#!/usr/bin/env python3
"""Refresh the MFSB / POF 2569 figures on the co-rev-exp dashboard.

Reads the RECORD sheet of each foundation's Revenue-Expense workbook,
regenerates the single AUTO-DATA block in index.html, and (only if something
changed) commits + pushes so GitHub Pages republishes.

Safe to run any time: idempotent, no commit when nothing changed, and it
refuses to publish data that is older than what is already live.

Sources
  drive  Google Drive API via a service account (used by the GitHub Action).
         Needs env GOOGLE_SERVICE_ACCOUNT_JSON and the two workbooks shared
         with that service account (Viewer).
  local  The synced Google Drive Desktop copy (G:\\My Drive\\...), for manual runs.

  python tools/refresh.py --ci                 GitHub Action: drive source, commit, push
  python tools/refresh.py --source local       manual run on the PC
  python tools/refresh.py --source local --dry-run    rewrite index.html only
"""
import argparse
import datetime as dt
import io
import json
import os
import re
import subprocess
import sys
import traceback
import warnings

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
YEAR = 2026  # CE year of the workbooks; the dashboard's 2569 section is tied to it

DRIVE_NAME = {
    "mfsb": "MFSB-Revenue-Expense-%d.xlsx" % YEAR,
    "pof": "POF-Revenue-Expense-%d.xlsx" % YEAR,
}
LOCAL_BASE = r"G:\My Drive\Documents\+MFSB\+ALL Co_Rev-Exp"
LOCAL_PATH = {
    "mfsb": os.path.join(LOCAL_BASE, "มูลนิธิธรรมทานกุศลจิต", "Account", DRIVE_NAME["mfsb"]),
    "pof": os.path.join(LOCAL_BASE, "มูลนิธิปิลันธน์โอวาท", "Account", DRIVE_NAME["pof"]),
}

BLOCK_RE = re.compile(r"/\*AUTO-DATA-START\*/.*?/\*AUTO-DATA-END\*/", re.S)


def log(msg):
    print("%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg), flush=True)


# ───────────────────────── sources ─────────────────────────
def open_drive(name):
    """Download a workbook by exact name via the service account; return BytesIO."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not set")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(raw), scopes=["https://www.googleapis.com/auth/drive.readonly"])
    svc = build("drive", "v3", credentials=creds, cache_discovery=False)
    safe = name.replace("'", "\\'")
    files = svc.files().list(q="name = '%s' and trashed = false" % safe,
                             fields="files(id,name,modifiedTime)", pageSize=10).execute().get("files", [])
    if not files:
        raise RuntimeError("'%s' not visible to the service account — share the workbook with it (Viewer)" % name)
    if len(files) > 1:
        raise RuntimeError("'%s' matched %d files — ambiguous, refusing to guess" % (name, len(files)))
    log("drive: %s modified %s" % (name, files[0]["modifiedTime"]))
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=files[0]["id"]))
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf


def open_local(path):
    if not os.path.exists(path):
        raise RuntimeError("source not found (is Google Drive mounted?): %s" % path)
    with open(path, "rb") as f:
        return io.BytesIO(f.read())


def read_entity(name, source):
    import openpyxl

    buf = open_drive(DRIVE_NAME[name]) if source == "drive" else open_local(LOCAL_PATH[name])
    ws = openpyxl.load_workbook(buf, data_only=True)["RECORD"]

    opening = None
    rev = [0.0] * 12
    exp = [0.0] * 12
    ins, outs = [], []
    last = None
    warns = 0
    for _seq, d, item, typ, amt, cat, note in ws.iter_rows(min_row=4, min_col=2, max_col=8, values_only=True):
        if amt is None:
            continue
        if not isinstance(amt, (int, float)):
            warns += 1
            continue
        if cat and "ยกมา" in str(cat):  # carry-forward balance, not a flow
            if opening is None:
                opening = float(amt)
            continue
        if not isinstance(d, dt.datetime) or d.year != YEAR:
            warns += 1
            continue
        m = d.month - 1
        label = " ".join(str(note or item or "").split())
        if len(label) > 34:
            label = label[:34] + "…"
        if typ == "รายรับ":
            rev[m] += amt
            if amt > 0:
                ins.append((d, amt, label))
        elif typ == "รายจ่าย":
            exp[m] += -amt
            if -amt > 0:
                outs.append((d, -amt, label))
        else:
            warns += 1
            continue
        last = d if last is None or d > last else last
    if opening is None:
        raise RuntimeError("%s: carry-forward (ยอดยกมา) row not found" % name)
    if last is None:
        raise RuntimeError("%s: no %d transactions found" % (name, YEAR))
    if warns:
        log("WARN %s: %d row(s) skipped (bad date/amount/type) — check the RECORD sheet" % (name, warns))

    n = last.month
    rev, exp = rev[:n], exp[:n]
    if min(rev) < 0 or min(exp) < 0:
        raise RuntimeError("%s: negative monthly revenue/expense — refusing to publish" % name)
    nraw = sum(rev) - sum(exp)
    top = lambda rows: [[d.strftime("%Y-%m-%d"), round(a), l]
                        for d, a, l in sorted(rows, key=lambda r: (-r[1], r[0]))[:3]]
    return {
        "asOf": last.strftime("%Y-%m-%d"),
        "opening": round(opening),
        "rev": [round(x) for x in rev],
        "exp": [round(x) for x in exp],
        "net": [round(r - e) for r, e in zip(rev, exp)],
        "R": round(sum(rev)),
        "E": round(sum(exp)),
        "N": round(nraw),
        "close": round(opening + nraw),
        "top": {"in": top(ins), "out": top(outs)},
    }


# ───────────────────────── page + git ─────────────────────────
def build_block(data):
    as_of = {k: v.pop("asOf") for k, v in data.items()}
    dump = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    body = "{\n \"asOf\":%s,\n \"mfsb\":%s,\n \"pof\":%s\n}" % (dump(as_of), dump(data["mfsb"]), dump(data["pof"]))
    return "/*AUTO-DATA-START*/\nconst AUTO=%s;\n/*AUTO-DATA-END*/" % body, as_of


def parse_existing(block):
    m = re.search(r"const AUTO=(\{.*\});", block, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except ValueError:
        return None


def git(repo, *args, check=True):
    r = subprocess.run(["git", "-C", repo, *args], capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (r.stdout + r.stderr).strip()
    if out:
        log("git %s -> %s" % (" ".join(args[:2]), out.replace("\n", " | ")[:300]))
    if check and r.returncode != 0:
        raise RuntimeError("git %s failed (exit %d)" % (" ".join(args), r.returncode))
    return r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ci", action="store_true", help="GitHub Action mode (drive source, bot identity, rebase before push)")
    ap.add_argument("--source", choices=("drive", "local"))
    ap.add_argument("--repo", default=os.path.abspath(os.path.join(HERE, "..")))
    ap.add_argument("--dry-run", action="store_true", help="rewrite index.html but do not commit/push")
    args = ap.parse_args()
    source = args.source or ("drive" if args.ci or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") else "local")
    repo = args.repo

    data = {k: read_entity(k, source) for k in DRIVE_NAME}
    new_block, as_of = build_block({k: dict(v) for k, v in data.items()})

    page = os.path.join(repo, "index.html")
    with open(page, encoding="utf-8", newline="") as f:
        html = f.read()
    m = BLOCK_RE.search(html)
    if not m:
        raise RuntimeError("AUTO-DATA markers not found in index.html")
    old_block = m.group(0)

    if new_block == old_block:
        log("no change (MFSB through %s, POF through %s)" % (as_of["mfsb"], as_of["pof"]))
        return 0

    prev = parse_existing(old_block)
    if prev:
        for k in ("mfsb", "pof"):
            if as_of[k] < prev["asOf"][k]:
                raise RuntimeError("%s source is older than what is live (%s < %s) — refusing to publish a stale copy"
                                   % (k, as_of[k], prev["asOf"][k]))
            if len(data[k]["rev"]) < len(prev[k]["rev"]):
                raise RuntimeError("%s now has fewer months than the live page — refusing" % k)
            if abs(data[k]["opening"] - prev[k]["opening"]) > 1:
                log("WARN %s opening balance changed %s -> %s" % (k, prev[k]["opening"], data[k]["opening"]))

    html = html[: m.start()] + new_block + html[m.end():]
    with open(page, "w", encoding="utf-8", newline="") as f:
        f.write(html)
    log("index.html updated (MFSB through %s, POF through %s)" % (as_of["mfsb"], as_of["pof"]))

    if args.dry_run:
        log("dry-run: not committing")
        return 0

    if args.ci:
        git(repo, "config", "user.name", "github-actions[bot]")
        git(repo, "config", "user.email", "41898282+github-actions[bot]@users.noreply.github.com")
    git(repo, "add", "index.html")
    msg = "Auto-refresh %d data: MFSB through %s, POF through %s" % (YEAR, as_of["mfsb"], as_of["pof"])
    git(repo, "commit", "-m", msg)
    if args.ci:
        git(repo, "pull", "--rebase", "origin", "main")
    git(repo, "push", "origin", "HEAD:main")
    log("pushed: " + msg)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        log("FAILED: %s" % e)
        log(traceback.format_exc().strip().replace("\n", " | ")[-600:])
        sys.exit(1)
