"""بيشتغل على GitHub: ينزّل ملفات الإكسيل من فولدر Google Drive، يبني بيانات المخزن، ويجهّز فولدر site/ للنشر على Cloudflare Pages.
لو مفيش تغيير في الأرقام مبيعملش نشر (عشان مانستهلكش حد النشر الشهري)."""
import datetime as dt, hashlib, json, os, re, shutil, sys, tempfile, urllib.request
import requests
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession
from build_data import read_all

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "site")
FOLDER = os.environ["DRIVE_FOLDER_ID"].strip()
PROJECT = os.environ.get("CF_PROJECT", "").strip()
SITE_URL = (os.environ.get("CF_URL", "").strip().rstrip("/") or (f"https://{PROJECT}.pages.dev" if PROJECT else "")).replace("https://https://", "https://")
if SITE_URL and not SITE_URL.startswith("http"): SITE_URL = "https://" + SITE_URL
FORCE = os.environ.get("FORCE", "") == "1"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET = "application/vnd.google-apps.spreadsheet"

def out(k, v):   # بيبلّغ الـ workflow
    p = os.environ.get("GITHUB_OUTPUT")
    if p: open(p, "a").write(f"{k}={v}\n")
    print(f"[{k}] {v}")

def drive():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    cr = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return AuthorizedSession(cr)

MODS = []   # آخر تعديل فعلي لكل ملف على الدرايف

def download(s, dest):
    """بينزّل كل xlsx (وأي Google Sheet كـ xlsx) من الفولدر ومن الفولدرات اللي جواه."""
    files, stack = [], [FOLDER]
    while stack:
        fid = stack.pop(); tok = None; dated = []
        while True:
            r = s.get("https://www.googleapis.com/drive/v3/files", params={
                "q": f"'{fid}' in parents and trashed=false", "pageSize": 200, "pageToken": tok,
                "fields": "nextPageToken,files(id,name,mimeType,modifiedTime,size)",
                "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}, timeout=60)
            r.raise_for_status(); j = r.json()
            for f in j.get("files", []):
                if f["mimeType"] == "application/vnd.google-apps.folder":
                    if "الرصيد في" in f["name"].replace("ى", "ي"): dated.append(f)      # فولدرات اليوم: ناخد الأحدث بس
                    else: stack.append(f["id"])
                elif f["name"].startswith("~$"): continue
                elif f["mimeType"] == GSHEET or f["name"].lower().endswith(".xlsx"): files.append(f)
            tok = j.get("nextPageToken")
            if not tok: break
        if dated:
            def key(f):
                n = f["name"]; m = re.search(r"(\d{1,2})[-_/.](\d{1,2})[-_/.](\d{4})", n)      # يوم-شهر-سنة
                if m: d = (int(m[3]), int(m[2]), int(m[1]))
                else:
                    m = re.search(r"(\d{4})[-_/.](\d{1,2})[-_/.](\d{1,2})", n); d = (int(m[1]), int(m[2]), int(m[3])) if m else (0, 0, 0)
                return (d, f["modifiedTime"])
            best = max(dated, key=key); print("  using dated folder:", best["name"]); stack.append(best["id"])
    got = []
    for f in sorted(files, key=lambda f: f["name"]):
        name = f["name"] if f["name"].lower().endswith(".xlsx") else f["name"] + ".xlsx"
        name = name.replace("/", "_")
        if f["mimeType"] == GSHEET:
            r = s.get(f"https://www.googleapis.com/drive/v3/files/{f['id']}/export", params={"mimeType": XLSX}, timeout=120)
        else:
            r = s.get(f"https://www.googleapis.com/drive/v3/files/{f['id']}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=120)
        r.raise_for_status()
        p = os.path.join(dest, name)
        if os.path.exists(p): p = os.path.join(dest, f["id"][:6] + "_" + name)
        open(p, "wb").write(r.content)
        t = dt.datetime.fromisoformat(f["modifiedTime"].replace("Z", "+00:00")).timestamp()
        MODS.append(t)
        os.utime(p, (t, t))        # الأحدث الأول (بيفرق لو فيه نسختين من الشرائح)
        got.append(p); print("  downloaded:", name, len(r.content) // 1024, "KB")
    return got

def live_version():   # النسخة المنشورة حاليًا (بدون حفظ أي حالة على GitHub)
    if not SITE_URL: return {}
    try:
        with urllib.request.urlopen(f"{SITE_URL}/version.json?t={int(dt.datetime.now().timestamp())}", timeout=30) as r:
            return json.load(r)
    except Exception as e:
        print("no live version yet:", e); return {}

def main():
    tmp = tempfile.mkdtemp()
    paths = download(drive(), tmp)
    if not paths:
        out("changed", "false"); print("!! مفيش ملفات إكسيل في الفولدر (أو الفولدر مش متشير لحساب الخدمة). مش هننشر حاجة."); sys.exit(1)
    d = read_all(paths)
    for r in d["report"]: print("  ", r)
    for p in d["problems"]: print("   !", p)
    if MODS:   # آخر تعديل فعلي لأي ملف (بتوقيت القاهرة)، مش تاريخ الجرد المكتوب جوه الشيت
        from zoneinfo import ZoneInfo
        d["mod"] = dt.datetime.fromtimestamp(max(MODS), ZoneInfo("Africa/Cairo")).isoformat(timespec="seconds")
    sig = hashlib.sha256(json.dumps({k: d.get(k) for k in ("items", "ins", "outs", "asOf", "cats", "problems", "mod")}, ensure_ascii=False).encode()).hexdigest()
    live = live_version()
    if live.get("s") == sig[:8] and not FORCE:
        out("changed", "false"); print("مفيش تغيير، مفيش نشر."); return
    d.update(v=int(live.get("v", 0)) + 1, sig=sig, built=dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds"))
    shutil.rmtree(SITE, ignore_errors=True); shutil.copytree(os.path.join(HERE, "web"), SITE)
    json.dump(d, open(os.path.join(SITE, "data.json"), "w", encoding="utf-8"), ensure_ascii=False)
    json.dump({"v": d["v"], "s": sig[:8]}, open(os.path.join(SITE, "version.json"), "w"))
    os.makedirs(os.path.join(SITE, "xlsx"))
    fl = []
    for p in paths:
        shutil.copy2(p, os.path.join(SITE, "xlsx", os.path.basename(p)))
        st = os.stat(p); fl.append({"name": os.path.basename(p), "size": st.st_size, "mtime": int(st.st_mtime)})
    json.dump(fl, open(os.path.join(SITE, "files.json"), "w", encoding="utf-8"), ensure_ascii=False)
    open(os.path.join(SITE, "_headers"), "w").write("/*.json\n  Cache-Control: no-store\n/xlsx/*\n  Cache-Control: no-store\n/index.html\n  Cache-Control: no-cache\n/\n  Cache-Control: no-cache\n")
    out("changed", "true"); out("version", d["v"])

if __name__ == "__main__": main()
