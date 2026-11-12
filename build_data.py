"""يقرأ ملف الإكسيل ويطلع بيانات المخزن (أصناف + إضافات + صرف) كقاموس واحد."""
import datetime as dt, hashlib, os, re, warnings, zipfile
import openpyxl
warnings.filterwarnings("ignore")

def _s(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).strip()

def _n(v):
    try: return float(v or 0)
    except (TypeError, ValueError): return 0.0

def _nn(v):
    return None if v is None or v == "" else _n(v)   # null = الخانة فاضية (مفيش رصيد محفوظ)

def _d(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, (dt.datetime, dt.date)) else _s(v)

WH = {"الكبير", "الشامى", "المصنع"}

def _txt(v):
    """نص في خانة رقمية (زي: "أمانة عند فلان" مكتوب مكان أعداد المخازن) — الأرقام وكل ما يتقرا كرقم بيرجع فاضي."""
    s = _s(v)
    if not s: return ""
    try: float(s.replace(",", "")); return ""
    except ValueError: return s

def _wh(v):
    v = _s(v)
    if not v: return ""
    return v if v in WH else "غير محدد"   # أي نص حر في خانة المخزن بيتحسب غير محدد

def _merges(path):
    """الخلايا المدموجة (Merged) لكل شيت: {اسم الشيت: [(صف1, عمود1, صف2, عمود2), ...]} — بتتقرا من ملف الإكسيل مباشرة (سريع وبيشتغل مع القراءة السريعة)."""
    out = {}
    try:
        with zipfile.ZipFile(path) as z:
            wb = z.read("xl/workbook.xml").decode("utf-8", "ignore")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
            tgt = {}
            for m in re.finditer(r"<Relationship\b[^>]*>", rels):
                t = m.group(0); i = re.search(r'\bId="([^"]+)"', t); g = re.search(r'\bTarget="([^"]+)"', t)
                if i and g: tgt[i.group(1)] = g.group(1)
            for m in re.finditer(r"<sheet\b[^>]*>", wb):
                t = m.group(0); n = re.search(r'\bname="([^"]*)"', t); r = re.search(r'\br:id="([^"]+)"', t)
                if not (n and r and r.group(1) in tgt): continue
                f = tgt[r.group(1)].lstrip("/"); f = f if f.startswith("xl/") else "xl/" + f
                name = n.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")
                xml = z.read(f).decode("utf-8", "ignore")
                lst = []
                for mm in re.finditer(r'<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml):
                    lst.append((int(mm.group(2)), openpyxl.utils.column_index_from_string(mm.group(1)),
                                int(mm.group(4)), openpyxl.utils.column_index_from_string(mm.group(3))))
                out[name] = lst
    except Exception:
        pass   # لو حصل أي خطأ نكمل من غير دمج (زي الأول)
    return out

def read_workbook(path):
    try: return _read(path, True)        # قراءة سريعة
    except Exception: return _read(path, False)   # لو فشلت نرجع للطريقة العادية

def _rows(ws, lo, hi, ro):
    if ro: ws.reset_dimensions()
    return ws.iter_rows(min_row=lo, max_row=hi, min_col=1, max_col=15, values_only=True)

def _nzh(v):
    """تطبيع نص العناوين: من غير تشكيل/تطويل، وتوحيد الألف والياء والتاء المربوطة."""
    t = _s(v).lower()
    t = re.sub(r"[ً-ٟـ]", "", t)
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", t).strip()

def _rowsw(ws, lo, hi, ro, mc=40):
    if ro: ws.reset_dimensions()
    return ws.iter_rows(min_row=lo, max_row=hi, min_col=1, max_col=mc, values_only=True)

# أماكن أعمدة شيتات الإضافات/الصرف في ملفاتك (لو العناوين مش واضحة بنرجع لدول)
MV_DEFAULT = {"date": 0, "doc": 1, "wo": 2, "name": 4, "w": 6, "c": 7, "p": 8, "h": 9, "note": 10}

def _moves(ws, out, mg, ro, cnt, tag, cols=None):
    """بيقرا شيت إضافات أو صرف. cols = أماكن الأعمدة (الافتراضي = شكل ملفاتك)."""
    c = dict(MV_DEFAULT, **(cols or {}))
    # الأعمدة اللي ممكن تتدمج (Merged) ومكانها في الداتا: التاريخ 0، الإذن 1، أمر الشغل 2، وارد من/منصرف إلى 6، المخزن 7، ملاحظات 8
    TXT = {c["date"]: 0, c["doc"]: 1, c["wo"]: 2, c["p"]: 6, c["h"]: 7, c["note"]: 8}
    raw = {i: list(r) for i, r in enumerate(_rowsw(ws, 2, None, ro, max(c.values()) + 5), 2)}
    span = {}                                  # (صف, عمود الداتا) -> (أول صف, آخر صف) للدمج الرأسي
    for r1, c1, r2, c2 in mg.get(ws.title, []):
        if r1 < 2 or r1 not in raw: continue
        for sc, oc in TXT.items():
            if not (c1 - 1 <= sc <= c2 - 1): continue
            top = raw[r1][c1 - 1] if len(raw[r1]) >= c1 else None
            for rr in range(r1, min(r2, max(raw)) + 1):
                if rr in raw and len(raw[rr]) > sc: raw[rr][sc] = top      # كل الصفوف المدموجة ليها نفس القيمة
                if r2 > r1: span[(rr, oc)] = (r1, r2)
    g = lambda r, k: r[c[k]] if len(r) > c[k] else None
    rows = []
    for i, r in sorted(raw.items()):
        if not _s(g(r, "name")): continue
        mgd = {str(oc): list(span[(i, oc)]) for oc in TXT.values() if (i, oc) in span}
        rows.append([_d(g(r, "date")), _s(g(r, "doc")), _s(g(r, "wo")) if out else "", _s(g(r, "name")),
                     _n(g(r, "w")), _n(g(r, "c")), _s(g(r, "p")), _s(g(r, "h")) if cnt else _wh(g(r, "h")), _s(g(r, "note")), i]
                    + ([mgd or None, tag] if cnt else [mgd] if mgd else []))
    return rows

def _read(path, ro, pipes=False, coils=False, tag=None):
    """القوالب المعروفة: الشرائح (مخازن) / المواسير / اللفائف."""
    cnt = pipes or coils   # المواسير واللفائف: عدد بس (من غير مخازن)، والفرق في مكان الأعمدة
    tag = tag or ("c" if coils else "p")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=ro)
    bal = wb["الرصيد"]
    items = []
    st = 5 if coils else 6   # اللفائف: الهيدر 4 صفوف فالأصناف من صف 5
    for i, r in enumerate(_rows(bal, st, None, ro), st):
        if not _s(r[1]): continue
        if coils:   # اللفائف: العدد بس (الوزن صفر في الشيت)، العمود C = المصدر، العمود L = أرقام اللفات (نص) بنعرضه ملاحظات
            cn = _nn(r[10])
            items.append([_s(r[1]), _s(r[2]), 0.0, _n(r[4]), 0.0, 0.0, 0.0, _s(r[11]), i, None if cn is None else 0.0, cn, tag, _nn(r[6]), _nn(r[8])])
            continue
        if pipes:   # المواسير: عدد بس (من غير وزن ومن غير مخازن)، العمود H = ملاحظات
            cn = _nn(r[6])
            items.append([_s(r[1]), _s(r[2]), 0.0, _n(r[3]), 0.0, 0.0, 0.0, _s(r[7]), i, None if cn is None else 0.0, cn, tag, _nn(r[4]), _nn(r[5])])   # آخر رقمين: الإضافات والصرف زي ما هم مكتوبين في الشيت
        else:
            cust = " · ".join(dict.fromkeys(t for t in (_txt(r[11]), _txt(r[12]), _txt(r[13])) if t))   # نص مكتوب مكان أعداد المخازن (أمانة / عند عميل...)
            items.append([_s(r[1]), _s(r[2]), _n(r[3]), _n(r[4]), _n(r[11]), _n(r[12]), _n(r[13]), _s(r[14]), i, _nn(r[9]), _nn(r[10])] + (["", None, None, cust] if cust else []))
    mg = _merges(path)
    ins, outs = _moves(wb["الاضافات"], False, mg, ro, cnt, tag), _moves(wb["الصرف"], True, mg, ro, cnt, tag)
    dr = 3 if pipes else 1 if coils else 2
    last = _d(next(_rows(bal, dr, dr, ro))[2])   # تاريخ الجرد المكتوب في شيت الرصيد
    titles = [" ".join(_d(c) for c in r[:5] if c not in (None, "")) for r in _rows(bal, 1, 3, ro)]
    return {"titles": titles, "asOf": last, "items": items, "ins": ins, "outs": outs, "wt": False, "warnings": []}

# ---------------- القارئ العام: أي ملف من نفس العيلة (رصيد + إضافات + صرف) بيتفهم من عناوينه ----------------
WH_NAMES = {"الكبير", "الشامى", "الجديد", "المصنع", "الشامي"}

def _find_sheet(wb, *keys):
    for n in wb.sheetnames:
        t = _nzh(n)
        if any(k in t for k in keys): return wb[n]
    return None

def _profile(rows):
    """بيحلل صفوف الهيدر في شيت الرصيد ويرجع أماكن الأعمدة. بيرمي ValueError برسالة مفهومة لو مش فاهم."""
    hr = nc = None
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            t = _nzh(c)
            if "اسم" in t and "صنف" in t: hr, nc = i, j; break
        if hr is not None: break
    if hr is None: raise ValueError("مش لاقي عنوان 'اسم الصنف' في شيت الرصيد")
    head = rows[hr]; sub = rows[hr + 1] if hr + 1 < len(rows) else ()
    subl = lambda j: _nzh(sub[j]) if j < len(sub) else ""
    groups = []
    for j, c in enumerate(head):
        t = _nzh(c)
        if not t or j == nc: continue
        if "رصيد" in t and "اول" in t: groups.append(("open", j))
        elif "اضاف" in t or "وارد" in t: groups.append(("add", j))
        elif "صرف" in t or "منصرف" in t: groups.append(("out", j))
        elif "رصيد" in t and ("اخر" in t or "ختام" in t or "نهاي" in t): groups.append(("close", j))
    if not any(k in ("open", "close") for k, _ in groups): raise ValueError("مش لاقي عمود 'رصيد اول المدة' ولا 'رصيد اخر المدة'")
    starts = sorted(j for _, j in groups)
    cols = {}
    for key, j in groups:
        nxt = min([x for x in starts if x > j] or [j + 3])
        w = n = None
        for k in range(j, min(nxt, j + 3)):
            t = subl(k)
            if t == "وزن" and w is None: w = k
            elif t in ("عدد", "كميه") and n is None: n = k
        if w is None and n is None: n = j     # من غير عناوين فرعية = عمود عدد واحد
        cols[key] = (w, n)
    has_sub = any(subl(j) in ("وزن", "عدد") for j in range(len(sub)))
    warn = []
    if all(v[1] is None for v in cols.values()) and any(v[0] is not None for v in cols.values()):
        cols = {k: (None, w) for k, (w, n) in cols.items()}; warn.append("مفيش عمود 'عدد' فاعتبرت عمود 'وزن' هو العدد")
    store = None
    first = min(starts)
    for j, c in enumerate(head):
        t = _nzh(c)
        if j != nc and j < first and ("مخزن" in t or "مصدر" in t or "نوع" in t): store = j; break
    endc = max(c for v in cols.values() for c in v if c is not None)
    wh_cols = [j for j in range(endc + 1, len(head)) if _nzh(head[j]) == "المخزن" or _nzh(sub[j] if j < len(sub) else "") in {_nzh(x) for x in WH_NAMES} or _nzh(head[j]) in {_nzh(x) for x in WH_NAMES}]
    if wh_cols: warn.append("فيه أعمدة مخازن (الكبير/الشامى...) مش مدعومة في الملفات الجديدة، اتجاهلت")
    return {"hr": hr, "start": hr + (2 if has_sub else 1) + 1, "nc": nc, "store": store, "cols": cols, "endc": endc, "whc": set(wh_cols), "warn": warn, "head": head}

def _mv_cols(ws, ro):
    """أماكن أعمدة الإضافات/الصرف من العناوين (لو عنوان ناقص بنستخدم المكان المعتاد)."""
    hdr = None
    for r in _rowsw(ws, 1, 4, ro, 25):
        if any("تاريخ" in _nzh(c) for c in r): hdr = r; break
    if hdr is None: raise ValueError("مش لاقي صف العناوين (التاريخ، الصنف، عدد...) في شيت " + ws.title)
    want = {"date": ("تاريخ",), "doc": ("اذن",), "wo": ("شغل",), "name": ("صنف",), "w": ("وزن",), "c": ("عدد", "كميه"),
            "p": ("وارد", "منصرف", "الي", "جهه"), "h": ("القيمه", "مخزن"), "note": ("ملاحظ",)}
    out, used = {}, set()
    for k, keys in want.items():
        for j, c in enumerate(hdr):
            t = _nzh(c)
            if j not in used and t and any(x in t for x in keys) and not (k == "p" and "حركه" in t):
                out[k] = j; used.add(j); break
    if "name" not in out: raise ValueError("مش لاقي عمود 'الصنف' في شيت " + ws.title)
    return out

def _read_generic(path, ro, tag):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=ro)
    bal = _find_sheet(wb, "رصيد")
    ins_ws, outs_ws = _find_sheet(wb, "اضاف"), _find_sheet(wb, "صرف")
    if bal is None: raise ValueError("مفيش شيت اسمه 'الرصيد'")
    if ins_ws is None or outs_ws is None: raise ValueError("مفيش شيت 'الاضافات' أو شيت 'الصرف'")
    hdr_rows = list(_rowsw(bal, 1, 16, ro))
    pf = _profile(hdr_rows)
    cols = pf["cols"]
    g = lambda r, j: r[j] if j is not None and j < len(r) else None
    note_c = None
    for j in range(pf["endc"] + 1, len(pf["head"])):
        if j in pf["whc"]: continue
        if "ملاحظ" in _nzh(pf["head"][j]): note_c = j; break
    body = list(_rowsw(bal, pf["start"], None, ro))
    if note_c is None:
        for j in range(pf["endc"] + 1, 40):
            if j in pf["whc"]: continue
            if any(isinstance(r[j] if j < len(r) else None, str) and _txt(r[j] if j < len(r) else None) for r in body[:80]): note_c = j; break
    items, wsum = [], 0.0
    for i, r in enumerate(body, pf["start"]):
        if not _s(g(r, pf["nc"])): continue
        ow, on = _n(g(r, cols.get("open", (None, None))[0])), _n(g(r, cols.get("open", (None, None))[1]))
        cw, cn = _nn(g(r, cols.get("close", (None, None))[0])), _nn(g(r, cols.get("close", (None, None))[1]))
        se = _nn(g(r, cols["add"][1])) if "add" in cols else None
        sx = _nn(g(r, cols["out"][1])) if "out" in cols else None
        wsum += abs(ow) + abs(cw or 0) + sum(abs(_n(g(r, cols[k][0]))) for k in ("add", "out") if k in cols)
        items.append([_s(g(r, pf["nc"])), _s(g(r, pf["store"])) if pf["store"] is not None else "", ow, on, 0.0, 0.0, 0.0,
                      _s(g(r, note_c)) if note_c is not None else "", i, cw, cn, tag, se, sx])
    mg = _merges(path)
    ins = _moves(ins_ws, False, mg, ro, True, tag, _mv_cols(ins_ws, ro))
    outs = _moves(outs_ws, True, mg, ro, True, tag, _mv_cols(outs_ws, ro))
    wt = wsum > 0 or any(m[4] for m in ins + outs)
    if not wt:   # الوزن كله صفر: بنتعامل مع القسم كأنه عدد بس (زي المواسير واللفائف)
        for a in items: a[2] = 0.0; a[9] = None if a[10] is None else 0.0
        for m in ins + outs: m[4] = 0.0
    date = ""
    for r in hdr_rows[:6]:
        for c in r:
            if isinstance(c, (dt.datetime, dt.date)): date = _d(c); break
        if date: break
    titles = [" ".join(_d(c) for c in r[:5] if c not in (None, "")) for r in hdr_rows[:3]]
    return {"titles": titles, "asOf": date, "items": items, "ins": ins, "outs": outs, "wt": wt, "warnings": list(pf["warn"])}


def _kind(path):
    """نوع الملف: slices/pipes/coils = القوالب المعروفة، generic = أي ملف تاني بيتفهم من عناوينه، None = مش ملف مخزن."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if "الرصيد" not in wb.sheetnames:
            return "generic" if _find_sheet(wb, "رصيد") is not None else None
        ws = wb["الرصيد"]; ws.reset_dimensions()
        rows = list(ws.iter_rows(min_row=3, max_row=5, min_col=1, max_col=12, values_only=True))
        r3, r4, r5 = ([list(x) + [None] * 12 for x in rows] + [[None] * 12] * 3)[:3]
        if _s(r5[1]).startswith("لفة") and _nzh(r3[5]) == "الاضافات" and _nzh(r3[3]).startswith("رصيد اول"): return "coils"
        if _nzh(r4[4]) == "الاضافات" and _nzh(r4[3]).startswith("رصيد اول"): return "pipes"   # المواسير: الاضافات في العمود E (من غير وزن)
        if _nzh(r4[3]).startswith("رصيد اول") and _nzh(r4[5]) == "الاضافات" and _nzh(r4[7]) == "الصرف" and _nzh(r4[11]) == "المخزن": return "slices"
        return "generic"
    finally:
        wb.close()

def _catname(fname):
    t = os.path.splitext(fname)[0].replace("_", " ")
    t = re.sub(r"\b20\d\d\b", " ", t)
    return re.sub(r"\s+", " ", t).strip() or fname

def read_all(paths):
    """بيقرا كل ملفات الإكسيل ويدمجهم في بيانات واحدة. كل ملف (غير الشرائح) بيبقى قسم لوحده.
    الملفات اللي ما اتفهمتش بتتسجل في problems بدل ما توقف الباقي."""
    parts, problems, seen_slices = [], [], False
    for p in sorted(paths, key=lambda x: -os.path.getmtime(x)):        # الأحدث الأول (عشان لو فيه نسختين من الشرائح ناخد الأحدث)
        name = os.path.basename(p)
        try: k = _kind(p)
        except Exception as e:
            problems.append({"file": name, "reason": "الملف مش بيتفتح: " + str(e)[:100]}); continue
        if k is None:
            problems.append({"file": name, "reason": "مفيش فيه شيت اسمه 'الرصيد'، فمش شبه ملفات المخزن"}); continue
        if k == "slices":
            if seen_slices: problems.append({"file": name, "reason": "نسخة تانية من الشرائح، اتجاهلت (بنستخدم الأحدث)"}); continue
            seen_slices = True
        tag = None if k == "slices" else "x" + hashlib.md5(name.encode("utf-8")).hexdigest()[:6]
        try:
            fn = (lambda ro: _read(p, ro)) if k == "slices" else (lambda ro: _read(p, ro, k == "pipes", k == "coils", tag)) if k != "generic" else (lambda ro: _read_generic(p, ro, tag))
            try: d = fn(True)
            except ValueError: raise
            except Exception: d = fn(False)
        except ValueError as e:
            problems.append({"file": name, "reason": str(e)}); continue
        except Exception as e:
            problems.append({"file": name, "reason": "حصل خطأ وانا بقرا الملف: " + str(e)[:100]}); continue
        for w in d["warnings"]: problems.append({"file": name, "reason": w, "level": "info"})
        parts.append((name, k, tag, d))
    if not parts: raise ValueError("مفيش ملف إكسيل مفهوم في الفولدر")
    parts.sort(key=lambda x: x[0])
    sl = next((d for n, k, t, d in parts if k == "slices"), None)
    dates = list(dict.fromkeys(d["asOf"] for _, _, _, d in parts if d["asOf"]))
    return {"titles": sl["titles"] if sl else [],
            "cats": [{"id": t, "n": _catname(n), "wt": d["wt"], "titles": d["titles"], "file": n} for n, k, t, d in parts if t],
            "asOf": dates[0] if len(dates) == 1 else " · ".join(dates),
            "items": sum((d["items"] for *_, d in parts), []), "ins": sum((d["ins"] for *_, d in parts), []), "outs": sum((d["outs"] for *_, d in parts), []),
            "problems": problems,
            "report": [{"file": n, "kind": k, "cat": _catname(n) if t else "شرائح", "items": len(d["items"]), "ins": len(d["ins"]), "outs": len(d["outs"]), "wt": d["wt"]} for n, k, t, d in parts]}

if __name__ == "__main__":
    import sys, json
    d = read_all(sys.argv[1:])
    for r in d["report"]: print(r)
    for p in d["problems"]: print("!", p)
