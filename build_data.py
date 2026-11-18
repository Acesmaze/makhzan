"""يقرأ ملف الإكسيل ويطلع بيانات المخزن (أصناف + إضافات + صرف) كقاموس واحد."""
import copy, datetime as dt, hashlib, json, os, re, tempfile, warnings, zipfile
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

_ERR = re.compile(r"^#(DIV/0!|N/A|VALUE!|REF!|NAME\?|NUM!|NULL!)$", re.I)
def _clean(v):
    """نص الخلية، بس أخطاء الإكسيل (#DIV/0! وأخواتها) بتتحول لفاضي."""
    t = _s(v)
    return "" if _ERR.match(t) else t

# نص زي "986/19900+1044/17340": رقم اللفة/وزنها. الوزن = مجموع الأرقام اللي بعد "/" (من غير اللي بين أقواس زي "(مقص حجاج 949/4180)")
_COIL = re.compile(r"(?<!\d/)(?<!\d)(\d{1,6})[^\d/+()\s]{0,3}\s*/\s*(\d{3,6})(?!\d|/\d)")
_PLAIN = re.compile(r"^\s*\d{3,6}(?:\s*\+\s*\d{3,6})*\s*$")
def _entries(t):
    ws = [int(b) for a, b in _COIL.findall(t) if not (int(a) <= 12 and 2020 <= int(b) <= 2035)]   # تاريخ شهر/سنة مش لفة
    if not ws and _PLAIN.match(t): ws = [int(x) for x in re.findall(r"\d+", t)]       # لفات صغيرة مكتوبة أوزان بس: 972+1828
    return ws

_PLAINW = re.compile(r"وزن\s+(?!ال[اإ]ذن|اذن)[^\d/()]{0,20}?(\d{3,6})(?![\d/])")
def _coil_w(text, cn=None):
    """(مجموع الوزن، عدد اللفات). القوس اللي فيه "مقص" بيتشال (اللفة دي مش في العدد)، والقوس التاني زي (9س/4990) لفة عادية.
    لو cn (العدد في الرصيد) معروف بنختار القراءة اللي عدد لفاتها = العدد؛ ولو مفيش واحدة بتطابق بنرجع (0, 0)."""
    t_b = re.sub(r"\([^)]*مقص[^)]*\)", " ", text); t_b = t_b.replace("(", " ").replace(")", " ")
    t_a = re.sub(r"\([^)]*\)", " ", text)
    if cn is None:
        ws = _entries(t_b)
        if not ws:
            m = _PLAINW.search(text)                       # وزن الجلاد 30450: إجمالي وزن الحركة
            if m: ws = [int(m.group(1))]
        return (float(sum(ws)), len(ws))
    for t in (t_b, t_a):
        ws = _entries(t)
        if ws and len(ws) == round(cn): return (float(sum(ws)), len(ws))
    return (0.0, 0)

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

def _hidden_rows(path):
    """الصفوف المخفية (فلتر أو Hide) لكل شيت: {اسم الشيت: {أرقام الصفوف}} — من ملف الإكسيل مباشرة."""
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
                xml = z.read(f).decode("utf-8", "ignore"); hs = set()
                for mm in re.finditer(r"<row\b[^>]*>", xml):
                    tt = mm.group(0)
                    if re.search(r'\bhidden="(1|true)"', tt):
                        rn = re.search(r'\br="(\d+)"', tt)
                        if rn: hs.add(int(rn.group(1)))
                out[n.group(1).replace("&amp;", "&")] = hs
    except Exception:
        pass
    return out

def _mark_hidden(path, d):
    """بيعلّم الأصناف اللي صفها مخفي في شيت الرصيد بالإكسيل الأصلي (آخر خانة في الصنف = 1)."""
    h = _hidden_rows(path)
    name = next((n for n in h if "رصيد" in n), None)
    hs = h.get(name) if name else None
    if not hs: return
    if all(a[8] in hs for a in d["items"]): return          # لو كل الأصناف مخفية (فلتر غلط أو ملف متصدّر) الإخفاء ملوش معنى، فنتجاهله
    for a in d["items"]:
        a += [None] * (15 - len(a))
        a.append(1 if a[8] in hs else None)

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

def _read(path, ro, pipes=False, coils=False, tag=None):
    cnt = pipes or coils   # المواسير واللفائف: عدد بس (من غير مخازن)، والفرق في مكان الأعمدة وعلامة النوع
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
    # أعمدة الشيت (من 0) اللي ممكن تتدمج: التاريخ، الإذن، أمر الشغل، وارد من/منصرف إلى، المخزن، ملاحظات
    # ومكانها في الداتا:   0        1       2              6               7        8
    TXT = {0: 0, 1: 1, 2: 2, 8: 6, 9: 7, 10: 8}
    def moves(ws, out):
        raw = {i: list(r) for i, r in enumerate(_rows(ws, 2, None, ro), 2)}
        span = {}                                  # (صف, عمود الداتا) -> (أول صف, آخر صف) للدمج الرأسي
        for r1, c1, r2, c2 in mg.get(ws.title, []):
            if r1 < 2 or r1 not in raw: continue
            for sc, oc in TXT.items():
                if not (c1 - 1 <= sc <= c2 - 1): continue
                top = raw[r1][c1 - 1] if len(raw[r1]) >= c1 else None
                for rr in range(r1, min(r2, max(raw)) + 1):
                    if rr in raw and len(raw[rr]) > sc: raw[rr][sc] = top      # كل الصفوف المدموجة ليها نفس القيمة
                    if r2 > r1: span[(rr, oc)] = (r1, r2)
        rows = []
        for i, r in sorted(raw.items()):
            if not _s(r[4]): continue
            mgd = {str(oc): list(span[(i, oc)]) for oc in TXT.values() if (i, oc) in span}
            rows.append([_d(r[0]), _s(r[1]), _s(r[2]) if out else "", _s(r[4]),
                         _n(r[6]), _n(r[7]), _s(r[8]), _s(r[9]) if cnt else _wh(r[9]), _s(r[10]), i] + ([mgd or None, tag] if cnt else [mgd] if mgd else []))
        return rows
    ins, outs = moves(wb["الاضافات"], False), moves(wb["الصرف"], True)
    dr = 3 if pipes else 1 if coils else 2
    last = _d(next(_rows(bal, dr, dr, ro))[2])   # تاريخ الجرد المكتوب في شيت الرصيد
    titles = [" ".join(_d(c) for c in r[:5] if c not in (None, "")) for r in _rows(bal, 1, 3, ro)]
    return {"titles": titles, "asOf": last, "items": items, "ins": ins, "outs": outs}



# ---------------------------------------------------------------------------------------------
# القارئ الذكي: أي ملف إكسيل جديد (شيت رصيد + إضافات + صرف) بيتفهم من عناوين الأعمدة نفسها
# ---------------------------------------------------------------------------------------------
def _nzh(v):
    """تطبيع نص العناوين: من غير تشكيل/تطويل، وتوحيد الألف والياء والتاء المربوطة."""
    t = _s(v).lower()
    t = re.sub(r"[\u064b-\u065f\u0640]", "", t)
    t = t.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", t).strip()

def _grid(ws, lo, hi, ro, mc=40):
    if ro: ws.reset_dimensions()
    return [list(r) for r in ws.iter_rows(min_row=lo, max_row=hi, min_col=1, max_col=mc, values_only=True)]

def _find_sheet(wb, *keys):
    for n in wb.sheetnames:
        z = _nzh(n)
        if any(k in z for k in keys): return wb[n]
    return None

MV_DEFAULT = {"date": 0, "doc": 1, "wo": 2, "name": 4, "w": 6, "c": 7, "p": 8, "h": 9, "note": 10}

def _mv_cols(head):
    """أماكن أعمدة شيت الإضافات/الصرف من العناوين (اللي ما اتلقاش بيرجع للمكان الافتراضي)."""
    c = {}
    for j, v in enumerate(head):
        t = _nzh(v)
        if not t: continue
        if "تاريخ" in t: c.setdefault("date", j)
        elif "اذن" in t: c.setdefault("doc", j)
        elif "امر" in t and "شغل" in t: c.setdefault("wo", j)
        elif "صنف" in t: c.setdefault("name", j)
        elif t == "وزن" or t.startswith("وزن "): c.setdefault("w", j)
        elif t == "عدد" or t.startswith("عدد "): c.setdefault("c", j)
        elif "وارد من" in t or "منصرف" in t or t.startswith("من") or "الي" in t.split(): c.setdefault("p", j)
        elif "طبليه" in t or "بسكول" in t or "مخزن" in t: c.setdefault("h", j)
        elif "ملاحظ" in t: c.setdefault("note", j)
    return dict(MV_DEFAULT, **c)

def _moves(ws, out, mg, ro, tag):
    g = _grid(ws, 1, None, ro)
    hr = next((i for i, r in enumerate(g[:6]) if any("صنف" in _nzh(x) for x in r)), 0)
    c = _mv_cols(g[hr] if g else [])
    TXT = {c["date"]: 0, c["doc"]: 1, c["wo"]: 2, c["p"]: 6, c["h"]: 7, c["note"]: 8}
    width = max(c.values()) + 1
    raw = {i + 1: (r + [None] * width)[:max(width, len(r))] for i, r in enumerate(g) if i > hr}
    span = {}
    for r1, c1, r2, c2 in mg.get(ws.title, []):
        if r1 not in raw: continue
        for sc, oc in TXT.items():
            if not (c1 - 1 <= sc <= c2 - 1): continue
            top = raw[r1][c1 - 1] if len(raw[r1]) >= c1 else None
            for rr in range(r1, min(r2, max(raw)) + 1):
                if rr in raw and len(raw[rr]) > sc: raw[rr][sc] = top
                if r2 > r1: span[(rr, oc)] = (r1, r2)
    rows = []
    for i, r in sorted(raw.items()):
        if not _s(r[c["name"]]): continue
        mgd = {str(oc): list(span[(i, oc)]) for oc in TXT.values() if (i, oc) in span}
        rows.append([_d(r[c["date"]]), _s(r[c["doc"]]), _s(r[c["wo"]]) if out else "", _s(r[c["name"]]),
                     _nn(r[c["w"]]), _n(r[c["c"]]), _s(r[c["p"]]), _s(r[c["h"]]), _s(r[c["note"]]), i, mgd or None, tag])
    return rows

def _read_generic(path, ro, tag):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=ro)
    try:
        bal = _find_sheet(wb, "رصيد")
        if bal is None: raise ValueError("مفيش فيه شيت اسمه 'الرصيد'، فمش شبه ملفات المخزن")
        si = _find_sheet(wb, "اضاف", "وارد"); so = _find_sheet(wb, "صرف", "منصرف")
        if si is None or so is None: raise ValueError("ناقص شيت الإضافات أو الصرف")
        g = _grid(bal, 1, None, ro)
        hr = next((i for i, r in enumerate(g[:12]) if any("اسم الصنف" in _nzh(x) or _nzh(x) == "الصنف" for x in r)), None)
        if hr is None: raise ValueError("مش لاقي عنوان 'اسم الصنف' في شيت الرصيد")
        H = (g[hr] + [None] * 40)[:40]
        sub = (g[hr + 1] + [None] * 40)[:40] if hr + 1 < len(g) else [None] * 40
        has_sub = any(_nzh(x) in ("وزن", "عدد") for x in sub)
        K = {}
        for j, v in enumerate(H):
            t = _nzh(v)
            if not t: continue
            key = ("name" if "اسم الصنف" in t or t == "الصنف" else "type" if t in ("النوع", "المصدر", "من") or (t == "المخزن" and j <= 3) else
                   "open" if "رصيد" in t and "اول" in t else "close" if "رصيد" in t and "اخر" in t else
                   "in" if "اضاف" in t or t.startswith("وارد") else "out" if "صرف" in t or "منصرف" in t else
                   "note" if "ملاحظ" in t else None)
            if not key or key in K: continue
            if key in ("open", "in", "out", "close"):
                w = n = j
                if has_sub:
                    a, b = _nzh(sub[j]), _nzh(sub[j + 1]) if j + 1 < 40 else ""
                    if a == "وزن" and b == "عدد": w, n = j, j + 1
                    elif a == "وزن": w, n = j, None
                    else: w, n = None, j
                K[key] = (w, n)
            else: K[key] = j
        if "note" not in K and "close" in K:
            nj = max(x for x in K["close"] if x is not None) + 1   # لو مفيش عنوان ملاحظات: العمود اللي بعد رصيد آخر المدة، بس لو عنوانه فاضي (مش عمود تاني زي "الوزن الفعلى")
            if nj < len(H) and not _nzh(H[nj]): K["note"] = nj
        if "name" not in K: raise ValueError("مش لاقي عمود اسم الصنف")
        for need in ("open", "close"):
            if need not in K: raise ValueError("مش لاقي عمود 'رصيد %s المدة'" % ("أول" if need == "open" else "آخر"))
        cmax = max(x for x in K["close"] if x is not None)
        DET = [j for j, v in enumerate(H) if j > cmax and "مخزن" in _nzh(v)]      # أعمدة نص تفاصيل اللفائف (المخزن الكبير)
        st = hr + 1 + (1 if has_sub else 0)
        gv = lambda r, j: r[j] if j is not None and j < len(r) else None
        items, wt = [], False
        for i, r in enumerate(g[st:], st + 1):
            nm = _s(gv(r, K["name"]))
            if not nm or re.match(r"^(الاجمالي|اجمالي|المجموع|مجموع)", _nzh(nm)): continue
            ow, on = (gv(r, K["open"][0]), gv(r, K["open"][1])); cw, cn = (gv(r, K["close"][0]), gv(r, K["close"][1]))
            iw, inn = K.get("in", (None, None)); xw, xn = K.get("out", (None, None))
            sw_in, sw_out = gv(r, K["in"][1]) if "in" in K else None, gv(r, K["out"][1]) if "out" in K else None
            det = " + ".join(t for t in (_clean(gv(r, j)) for j in DET) if t)
            cwv = _nn(cw) if K["close"][0] is not None else (None if _nn(cn) is None else 0.0)
            if det and not cwv and (_nn(cn) or 0) > 0:            # وزن آخر المدة مش مكتوب في الرصيد: بنحسبه من تفاصيل اللفائف
                dw, dn = _coil_w(det, _nn(cn))
                if dw: cwv = dw          # بس لو عدد اللفات المكتوبة = العدد في الرصيد (عشان مانعرضش وزن ناقص)
            if _n(ow) or _n(cw) or (cwv or 0): wt = True
            items.append([nm, _clean(gv(r, K.get("type"))), (_nn(ow) if K["open"][0] is not None else None), _n(on), 0.0, 0.0, 0.0, _clean(gv(r, K.get("note"))) or det, i,
                          cwv, _nn(cn), tag, _nn(sw_in), _nn(sw_out)])
        mg = _merges(path)
        ins, outs = _moves(si, False, mg, ro, tag), _moves(so, True, mg, ro, tag)
        for m in ins + outs:                     # وزن الحركة مش مكتوب لكن مكتوب في الملاحظات (وزن كاما 798/8470): بناخده منها
            if not m[4] and (m[5] or 0) > 0:
                dw, dn = _coil_w(" ".join(str(x) for x in (m[7], m[8]) if x))
                if dw: m[4] = dw
        mvn = {}
        for m in ins: mvn.setdefault(m[3], []).append((1, m))
        for m in outs: mvn.setdefault(m[3], []).append((-1, m))
        for it in items:                         # لسه مفيش وزن لآخر المدة: بنحسبه من حركات الصنف (أول المدة + وارد - منصرف) لو كل حركاته ليها وزن
            if (it[10] or 0) > 0 and not it[9] and it[0] in mvn:
                mv = [(sg, m) for sg, m in mvn[it[0]] if (m[5] or 0) > 0]
                if mv and all(m[4] for _, m in mv):
                    w = (it[2] or 0) + sum(sg * m[4] for sg, m in mv)
                    if w > 0: it[9] = w
        if any(m[4] for m in ins + outs) or any(it[9] for it in items): wt = True
        titles = [" ".join(_d(c) for c in r[:5] if c not in (None, "")) for r in g[:max(hr, 3)][:3]]
        last = ""
        for r in g[:hr + 1]:
            for c in r[:15]:
                if isinstance(c, (dt.datetime, dt.date)): last = _d(c); break
                m = re.search(r"\d{4}-\d{2}-\d{2}", _s(c))
                if m: last = m.group(0); break
            if last: break
        if not last: last = max([m[0] for m in ins + outs if m[0]] or [""])
        if not items: raise ValueError("شيت الرصيد مفيهوش أصناف")
        return {"titles": titles, "asOf": last, "items": items, "ins": ins, "outs": outs, "wt": wt}
    finally:
        try: wb.close()
        except Exception: pass


def _kind(path):
    """بيفرّق بين الشرائح والمواسير واللفائف من شكل شيت الرصيد، وأي حاجة تانية = generic (None = مش ملف مخزن)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if "الرصيد" not in wb.sheetnames:
            return "generic" if _find_sheet(wb, "رصيد") is not None else None
        ws = wb["الرصيد"]; ws.reset_dimensions()
        rows = list(ws.iter_rows(min_row=3, max_row=5, min_col=1, max_col=12, values_only=True))
        r3, r4, r5 = ([list(x) + [None] * 12 for x in rows] + [[None] * 12] * 3)[:3]
        if _s(r5[1]).startswith("لفة") and _s(r5[2]) != "" or (_s(r5[1]).startswith("لفة") and "رصيد" in _nzh(r4[3])): return "coils"
        if _s(r4[4]) == "الاضافات" and _s(r4[1]).replace("ـ", "").replace(" ", "") == "اسمالصنف": return "pipes"
        if _s(r4[3]).startswith("رصيد") and _s(r4[5]) == "الاضافات" and any("الكبير" in _s(x) for x in r5): return "slices"
        return "generic"
    finally:
        wb.close()

def _catname(path):
    import os
    t = os.path.splitext(os.path.basename(path))[0].replace("_", " ")
    t = re.sub(r"\b20\d\d\b", "", t)
    return re.sub(r"\s+", " ", t).strip() or "قسم جديد"

def _catid(path):
    import hashlib, os
    return "x" + hashlib.md5(_nzh(_catname(path)).encode("utf-8")).hexdigest()[:6]

WJ = "\u2060"   # علامة غير مرئية بتفرّق الأصناف اللي ليها نفس الاسم في ملفين مختلفين (عشان حركات ملف متتحسبش على ملف تاني)

def _merge_rows(rows):
    """نفس الصنف مكتوب أكتر من مرة في نفس الملف: بنجمعه صف واحد. الإضافات والصرف في الشيت بتتحسب على الاسم
    (فكل صف فيهم بيعرض الإجمالي كله)، فبناخدها مرة واحدة، وبنجمع أرصدة أول المدة بس، ونطلع الرصيد = مجموع الأول + الإضافات - الصرف."""
    a = list(rows[0]); k = len(rows)
    for r in rows[1:]:
        for j in (2, 3, 4, 5, 6): a[j] = None if a[j] is None and r[j] is None else (a[j] or 0) + (r[j] or 0)
        for j in (1, 7):
            t = [x for x in dict.fromkeys([a[j], r[j]]) if x]
            a[j] = " | ".join(t)
    f = rows[0]
    for cj, oj in ((9, 2), (10, 3)):
        if all(r[cj] is not None for r in rows):
            delta = (f[cj] or 0) - (f[oj] or 0)          # الإضافات - الصرف (نفسها في كل صف)
            a[cj] = sum((r[cj] or 0) for r in rows) - (k - 1) * delta
    if len(a) > 15: a[15] = 1 if all(len(r) > 15 and r[15] for r in rows) else None     # مخفي بس لو كل الصفوف المكررة مخفية
    return a

def _dedupe(got, items, ins, outs):
    import collections
    # 1) تكرار الاسم جوه نفس الملف (الأقسام الجديدة بس؛ الشرائح زي ما هي)
    seen = collections.defaultdict(list)
    for a in items:
        if len(a) > 11 and a[11]: seen[(a[11], a[0])].append(a)
    drop = set()
    for k, v in seen.items():
        if len(v) > 1:
            m = _merge_rows(v); v[0][:] = m; drop.update(id(x) for x in v[1:])
    items[:] = [a for a in items if id(a) not in drop]
    # 2) نفس الاسم في ملفين مختلفين: كل ملف يبقى صنف لوحده
    owners = collections.defaultdict(list)
    def tagof(r): return r[11] if len(r) > 11 and r[11] else "sh"
    for rows, ni in ((items, 0), (ins, 3), (outs, 3)):   # الحركات كمان بتعدّ: صنف ليه حركات في ملف من غير ما يكون في شيت رصيده
        for a in rows:
            if tagof(a) not in owners[a[ni]]: owners[a[ni]].append(tagof(a))
    for rows, ni in ((items, 0), (ins, 3), (outs, 3)):
        for r in rows:
            o = owners.get(r[ni])
            if o and len(o) > 1 and tagof(r) in o: r[ni] = r[ni] + WJ * o.index(tagof(r))

# ---- كاش قراءة الملفات: كل ملف بيتقرا مرة واحدة، ولما حاجة تتحفظ بنقرا الملف اللي اتغيّر بس (بدل الـ 14 كلهم) ----
_CF = os.path.join(tempfile.gettempdir(), "mkhzn_parse_cache.json")
try: _CV = hashlib.md5(open(__file__, "rb").read()).hexdigest()   # أي تعديل في الكود بيلغي الكاش القديم
except Exception: _CV = "x"
_MEM = None

def _cache():
    global _MEM
    if _MEM is None:
        try:
            c = json.load(open(_CF, encoding="utf-8"))
            _MEM = c.get("e", {}) if c.get("v") == _CV else {}
        except Exception: _MEM = {}
    return _MEM

def _ckey(p):
    st = os.stat(p); return "%s|%d|%d" % (os.path.basename(p), int(st.st_mtime * 1000), st.st_size)

def _cache_save(keep):
    c = _cache()
    for k in [k for k in c if k not in keep]: del c[k]
    try:
        json.dump({"v": _CV, "e": c}, open(_CF + ".tmp", "w", encoding="utf-8"), ensure_ascii=False)
        os.replace(_CF + ".tmp", _CF)
    except Exception: pass

def read_all(paths):
    """بيقرا كل ملفات الإكسيل ويدمجهم في بيانات واحدة: الشرائح هي الأساسية، وكل ملف تاني بيبقى قسم لوحده.
    لو فيه أكتر من ملف بنفس الاسم بياخد الأحدث. الملفات اللي ما اتفهمتش بترجع في problems."""
    import os
    sl, cats, items, ins, outs, report, problems = None, [], [], [], [], [], []
    got = {}
    cache, keep, dirty = _cache(), set(), False
    for p in sorted(paths, key=os.path.getmtime):
        fn = os.path.basename(p)
        try:
            ck = _ckey(p); keep.add(ck)
            hit = cache.get(ck)
            if hit is None:                      # ملف جديد أو اتغيّر: نقراه (ونفتكره)
                k = _kind(p)
                if not k: problems.append({"file": fn, "reason": "مفيش فيه شيت اسمه 'الرصيد'، فمش شبه ملفات المخزن"}); continue
                if k == "slices":
                    key = "slices"
                    try: d = _read(p, True)
                    except Exception: d = _read(p, False)
                else:
                    key = _catid(p)
                    f = lambda ro: _read_generic(p, ro, key)   # القارئ الذكي لكل الأقسام الجديدة (بيفهم الأوزان لو موجودة من العناوين)
                    try: d = f(True)
                    except Exception:
                        d = f(False)
                    d.setdefault("wt", False)
                _mark_hidden(p, d)
                hit = cache[ck] = [key, k, d]; dirty = True
            key, k, d = hit
            got[key] = (p, k, copy.deepcopy(d))   # نسخة، لأن الدمج بعد كده بيعدّل في الصفوف
        except Exception as e:
            problems.append({"file": fn, "reason": "الملف مش بيتفتح: %s" % e})
    if dirty or len(cache) != len(keep): _cache_save(keep)
    if not got and not problems: raise ValueError("مفيش ملف إكسيل فيه شيت اسمه الرصيد")
    parts = []
    for key, (p, k, d) in got.items():
        parts.append(d)
        items += d["items"]; ins += d["ins"]; outs += d["outs"]
        cat = None
        if k != "slices":
            cat = _catname(p)
            cats.append({"id": key, "n": cat, "titles": d["titles"], "wt": bool(d.get("wt"))})
        report.append({"file": os.path.basename(p), "kind": k, "cat": cat or "شرائح", "items": len(d["items"]), "ins": len(d["ins"]), "outs": len(d["outs"]), "wt": bool(d.get("wt"))})
    _dedupe(got, items, ins, outs)
    sl = got["slices"][2] if "slices" in got else None
    dates = list(dict.fromkeys(x["asOf"] for x in parts if x["asOf"]))
    return {"titles": sl["titles"] if sl else [], "cats": cats, "report": report, "problems": problems,
            "asOf": dates[0] if len(dates) == 1 else " · ".join(dates), "items": items, "ins": ins, "outs": outs}

if __name__ == "__main__":
    import sys
    d = read_all(sys.argv[1:])
    print({k: (len(v) if isinstance(v, list) else v) for k, v in d.items()})
