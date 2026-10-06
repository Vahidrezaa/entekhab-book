#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_booklet.py  --  استخراج جدول رشته‌محل‌های دفترچه انتخاب رشته (PDF با لایه‌ی متنی) به بانک ابزار.

نیازمندی‌ها:  pip install pdfplumber      (و برای اعتبارسنجی مستقل: poppler-utils => دستور pdftotext)

مراحل کار:
  1) python extract_booklet.py book.pdf --calibrate [--page 2]      # مختصات ستون‌ها/فونت عنوان را نشان می‌دهد
  2) ثابت‌های بخش CONFIG را (اگر لازم بود) اصلاح کن
  3) python extract_booklet.py book.pdf --out out_dir               # استخراج + اعتبارسنجی + خروجی
  4) فایل out_dir/validation_report.txt باید بگوید: PROBLEMS = 0  و  همه‌ی چک‌ها PASS

خروجی‌ها:
  bank.json  (فرمت فایل بانک؛ در ابزار: «بازیابی پشتیبان JSON» یا دکمه‌ی بانک‌های آماده)
  bank.csv   (با ستون‌های «استان» و «دانشگاه»؛ در ابزار: «بارگذاری Excel / CSV»)
  validation_report.txt
"""
import argparse, collections, csv, json, os, re, subprocess, sys
import pdfplumber

# =====================================================================
# CONFIG  — مقادیر دفترچه ۱۴۰۴ (گروه ریاضی و فنی). برای دفترچه‌ی جدید با --calibrate تطبیق بده.
# =====================================================================
# مرز افقی ستون‌ها (محور x صفحه، از راست به چپ). هر ستون: (نام, x_min, x_max) روی «مرکز» کاراکتر سنجیده می‌شود.
COLS = [
    ('admission', 455, 530),   # نحوه پذیرش
    ('degree',    399, 455),   # دوره تحصیلی
    ('code',      365, 399),   # کد رشته محل
    ('major',     260, 365),   # عنوان رشته
    ('capFirst',  240, 260),   # ظرفیت نیمسال اول
    ('capSecond', 221, 240),   # ظرفیت نیمسال دوم
    ('cellF',     201, 221),   # جنس پذیرش: «زن» (یا عدد ظرفیت زنان / خط تیره)
    ('cellM',     181, 201),   # جنس پذیرش: «مرد» (یا عدد ظرفیت مردان / خط تیره)
    ('notes',      70, 181),   # توضیحات
]
# ستون‌هایی که خطوط افقی‌شان «فقط» مرز واقعی ردیف‌هاست (x0,x1 مستطیل‌های باریک). ستون عنوان/توضیحات را
# عمداً نمی‌گذاریم چون برای هر خط متن یک زیرسلول جدا می‌کشند و ردیف‌های چندخطی را می‌شکنند.
ROW_BORDER_X = [(455, 519), (399, 455), (365, 399)]
TITLE_FONT_SUBSTR = 'BTitrBold'   # فونت تیتر «استان … – دانشگاه …»
TITLE_SIZE = 10
CODE_RE = r'\d{4,6}'              # شکل کد رشته‌محل
X_TABLE_MIN, X_TABLE_MAX = 77, 519  # محدوده‌ی x جدول برای چک «کلمه‌ی تخصیص‌نشده»
SPACE_GAP = 6.0                    # فاصله‌ی افقی (pt) که اگر بین دو کاراکتر باشد کلمه جدا می‌شود (فاصله‌ی واقعی خودش کاراکتر ' ' دارد)

HEADER_KEYS = ('کد', 'ظرفیت', 'پذیرش', 'مرد', 'زن', 'اول', 'دوم', 'جنس', 'نیمسال', 'محل', 'دوره',
               'تحصیلی', 'عنوان', 'توضیحات', 'نحوه', 'رشته')


# =====================================================================
# متن فارسی: ترتیب بصری -> منطقی + نرمال‌سازی
# =====================================================================
def logical(word_visual):
    """pdfplumber کاراکترها را چپ‌به‌راست (ترتیب بصری) می‌دهد. برای کلمه‌ی فارسی باید معکوس شود،
    ولی رشته‌های لاتین/عدد (که در PDF چپ‌به‌راست‌اند) باید دست‌نخورده بمانند."""
    v = word_visual
    if re.search(r'[A-Za-z]', v) and not re.search(r'[\u0600-\u06FF]', v):
        return v                                   # کلمه‌ی لاتین/URL
    runs = re.findall(r'[0-9A-Za-z]+|[^0-9A-Za-z]+', v)
    out = []
    for r in reversed(runs):                       # ترتیب run ها معکوس
        out.append(r if re.fullmatch(r'[0-9A-Za-z]+', r) else ''.join(reversed(r)))  # ارقام/لاتین سالم، فارسی معکوس
    return ''.join(out)


def norm(s):
    s = s.replace('\u064a', '\u06cc').replace('\u0643', '\u06a9').replace('\u0649', '\u06cc')   # ي ك ى عربی -> فارسی
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2066-\u2069\u064b-\u065f\u0670]', '', s)          # RLM/LRM/... و اعراب
    s = s.translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789'))              # ارقام فارسی/عربی -> لاتین
    return re.sub(r'\s+', ' ', s).strip()


def chars_to_lines(chars):
    """کاراکترهای یک «سلول» -> لیست خطوط متن منطقی. کلمه‌ها فقط روی کاراکتر فاصله‌ی واقعی (یا فاصله‌ی خیلی زیاد) جدا می‌شوند؛
    استفاده از extract_words() کلمه‌ها را می‌شکست چون فاصله‌ی بین حروف فارسی در این PDF نامنظم است."""
    chars = list(chars)
    if not chars:
        return []
    chars.sort(key=lambda c: c['top'])
    lines = []
    for c in chars:                                           # گروه‌بندی خطوط بر اساس top
        if lines and abs(lines[-1]['top'] - c['top']) < 3.5:
            lines[-1]['cs'].append(c)
        else:
            lines.append({'top': c['top'], 'cs': [c]})
    out = []
    for ln in lines:
        cs = sorted(ln['cs'], key=lambda c: c['x0'])          # چپ -> راست (بصری)
        words, cur, prev = [], [], None
        for c in cs:
            if c['text'].strip() == '':                        # فاصله‌ی واقعی
                if cur:
                    words.append(''.join(cur)); cur = []
                prev = c
                continue
            if prev is not None and cur and (c['x0'] - prev['x1']) > SPACE_GAP and prev['text'].strip() != '':
                words.append(''.join(cur)); cur = []
            # گلیف لیگاتور «لا» یک کاراکتر دو-کدپوینتی است که از قبل «منطقی» (ل سپس ا) ذخیره شده؛
            # چون بعداً کل کلمه معکوس می‌شود باید از قبل «ال» شود تا بعد از معکوس‌سازی «لا» درآید.
            cur.append('\u0627\u0644' if c['text'] == '\u0644\u0627' else c['text'])
            prev = c
        if cur:
            words.append(''.join(cur))
        # کلمات از چپ به راست جمع شده‌اند؛ فارسی راست‌به‌چپ است => لیست کلمات هم معکوس می‌شود
        out.append(norm(' '.join(logical(w) for w in reversed(words))))
    return out


def chars_to_text(chars):
    return norm(' '.join(chars_to_lines(chars)))


# =====================================================================
# هندسه‌ی جدول
# =====================================================================
def near(a, b, t=2.5):
    return abs(a - b) < t


def col_of(xc):
    for name, a, b in COLS:
        if a <= xc < b:
            return name
    return None


def page_borders(page):
    """y خطوط افقی جداکننده‌ی ردیف‌ها: مستطیل‌های باریک (ارتفاع < 0.8pt) در ستون‌های ROW_BORDER_X."""
    ys = []
    for r in page.rects:
        if r['bottom'] - r['top'] < 0.8 and any(near(r['x0'], a) and near(r['x1'], b) for a, b in ROW_BORDER_X):
            ys.append((r['top'] + r['bottom']) / 2)
    ys.sort()
    out = []
    for y in ys:                                   # حذف خطوط تکراری (فاصله < 1.2pt)
        if not out or y - out[-1] >= 1.2:
            out.append(y)
    return out


# =====================================================================
# استخراج اصلی
# =====================================================================
def extract(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    records, problems, page_stats, seen_titles = [], [], [], []
    block = None                                   # بلوک (استان/دانشگاه) جاری؛ بین صفحات حمل می‌شود
    for pi, page in enumerate(pdf.pages, start=1):
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        borders = page_borders(page)
        if len(borders) < 2:
            problems.append(f'p{pi}: fewer than 2 row borders found')
            continue
        intervals = [(borders[i], borders[i + 1]) for i in range(len(borders) - 1)]

        # ---- تیترها (استان – دانشگاه): کاراکترهای فونت تیتر، گروه‌بندی بر اساس خط
        tchars = [c for c in page.chars if TITLE_FONT_SUBSTR in c['fontname'] and abs(c['size'] - TITLE_SIZE) < .1]
        tlines = []
        for c in sorted(tchars, key=lambda c: c['top']):
            if tlines and abs(tlines[-1]['top'] - c['top']) < 3:
                tlines[-1]['chars'].append(c)
            else:
                tlines.append({'top': c['top'], 'chars': [c]})
        titles = []
        for tl in tlines:
            xa = min(c['x0'] for c in tl['chars']) - 3
            xb = max(c['x1'] for c in tl['chars']) + 3
            tcs = [c for c in page.chars if abs(c['top'] - tl['top']) < 3 and c['x0'] >= xa and c['x1'] <= xb]
            txt_ = chars_to_text(tcs)
            if not re.search(r'[\u0600-\u06FF]', txt_):
                continue                           # خط فقط-نماد (مثل ✱ ستاره) تیتر نیست
            titles.append({'top': tl['top'], 'text': txt_, 'used': False})

        code_words = [w for w in words if COLS[2][1] <= (w['x0'] + w['x1']) / 2 < COLS[2][2]
                      and re.fullmatch(CODE_RE, w['text']) and borders[0] <= (w['top'] + w['bottom']) / 2 <= borders[-1]]
        assigned, n_rows, events = set(), 0, []
        for t in titles:
            events.append((t['top'], 0, t))
        for (a, b) in intervals:
            events.append((a, 1, (a, b, [w for w in words if a <= (w['top'] + w['bottom']) / 2 < b])))
        events.sort(key=lambda e: (e[0], e[1]))    # بالا->پایین؛ تیتر قبل از ردیف هم‌ارتفاع

        for y, kind, payload in events:
            if kind == 0:                          # ---------- تیتر بلوک
                t = payload
                t['used'] = True
                text = t['text']
                cont = text.startswith('ادامه')    # «ادامه استان …» = ادامه‌ی همان بلوک صفحه‌ی قبل
                text = re.sub(r'^ادامه\s*', '', text)
                star = bool(re.search(r'[*✱❁✽❋]', text))
                text = re.sub(r'[\s*✱❁✽❋]+$', '', text).strip()
                m = re.match(r'^استان\s+(.+?)\s*[–—\-‐‑]\s*(.+)$', text)
                if m:
                    new = {'province': m.group(1).strip(), 'university': m.group(2).strip(), 'raw': text, 'star': star}
                else:
                    new = {'province': '', 'university': text, 'raw': text, 'star': star}
                    problems.append(f'p{pi}: title not parsed as استان–دانشگاه: {text!r}')
                if cont and block and new['raw'] != block['raw']:
                    problems.append(f'p{pi}: «ادامه» title differs from carried block: {new["raw"]!r} vs {block["raw"]!r}')
                block = new
                seen_titles.append((pi, text, cont))
                continue

            a, b, rw = payload                     # ---------- ردیف داده
            for w in rw:
                assigned.add(id(w))
            cells = collections.defaultdict(list)
            for c in page.chars:
                if not (a <= (c['top'] + c['bottom']) / 2 < b):
                    continue
                if c['text'].strip() == '' and not (70 <= (c['x0'] + c['x1']) / 2 < 530):
                    continue
                cn = col_of((c['x0'] + c['x1']) / 2)
                if cn is None:
                    problems.append(f'p{pi} y={a:.0f}: char outside columns {c["text"]!r} x={c["x0"]:.0f}')
                else:
                    cells[cn].append(c)
            lines_ = {k: chars_to_lines(v) for k, v in cells.items()}
            txt = {k: norm(' '.join(v)) for k, v in lines_.items()}
            ncodes = sum(1 for w in rw if COLS[2][1] <= (w['x0'] + w['x1']) / 2 < COLS[2][2] and re.fullmatch(CODE_RE, w['text']))
            code = txt.get('code', '')
            if ncodes == 0:                        # ردیف بدون کد: باید هدر جدول باشد، وگرنه گزارش می‌شود
                if not rw:
                    continue
                alltxt = ' '.join(txt.values())
                if any(abs(t['top'] - w['top']) < 4 for t in titles for w in rw):
                    continue
                if any(k in alltxt for k in HEADER_KEYS) and not re.search(r'\d{4,6}', alltxt):
                    continue
                problems.append(f'p{pi} y={a:.0f}: row without code but with content {txt}')
                continue
            if ncodes > 1 or not re.fullmatch(CODE_RE, code):
                problems.append(f'p{pi} y={a:.0f}: bad code cell {code!r} (codes found {ncodes}) {txt}')
                continue
            n_rows += 1
            f, m_ = txt.get('cellF', ''), txt.get('cellM', '')
            gs = []
            for cv, nm in ((f, 'زن'), (m_, 'مرد')):
                if cv == nm or re.fullmatch(r'\d+', cv):
                    gs.append(nm)                  # کلمه یا عدد = این جنس پذیرفته می‌شود
                elif cv not in ('-', ''):
                    problems.append(f'p{pi} code {code}: unexpected gender cell {cv!r}')
            gender = ' و '.join(gs) if gs else 'نامشخص'
            if not gs:
                problems.append(f'p{pi} code {code}: both gender cells empty/dash')
            if not txt.get('admission') or not txt.get('degree') or not txt.get('major'):
                problems.append(f'p{pi} code {code}: missing admission/degree/major {txt}')
            for k in ('capFirst', 'capSecond'):
                v = txt.get(k, '')
                if v not in ('-', '') and not re.fullmatch(r'\d+', v):
                    problems.append(f'p{pi} code {code}: weird {k}={v!r}')
            if block is None:
                problems.append(f'p{pi} code {code}: no block title yet')
            records.append({
                'code': code, 'major': txt.get('major', ''),
                'university': block['university'] if block else '', 'province': block['province'] if block else '',
                'admission': txt.get('admission', ''), 'degree': txt.get('degree', ''),
                'gender': gender, 'notes': txt.get('notes', ''),
                'capFirst': '' if txt.get('capFirst', '-') in ('-', '') else txt['capFirst'],
                'capSecond': '' if txt.get('capSecond', '-') in ('-', '') else txt['capSecond'],
                'capFemale': f if re.fullmatch(r'\d+', f) else '', 'capMale': m_ if re.fullmatch(r'\d+', m_) else '',
                'star': bool(block and block.get('star')),
                'blockRaw': block['raw'] if block else '', 'page': pi,
                '_majorLines': lines_.get('major', []), '_notesLines': lines_.get('notes', []),
            })

        # ---- چک کامل بودن: هر کلمه‌ی داخل جدول باید به یک ردیف تخصیص یافته باشد
        for w in words:
            yc, xc = (w['top'] + w['bottom']) / 2, (w['x0'] + w['x1']) / 2
            if borders[0] <= yc <= borders[-1] and X_TABLE_MIN <= xc <= X_TABLE_MAX and id(w) not in assigned:
                problems.append(f'p{pi}: unassigned word in table area {logical(w["text"])!r} y={yc:.0f}')
        if len(code_words) != n_rows:
            problems.append(f'p{pi}: code-words={len(code_words)} but parsed rows={n_rows}')
        for t in titles:
            if not t['used']:
                problems.append(f'p{pi}: unused title {t["text"]!r}')
        page_stats.append((pi, n_rows, len(code_words), len(titles)))
    return records, problems, page_stats, seen_titles


# =====================================================================
# اعتبارسنجی مستقل با poppler (pdftotext)  — یک موتور استخراج متفاوت، با ترتیب منطقی متن
# =====================================================================
def pdftotext_page(pdf_path, p):
    return subprocess.run(['pdftotext', '-f', str(p), '-l', str(p), pdf_path, '-'],
                          capture_output=True).stdout.decode('utf8', 'ignore')


def _n(s, ours=False):
    # poppler گلیف لیگاتور «لا» را وارونه («ال») بیرون می‌دهد؛ برای مقایسه‌ی عادلانه فقط سمت «خودمان» را وارونه می‌کنیم.
    # (درستیِ «لا» در خروجی ما را عکس صفحه و جستجوی «گیلان/اطلاعات» تأیید کرده است.)
    if ours:
        s = s.replace('\u0644\u0627', '\u0627\u0644')
    s = s.replace('\u064a', '\u06cc').replace('\u0643', '\u06a9')
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2066-\u2069\u200c\u064b-\u065f\u0670]', '', s)
    s = s.translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789'))
    # فاصله/خط‌تیره/ستاره/پرانتز/گیومه/دونقطه نادیده؛ چون جای آن‌ها (به‌خصوص کنار اعداد) در دو موتور bidi فرق می‌کند
    return re.sub(r'[\s\-–—*()\[\]«»:،]+', '', s)


def validate_with_poppler(pdf_path, records, titles):
    lines = []
    ok_all = True
    pages = sorted({r['page'] for r in records})
    ptxt = {p: pdftotext_page(pdf_path, p) for p in range(1, len(pdfplumber.open(pdf_path).pages) + 1)}
    pn = {p: _n(t) for p, t in ptxt.items()}

    # (1) مجموعه‌ی کدهای هر صفحه == کدهای ۵رقمی متن poppler
    bad = []
    for p in pages:
        raw = ptxt[p].translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789'))
        cand = set(re.findall(r'(?<!\d)\d{5}(?!\d)', raw))
        ours = {r['code'] for r in records if r['page'] == p}
        if cand != ours:
            bad.append((p, sorted(cand - ours)[:5], sorted(ours - cand)[:5]))
    lines.append(f'[{"PASS" if not bad else "FAIL"}] code set per page vs poppler: {len(bad)} mismatching pages {bad[:5]}')
    ok_all &= not bad

    # (2) هر خط از «عنوان رشته» و «توضیحات» باید به‌عنوان زیررشته در متن همان صفحه‌ی poppler باشد
    miss = []
    for r in records:
        for key in ('_majorLines', '_notesLines'):
            for ln in r[key]:
                v, v2 = _n(ln, ours=True), _n(ln)
                if v and v not in pn[r['page']] and v2 not in pn[r['page']]:
                    miss.append((r['page'], r['code'], key, ln))
    lines.append(f'[{"PASS" if not miss else "FAIL"}] major/notes lines vs poppler text: {len(miss)} mismatches')
    for m in miss[:15]:
        lines.append('      ' + repr(m))
    ok_all &= not miss

    # (3) عنوان بلوک‌ها
    tb = [(p, t) for p, t, c in titles if _n(t, ours=True) not in pn[p] and _n(t) not in pn[p]]
    lines.append(f'[{"PASS" if not tb else "FAIL"}] block titles vs poppler: {len(tb)} mismatches {tb[:5]}')
    ok_all &= not tb

    # (4) یکتایی کد
    c = collections.Counter(r['code'] for r in records)
    dups = {k: v for k, v in c.items() if v > 1}
    lines.append(f'[{"PASS" if not dups else "FAIL"}] duplicate codes: {len(dups)} {list(dups)[:5]}')
    ok_all &= not dups
    return ok_all, lines


# =====================================================================
# خروجی
# =====================================================================
def write_outputs(records, outdir, title='', year=None, group=''):
    """خروجی با «فرمت فایل بانک» ابزار (docs/data-format.md). همین فایل مستقیم در ابزار بارگذاری می‌شود."""
    keys = ['code', 'major', 'university', 'province', 'admission', 'degree', 'gender', 'capFirst', 'capSecond',
            'capFemale', 'capMale', 'notes', 'star', 'page', 'blockRaw']
    bank = []
    for r in records:
        o = {k: r[k] for k in keys}
        o['page'] = str(o['page'])
        bank.append(o)
    meta = {'title': title or 'بانک استخراج‌شده', 'group': group, 'year': year, 'rows': len(bank),
            'source': 'extract_booklet.py'}
    data = {'format': 'entekhab-reshte-bank', 'version': 1, 'meta': meta, 'bank': bank}
    with open(os.path.join(outdir, 'bank.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    with open(os.path.join(outdir, 'bank.csv'), 'w', newline='', encoding='utf-8-sig') as f:   # BOM برای Excel
        w = csv.writer(f)
        w.writerow(['کد رشته', 'عنوان رشته', 'دانشگاه', 'استان', 'نحوه پذیرش', 'دوره تحصیلی', 'جنس پذیرش',
                    'ظرفیت اول', 'ظرفیت دوم', 'ظرفیت زن', 'ظرفیت مرد', 'توضیحات',
                    'ستاره‌دار (مصاحبه/گزینش)', 'صفحه دفترچه', 'نوع ردیف'])
        for r in records:
            w.writerow([r['code'], r['major'], r['university'], r['province'], r['admission'], r['degree'],
                        r['gender'], r['capFirst'], r['capSecond'], r['capFemale'], r['capMale'], r['notes'],
                        'بله' if r['star'] else '', r['page'], 'اصلی'])


def summary_lines(records):
    L = [f'records={len(records)}  pages={len({r["page"] for r in records})}  '
         f'provinces={len({r["province"] for r in records})}  blocks={len({r["blockRaw"] for r in records})}']
    for k in ('admission', 'degree', 'gender'):
        L.append(f'{k}: ' + str(collections.Counter(r[k] for r in records).most_common(10)))
    L.append('code length: ' + str(collections.Counter(len(r['code']) for r in records)))
    L.append(f'rows with both capacities empty: {sum(1 for r in records if not r["capFirst"] and not r["capSecond"] and not (r["capFemale"] or r["capMale"]))}')
    return L


# =====================================================================
# --calibrate : کمک برای دفترچه‌ی جدید
# =====================================================================
def calibrate(pdf_path, page_no):
    pdf = pdfplumber.open(pdf_path)
    p = pdf.pages[page_no - 1]
    print(f'page {page_no}: size={p.width:.0f}x{p.height:.0f}  chars={len(p.chars)}  rects={len(p.rects)}  lines={len(p.lines)}')
    print('\n-- فونت/اندازه‌ی کاراکترها (تیتر معمولاً کمتعدادترین فونت بولد بزرگ‌تر است):')
    for k, v in collections.Counter((c['fontname'], round(c['size'], 1)) for c in p.chars).most_common(10):
        print('  ', k, v)
    print('\n-- جفت (x0,x1) مستطیل‌ها (ستون‌های جدول؛ مرز ستون‌ها را از اینجا بخوان):')
    for k, v in collections.Counter((round(r['x0']), round(r['x1'])) for r in p.rects).most_common(25):
        print('  ', k, v)
    print('\n-- مرزهای x مستطیل‌های باریک افقی (کاندید ROW_BORDER_X):')
    thin = collections.Counter((round(r['x0']), round(r['x1'])) for r in p.rects if r['bottom'] - r['top'] < 0.8)
    for k, v in thin.most_common(15):
        print('  ', k, v)
    print('\n-- مرکز x کلمات هدر جدول (برای نگاشت ستون‌ها):')
    for w in p.extract_words():
        if any(h in logical(w['text']) for h in ('نحوه', 'دوره', 'عنوان', 'توضیحات', 'ظرفیت', 'جنس', 'کد')):
            print(f'   {logical(w["text"])!r:20} x0={w["x0"]:.0f} x1={w["x1"]:.0f} top={w["top"]:.0f}')
    print('\nحالا COLS / ROW_BORDER_X / TITLE_FONT_SUBSTR / TITLE_SIZE را در بخش CONFIG تنظیم کن و اسکریپت را بدون --calibrate اجرا کن.')
    print('اگر مرزها اشتباه باشند، اجرای اصلی با پیام‌های «char outside columns» / «unassigned word» / «code-words != rows» فریاد می‌زند.')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('--out', default='out')
    ap.add_argument('--calibrate', action='store_true')
    ap.add_argument('--page', type=int, default=2)
    ap.add_argument('--skip-poppler', action='store_true')
    ap.add_argument('--title', default='', help='عنوان بانک (مثلاً «ریاضی و فنی ۱۴۰۵»)')
    ap.add_argument('--year', type=int, default=None)
    ap.add_argument('--group', default='')
    a = ap.parse_args()
    if a.calibrate:
        calibrate(a.pdf, a.page)
        return
    os.makedirs(a.out, exist_ok=True)
    records, problems, stats, titles = extract(a.pdf)
    rep = [f'PDF: {a.pdf}', f'PROBLEMS = {len(problems)}'] + ['  ' + p[:220] for p in problems[:80]]
    rep += summary_lines(records)
    ok = not problems
    if not a.skip_poppler:
        try:
            ok2, lines = validate_with_poppler(a.pdf, records, titles)
            rep += lines
            ok &= ok2
        except FileNotFoundError:
            rep.append('[SKIP] pdftotext (poppler-utils) not installed — independent validation skipped')
            ok = False
    write_outputs(records, a.out, a.title, a.year, a.group)
    rep.append('RESULT: ' + ('ALL CHECKS PASSED' if ok else 'CHECKS FAILED — do not use the data before fixing'))
    open(os.path.join(a.out, 'validation_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
    print('\n'.join(rep))


if __name__ == '__main__':
    main()
