#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_booklet.py  --  استخراج جدول رشته‌محل‌های «دفترچه راهنمای انتخاب رشته» (PDF با لایه‌ی متنی) به فایل بانک ابزار.

نیازمندی‌ها:  pip install pdfplumber      (و برای اعتبارسنجی مستقل: poppler-utils => دستور pdftotext)

نسخه‌ی ۲ — ویژگی‌ها
  • ساختار ستون‌ها را خودکار از «سطر هدر» هر جدول (خانه‌های رنگی) تشخیص می‌دهد؛ برای چند چیدمان مختلف در یک دفترچه
    (جدول اصلی، سهمیه مناطق محروم، فرهنگیان، غیرانتفاعی، پیام‌نور، بهداشت، بومی) بدون تنظیم دستی کار می‌کند.
  • مرز ردیف‌ها از خطوط موی افقی (hairline) خوانده می‌شود، نه از مختصات ثابت؛ پس بلوک‌های جابه‌جاشده را هم می‌گیرد.
  • جدول «نوع گزینش» (کشوری/ناحیه‌ای/قطبی/استانی) را از ابتدای دفترچه می‌خواند و به رشته‌محل‌های روزانه نسبت می‌دهد.
  • هر صفحه یک‌بار پارس و (اختیاری) کش می‌شود (--cache)؛ پارس مجدد برای دفترچه‌ی بزرگ لازم نیست.

مراحل کار:
  1) python extract_booklet.py book.pdf --calibrate 8             # ساختار ستون‌ها/تیترهای یک صفحه را نشان می‌دهد
  2) python extract_booklet.py book.pdf --out out_dir --cache cache_dir --title "..." --year 1405 --group "ریاضی و فنی"
  3) out_dir/validation_report.txt باید بگوید: PROBLEMS = 0  و  RESULT: ALL CHECKS PASSED

خروجی‌ها:
  bank.json  (فرمت فایل بانک؛ در ابزار: «بازیابی پشتیبان JSON»)
  bank.csv   (در ابزار: «بارگذاری Excel / CSV»)
  validation_report.txt
"""
import argparse, bisect, collections, csv, json, os, pickle, re, subprocess, sys, time

# =====================================================================
# صفحه: هندسه‌ی لازم (کاراکترها، خطوط موی افقی، خانه‌های رنگی هدر) — قابل کش
# =====================================================================
def _color(c):
    if c is None: return None
    try: return tuple(round(float(x), 3) for x in c)
    except Exception: return None

def _is_colored(t):
    return t is not None and len(t) == 3 and not (abs(t[0] - t[1]) < .02 and abs(t[1] - t[2]) < .02)

class PageData:
    __slots__ = ('number', 'width', 'height', 'chars', 'thin', 'fills')
    def __init__(self, number, width, height, chars, thin, fills):
        self.number, self.width, self.height = number, width, height
        self.chars, self.thin, self.fills = chars, thin, fills
    def to_dict(self):
        return {'number': self.number, 'width': self.width, 'height': self.height,
                'chars': self.chars, 'thin': self.thin, 'fills': self.fills}

def from_plumber(page, number):
    chars = [{'text': c['text'], 'x0': round(c['x0'], 2), 'x1': round(c['x1'], 2), 'top': round(c['top'], 2),
              'bottom': round(c['bottom'], 2), 'size': round(c['size'], 1), 'font': c['fontname'].split('+')[-1],
              'color': _color(c.get('non_stroking_color'))} for c in page.chars]
    thin, fills = [], []
    for r in page.rects:
        h = r['bottom'] - r['top']
        if h < 0.8:
            thin.append((round(r['x0'], 2), round(r['x1'], 2), round((r['top'] + r['bottom']) / 2, 2)))
        col = _color(r.get('non_stroking_color'))
        if _is_colored(col) and h > 6 and 8 < (r['x1'] - r['x0']) < 300:
            fills.append({'x0': round(r['x0'], 2), 'x1': round(r['x1'], 2), 'top': round(r['top'], 2),
                          'bottom': round(r['bottom'], 2), 'color': col})
    return PageData(number, float(page.width), float(page.height), chars, thin, fills)

class PageSource:
    """صفحه‌ها را از PDF (با pdfplumber) یا از کش می‌خواند؛ هر صفحه یک‌بار پارس و در کش ذخیره می‌شود."""
    def __init__(self, pdf_path, cache_dir=None):
        self.pdf_path, self.cache_dir, self._pdf = pdf_path, cache_dir, None
        if cache_dir: os.makedirs(cache_dir, exist_ok=True)
    def count(self):
        return len(self.pdf().pages)
    def pdf(self):
        if self._pdf is None:
            import pdfplumber
            self._pdf = pdfplumber.open(self.pdf_path)
        return self._pdf
    def get(self, n):
        path = os.path.join(self.cache_dir, f'p{n:04d}.pkl') if self.cache_dir else None
        if path and os.path.exists(path):
            with open(path, 'rb') as f:
                d = pickle.load(f)
            return PageData(d['number'], d['width'], d['height'], d['chars'], d['thin'], d['fills'])
        page = self.pdf().pages[n - 1]
        pd = from_plumber(page, n)
        page.flush_cache()
        if path:
            with open(path + '.tmp', 'wb') as f: pickle.dump(pd.to_dict(), f, protocol=4)
            os.replace(path + '.tmp', path)
        return pd

# =====================================================================
# CONFIG — مقادیر دفترچه‌های سری «راهنمای انتخاب رشته». برای دفترچه‌ی دیگر با --calibrate تطبیق بده.
# =====================================================================
DIGIT_MAP = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')
SPACE_GAP = 6.0

def logical(w):
    if re.search(r'[A-Za-z]', w) and not re.search(r'[\u0600-\u06FF]', w):
        return w
    runs = re.findall(r'[0-9A-Za-z]+|[^0-9A-Za-z]+', w)
    return ''.join(r if re.fullmatch(r'[0-9A-Za-z]+', r) else ''.join(reversed(r)) for r in reversed(runs))

def norm(s):
    s = s.replace('\u064a', '\u06cc').replace('\u0643', '\u06a9').replace('\u0649', '\u06cc')
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2066-\u2069\u064b-\u065f\u0670]', '', s)
    s = s.translate(DIGIT_MAP)
    return re.sub(r'\s+', ' ', s).strip()

def chars_to_lines(chars):
    chars = list(chars)
    if not chars: return []
    chars.sort(key=lambda c: c['top'])
    lines = []
    for c in chars:
        if lines and abs(lines[-1]['top'] - c['top']) < 3.5: lines[-1]['cs'].append(c)
        else: lines.append({'top': c['top'], 'cs': [c]})
    out = []
    for ln in lines:
        cs = sorted(ln['cs'], key=lambda c: c['x0'])
        words, cur, prev = [], [], None
        for c in cs:
            if c['text'].strip() == '':
                if cur: words.append(''.join(cur)); cur = []
                prev = c; continue
            if prev is not None and cur and (c['x0'] - prev['x1']) > SPACE_GAP and prev['text'].strip() != '':
                words.append(''.join(cur)); cur = []
            ch = c['text'].translate(DIGIT_MAP)
            cur.append('\u0627\u0644' if ch == '\u0644\u0627' else ch)
            prev = c
        if cur: words.append(''.join(cur))
        out.append(norm(' '.join(logical(w) for w in reversed(words))))
    return out

def chars_to_text(chars):
    return norm(' '.join(chars_to_lines(chars)))


# ---------------- CONFIG ----------------
TITLE_FONT = 'Titr'; TITLE_SIZE = 10; TITLE_COLOR = (0.502, 0.0, 0.0)
HEADING_MIN_SIZE = 12.5
CODE_RE = r'\d{4,6}'
FIELD_ORDER = ['cellM', 'cellF', 'notes', 'domain', 'service', 'admission', 'degree', 'major', 'code',
               'capSecond', 'capFirst', 'capacity', 'gender']

def near_color(c, t, tol=0.06):
    return c is not None and len(c) == len(t) and all(abs(a - b) <= tol for a, b in zip(c, t))

def field_of(label):
    t = re.sub(r'\s+', '', label)
    if 'مرد' in t: return 'cellM'
    if 'زن' in t: return 'cellF'
    if 'توضیحات' in t: return 'notes'
    if 'دامنه' in t: return 'domain'
    if 'خدمت' in t: return 'service'
    if 'نحوه' in t: return 'admission'
    if 'دوره' in t and 'نیمسال' not in t: return 'degree'
    if 'عنوان' in t: return 'major'
    if t.startswith('کد'): return 'code'
    if 'دوم' in t: return 'capSecond'
    if 'اول' in t: return 'capFirst'
    if 'ظرفیت' in t: return 'capacity'
    if 'جنس' in t: return 'gender'
    return None

class Layout:
    def __init__(self, cols, band_top, labels=None):
        self.labels = labels or {}
        self.cols = cols                       # [(field, x0, x1)]
        self.fields = [f for f, _, _ in cols]
        self.code = next((a, b) for f, a, b in cols if f == 'code')
        self.sig = '|'.join(self.fields)
        self.band_top = band_top
    def col_of(self, xc):
        for f, a, b in self.cols:
            if a - 0.5 <= xc < b + 0.5: return f
        return None

def header_bands(page):
    out = []
    for r in sorted(page.fills, key=lambda r: r['top']):
        if out and r['top'] < out[-1]['bottom'] + 3:
            out[-1]['rs'].append(r); out[-1]['bottom'] = max(out[-1]['bottom'], r['bottom'])
        else:
            out.append({'top': r['top'], 'bottom': r['bottom'], 'rs': [r]})
    return out

def build_layout(page, band, problems):
    xs = sorted({round(r['x0']) for r in band['rs']} | {round(r['x1']) for r in band['rs']})
    edges = []
    for x in xs:
        if not edges or x - edges[-1] > 2: edges.append(x)
    cols = []
    for a, b in zip(edges, edges[1:]):
        if b - a < 8: continue
        cs = [c for c in page.chars if band['top'] - 1 <= (c['top'] + c['bottom']) / 2 <= band['bottom'] + 1
              and a <= (c['x0'] + c['x1']) / 2 < b]
        label = chars_to_text(cs)
        f = field_of(label)
        cols.append((f, a, b, label))
    named = [(f, a, b) for f, a, b, l in cols if f]
    unnamed = [(a, b, l) for f, a, b, l in cols if not f]
    fields = [f for f, _, _ in named]
    ok = 'code' in fields and 'major' in fields and len(set(fields)) == len(fields)
    if not ok:
        return None, f'p{page.number}: header not recognised: ' + ' | '.join(f'{f}:{l}' for f, a, b, l in cols)
    if unnamed:
        problems.append(f'p{page.number}: unnamed header column(s) ignored: {unnamed}')
    return Layout(named, band['top'], {f: l for f, a, b, l in cols if f}), None

def row_borders(page, ncols, y0, y1):
    """Row borders = y positions where many (>= ncols-3) column-wide hairlines are drawn at once. Layout-agnostic,
    so it also works for blocks whose columns are shifted/scaled relative to the page header (e.g. blocks with a note box)."""
    items = sorted(((y, x0, x1) for x0, x1, y in page.thin if x1 - x0 > 8 and y0 - 0.5 <= y < y1))
    groups = []
    for y, x0, x1 in items:
        if groups and y - groups[-1]['y'] < 0.7:
            groups[-1]['spans'].add((round(x0), round(x1)))
        else:
            groups.append({'y': y, 'spans': {(round(x0), round(x1))}})
    need = max(4, ncols - 3)
    out = []
    for g in groups:
        if len(g['spans']) >= need:
            if out and g['y'] - out[-1][0] < 1.2:
                continue
            out.append((g['y'], sorted(g['spans'])))
    return out

def map_spans(spans, layout):
    """spans of one row (sorted by x) -> [(field,x0,x1)]; positional when counts match, otherwise by overlap, otherwise layout."""
    cols = layout.cols
    # merge duplicate/adjacent spans (e.g. (448,449)+(449,510))
    merged = []
    for a, b in spans:
        if merged and a <= merged[-1][1] + 1 and a >= merged[-1][0] and (b - a) < 3: continue
        merged.append((a, b))
    if len(merged) == len(cols):
        return [(f, a, b) for (f, _, _), (a, b) in zip(cols, merged)]
    out = []
    for f, ca, cb in cols:
        best = max(merged, key=lambda s: max(0, min(s[1], cb) - max(s[0], ca)), default=None)
        if best and max(0, min(best[1], cb) - max(best[0], ca)) > 0.5 * (cb - ca): out.append((f, best[0], best[1]))
        else: out.append((f, ca, cb))
    return out

CONT_COLOR = (0.0, 0.4, 0.0)      # کلمه‌ی «ادامه» در ابتدای تیتر بلوک‌های ادامه‌دار: سبز، همان فونت تیتر

def title_lines(page):
    tc = [c for c in page.chars if TITLE_FONT in c['font'] and abs(c['size'] - TITLE_SIZE) <= 0.6
          and (near_color(c['color'], TITLE_COLOR) or near_color(c['color'], CONT_COLOR))]
    stars = [c for c in page.chars if c['text'] in ('*', '✱', '❁', '✽', '❋') and TITLE_FONT in c['font']]
    lines = []
    for c in sorted(tc, key=lambda c: c['top']):
        if lines and abs(lines[-1]['top'] - c['top']) < 3: lines[-1]['cs'].append(c)
        else: lines.append({'top': c['top'], 'cs': [c]})
    merged = []
    for ln in lines:
        if merged and ln['top'] - merged[-1]['last'] < 24:
            merged[-1]['lines'].append(ln); merged[-1]['last'] = ln['top']
        else:
            merged.append({'top': ln['top'], 'last': ln['top'], 'lines': [ln]})
    out = []
    for m in merged:
        cs = [c for ln in m['lines'] for c in ln['cs']]
        lo, hi = m['top'] - 6, max(c['bottom'] for c in cs) + 6
        star = any(lo <= c['top'] <= hi for c in stars)           # نماد ✱ = «اسامی چندبرابر ظرفیت برای مصاحبه/گزینش»
        txt = norm(' '.join(chars_to_lines(cs)))
        out.append({'top': m['top'], 'bottom': hi - 6, 'text': txt, 'star': star, 'chars': cs})
    return out

def heading_lines(page):
    hc = [c for c in page.chars if TITLE_FONT in c['font'] and c['size'] >= HEADING_MIN_SIZE]
    lines = []
    for c in sorted(hc, key=lambda c: c['top']):
        if lines and abs(lines[-1]['top'] - c['top']) < 3: lines[-1]['cs'].append(c)
        else: lines.append({'top': c['top'], 'cs': [c]})
    return [(ln['top'], norm(chars_to_text(ln['cs']))) for ln in lines]

SECTION_RULES = [
    ('وزارت علوم', 'وزارت علوم، تحقیقات و فناوری'),
    ('سهمیه مناطق محروم', 'سهمیه مناطق محروم و بلایای طبیعی'),
    ('فرهنگیان', 'دانشگاه فرهنگیان و تربیت دبیر شهید رجایی'),
    ('غیرانتفاعی', 'مؤسسات آموزش عالی غیردولتی-غیرانتفاعی'),
    ('پیام نور', 'دانشگاه پیام‌نور'),
    ('وزارت بهداشت', 'وزارت بهداشت، درمان و آموزش پزشکی'),
    ('سهمیه بومی', 'سهمیه بومی (تعهد خدمت وزارت بهداشت)'),
]
def section_of(heading):
    h = heading.replace('\u200c', '').replace('پیامنور', 'پیام نور')
    for key, name in SECTION_RULES:
        if key in h: return name
    return None

STAR = '*✱❁✽❋'
def parse_title(text):
    t = norm(text)
    cont = t.startswith('ادامه')
    t = re.sub(r'^ادامه\s*', '', t)
    star = any(ch in t for ch in STAR)
    t = re.sub(r'^[\s' + re.escape(STAR) + r']+|[\s' + re.escape(STAR) + r']+$', '', t).strip()   # ستاره‌ی تیتر ممکن است ابتدا یا انتها باشد
    d = {'raw': t, 'cont': cont, 'star': star, 'kind': None, 'province': '', 'university': '', 'nativeProvince': '', 'note': ''}
    m = re.match(r'^(?:سهمیه\s+)?مخصوص متقاضیان بومی\s+(.*)$', t)
    if m or t.startswith('سهمیه مخصوص'):
        d['kind'] = 'native'
        rest = (m.group(1) if m else t)
        note = ''
        if re.search(r'\sبا\s+اولویت', rest):
            i = rest.index('با ' + 'اولویت') if 'با اولویت' in rest else len(rest)
            rest, note = rest[:i].strip(), rest[i:].strip()
        if 'بشاگرد' in rest:
            prov = 'هرمزگان (شهرستان بشاگرد)'
        else:
            ms = re.search(r'استان\s+(.+?)\s*(?:\(|$)', rest)
            prov = ms.group(1).strip() if ms else rest.strip()
            prov = re.sub(r'^محروم\s+', '', prov)               # «استان محروم ایلام» -> ایلام
        d['nativeProvince'] = prov; d['note'] = note
        return d
    m = re.match(r'^دانشگاه پیام\s*نور\s+استان\s+(.+?)\s*[–—\-‐‑]\s*(.+)$', t)
    if m:
        d.update(kind='main', province=m.group(1).strip(), university='دانشگاه پیام نور - ' + m.group(2).strip()); return d
    m = re.match(r'^استان\s+(.+?)\s*[–—\-‐‑]\s*(.+)$', t)
    if m:
        d.update(kind='main', province=m.group(1).strip(), university=m.group(2).strip()); return d
    return d


# ======================= row extraction =======================
NOTE_START = re.compile(r'^(دارای|فاقد|عدم|محدودیت|ممنوعیت|معرفی|ارائه|خوابگاه|تعهد|مشروط|شرایط|با\s|بدون|صرفا|پذیرش|متقاضیان|امکان|محل\s|مصاحبه|ظرفیت|طبق|حداکثر|حداقل|کلیه|تنها|فقط|پرداخت|واگذاری|اولویت|دوره|مقطع)')

def split_notes_university(text):
    """layouts whose notes header is «دانشگاه … محل تحصیل / توضیحات»: the university name comes first, then the notes,
    separated by ' - '. A campus/city suffix is also joined by ' - ' («دانشگاه هرمزگان - بندرعباس»), so segments are
    attached to the name until the first segment that starts like a note («دارای/فاقد/عدم/محدودیت/…»)."""
    segs, depth, cur = [], 0, ''
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '(': depth += 1
        elif ch == ')': depth = max(0, depth - 1)
        if depth == 0 and text[i:i + 3] == ' - ':
            segs.append(cur); cur = ''; i += 3; continue
        cur += ch; i += 1
    segs.append(cur)
    segs = [x.strip() for x in segs if x.strip()]
    for k in range(1, len(segs)):
        if NOTE_START.match(segs[k]):
            return ' - '.join(segs[:k]), ' - '.join(segs[k:])
    return ' - '.join(segs), ''

def split_domain(domain):
    """native (فرهنگیان) rows: «پردیس X / اولویت پذیرش با ...» -> campus, rest"""
    parts = [p.strip() for p in re.split(r'\s/\s|/', domain, maxsplit=1)]
    return parts[0], (parts[1] if len(parts) > 1 else '')

def extract(pages, progress=None):
    """pages: iterable of PageData in order. Returns records, problems, stats."""
    records, problems, stats, selection_rows = [], [], [], []
    layout = None; block = None; section = ''; block_first = 0; pending_star = False
    for page in pages:
        pn = page.number
        # sort chars by center-y for fast interval queries
        chars = sorted(page.chars, key=lambda c: (c['top'] + c['bottom']) / 2)
        ys = [(c['top'] + c['bottom']) / 2 for c in chars]
        def chars_in(y0, y1):
            return chars[bisect.bisect_left(ys, y0):bisect.bisect_left(ys, y1)]
        # headings -> section
        for top, h in heading_lines(page):
            s = section_of(h)
            if s and (h.startswith('رشته') or 'رشته‌محل' in h or 'رشتهمحل' in h.replace('\u200c', '')):
                section = s
        # layouts
        bands = header_bands(page)
        regions = []                                    # (y_start, y_end, layout)
        cur = layout
        starts = [(0.0, None)] + [(b['top'], b) for b in bands]
        for i, (ys0, b) in enumerate(starts):
            if b is not None:
                btxt = band_text(page, b)
                if 'گزینش' in btxt.replace('\u200c', '') and 'نوع' in btxt:
                    selection_rows.extend(parse_selection_page(page, b, chars, ys, problems))
                    cur = None
                else:
                    lay, err = build_layout(page, b, problems)
                    if err:
                        cur = None
                        if 'کد' in btxt and 'رشته' in btxt: problems.append(err)      # a data-table header we cannot read
                    else: cur = lay
            y_end = starts[i + 1][0] if i + 1 < len(starts) else 1e9
            regions.append((ys0, y_end, cur, b))
        layout = cur
        titles = title_lines(page)
        title_ids = {id(c) for t in titles for c in t['chars']}
        events = [(t['top'], 0, t) for t in titles]
        n_rows = 0; n_codes = 0
        for y0, y1, lay, band in regions:
            if lay is None:
                continue
            borders = row_borders(page, len(lay.cols), y0, y1)
            for i in range(len(borders) - 1):
                events.append((borders[i][0], 1, (borders[i][0], borders[i + 1][0], lay, borders[i][1])))
        events.sort(key=lambda e: (e[0], e[1]))
        for y, kind, payload in events:
            if kind == 0:
                t = payload
                if not t['text'] or not re.search(r'[\u0600-\u06FF]', t['text']):
                    continue
                nb = parse_title(t['text'])
                nb['star'] = nb['star'] or t.get('star', False) or pending_star
                pending_star = False
                if nb['kind'] is None:
                    problems.append(f'p{pn}: title not recognised: {t["text"]!r}')
                    continue
                if nb['cont'] and block and nb['raw'] != block['raw']:
                    problems.append(f'p{pn}: «ادامه» title differs from carried block: {nb["raw"]!r} vs {block["raw"]!r}')
                if nb['cont'] and block and nb['raw'] == block['raw']:
                    nb['star'] = nb['star'] or block['star']
                else:
                    block_first = len(records)
                block = nb
                continue
            a, b, lay, spans = payload
            colmap = map_spans(spans, lay)
            cells = collections.defaultdict(list)
            for c in chars_in(a, b):
                xc = (c['x0'] + c['x1']) / 2
                if id(c) in title_ids: continue
                f = next((ff for ff, ca, cb in colmap if ca - 0.5 <= xc < cb + 0.5), None)
                if f is None:
                    if c['text'].strip() == '': continue
                    problems.append(f'p{pn} y={a:.0f}: char outside columns {c["text"]!r} x={c["x0"]:.0f}')
                else:
                    cells[f].append(c)
            lines_ = {k: chars_to_lines(v) for k, v in cells.items()}
            txt = {k: norm(' '.join(v)) for k, v in lines_.items()}
            code = txt.get('code', '')
            if not code:
                # header or title-zone/footnote interval
                alltxt = ' '.join(txt.values())
                if not alltxt: continue
                if re.search(r'\d{4,6}', alltxt):
                    problems.append(f'p{pn} y={a:.0f}: interval without code cell but digits present {txt}')
                continue
            if not re.fullmatch(CODE_RE, code):
                if not re.search(r'\d', code) and re.search(r'کد|رشته|محل', code):
                    continue                                   # header row of the table
                if not re.search(r'\d', ' '.join(txt.values())) and len(' '.join(txt.values())) >= 40:
                    # long footnote paragraph between blocks. The ✱ note («اسامی چندبرابر ظرفیت … برای مصاحبه و گزینش»)
                    # marks the block right above it, even when the star glyph itself is an image / not in the text layer.
                    if 'چندبرابر' in ' '.join(txt.values()).replace(' ', ''):
                        in_gap = [t for t in titles if a <= t['top'] < b]
                        note_y = [c['top'] for c in chars_in(a, b) if id(c) not in title_ids and c['text'].strip()]
                        below_title = bool(in_gap) and note_y and (sum(note_y) / len(note_y)) > max(t['bottom'] for t in in_gap)
                        if below_title:
                            pending_star = True                 # the note sits right under a new block title (before its rows)
                        elif block is not None:                 # the note closes the current block (after its rows)
                            block['star'] = True
                            for r_ in records[block_first:]: r_['star'] = True
                    continue
                problems.append(f'p{pn} y={a:.0f}: bad code cell {code!r} {txt}')
                continue
            n_rows += 1
            if block is None:
                problems.append(f'p{pn} code {code}: no block title yet')
            rec = build_record(code, txt, lay, block, section, pn, problems)
            rec['_majorLines'] = lines_.get('major', []); rec['_notesLines'] = lines_.get('notes', [])
            records.append(rec)
        stats.append((pn, n_rows))
        if progress: progress(pn)
    return records, problems, stats, selection_rows

def build_record(code, txt, lay, block, section, pn, problems):
    f, m_ = txt.get('cellF', ''), txt.get('cellM', '')
    if 'gender' in lay.fields:
        g = txt.get('gender', '')
        gender = g if g in ('زن', 'مرد') else ''
        if not gender: problems.append(f'p{pn} code {code}: gender cell {g!r}')
        capF = capM = ''
    else:
        gs = []
        for cv, nm in ((f, 'زن'), (m_, 'مرد')):
            if cv == nm or re.fullmatch(r'\d+', cv): gs.append(nm)
            elif cv not in ('-', ''): problems.append(f'p{pn} code {code}: unexpected gender cell {cv!r}')
        gender = ' و '.join(gs) if gs else 'نامشخص'
        if not gs: problems.append(f'p{pn} code {code}: both gender cells empty/dash')
        capF = f if re.fullmatch(r'\d+', f) else ''; capM = m_ if re.fullmatch(r'\d+', m_) else ''
    def cap(k):
        v = txt.get(k, '')
        if v in ('-', ''): return ''
        if not re.fullmatch(r'\d+', v): problems.append(f'p{pn} code {code}: weird {k}={v!r}')
        return v
    capFirst = cap('capFirst') or cap('capacity'); capSecond = cap('capSecond')
    native = block is not None and block['kind'] == 'native'
    domain = txt.get('domain', ''); service = txt.get('service', '')
    university = block['university'] if block else ''; province = block['province'] if block else ''
    notes = txt.get('notes', '')
    if native and domain:
        campus, rest = split_domain(domain)
        university = campus
    elif 'دانشگاه' in lay.labels.get('notes', '') and notes:
        university, notes = split_notes_university(notes)
    rec = {
        'code': code, 'major': txt.get('major', ''), 'university': university, 'province': province,
        'admission': txt.get('admission', ''), 'degree': txt.get('degree', ''), 'gender': gender,
        'capFirst': capFirst, 'capSecond': capSecond, 'capFemale': capF, 'capMale': capM,
        'notes': notes, 'star': bool(block and block.get('star')), 'page': str(pn),
        'blockRaw': block['raw'] if block else '', 'section': section,
    }
    if native:
        rec['src'] = 'native'; rec['nativeProvince'] = block['nativeProvince']
        if domain: rec['domain'] = domain
        if service: rec['service'] = service
    if not rec['major']: problems.append(f'p{pn} code {code}: empty major')
    return rec


# =====================================================================
# جدول «نوع گزینش» (ابتدای دفترچه): کد رشته | نوع گزینش (دوره‌های روزانه) | نام رشته | مقطع | ترتیب
# =====================================================================
def band_text(page, band):
    cs = [c for c in page.chars if band['top'] - 1 <= (c['top'] + c['bottom']) / 2 <= band['bottom'] + 1]
    return chars_to_text(cs)

def sel_field_of(label):
    t = re.sub(r'\s+', '', label)
    if 'گزینش' in t: return 'selection'
    if 'مقطع' in t: return 'level'
    if 'ترتیب' in t: return 'order'
    if 'نام' in t: return 'name'
    if t.startswith('کد') or 'ردیف' in t: return 'code'
    return None

def parse_selection_page(page, band, chars, ys, problems):
    xs = sorted({round(r['x0']) for r in band['rs']} | {round(r['x1']) for r in band['rs']})
    edges = []
    for x in xs:
        if not edges or x - edges[-1] > 2: edges.append(x)
    cols = {}
    for a, b in zip(edges, edges[1:]):
        if b - a < 8: continue
        cs = [c for c in page.chars if band['top'] - 1 <= (c['top'] + c['bottom']) / 2 <= band['bottom'] + 1
              and a <= (c['x0'] + c['x1']) / 2 < b]
        f = sel_field_of(chars_to_text(cs))
        if f and f not in cols: cols[f] = (a - 1, b + 1)
    if not {'selection', 'name', 'order'} <= set(cols):
        problems.append(f'p{page.number}: selection-table header not recognised: {cols}')
        return []
    left, right = min(a for a, b in cols.values()), max(b for a, b in cols.values())
    body = [c for c in chars if c['top'] > band['bottom'] + 1 and c['bottom'] < page.height - 70
            and left <= (c['x0'] + c['x1']) / 2 <= right]
    oa, ob = cols['order']
    ordch = [c for c in body if oa <= (c['x0'] + c['x1']) / 2 < ob]
    lines = []
    for c in sorted(ordch, key=lambda c: c['top']):
        if lines and abs(lines[-1]['top'] - c['top']) < 3: lines[-1]['cs'].append(c)
        else: lines.append({'top': c['top'], 'cs': [c]})
    centers = []
    for ln in lines:
        t = chars_to_text(ln['cs'])
        if re.fullmatch(r'\d{1,3}', t):
            centers.append((int(t), sum((c['top'] + c['bottom']) / 2 for c in ln['cs']) / len(ln['cs'])))
    rows = []
    for i, (n, y) in enumerate(centers):
        lo = (centers[i - 1][1] + y) / 2 if i > 0 else y - 12
        hi = (y + centers[i + 1][1]) / 2 if i + 1 < len(centers) else y + 12
        row = {'order': n, 'page': page.number}
        for f, (a, b) in cols.items():
            if f == 'order': continue
            row[f] = chars_to_text([c for c in body if lo <= (c['top'] + c['bottom']) / 2 < hi and a <= (c['x0'] + c['x1']) / 2 < b])
        row['selection'] = row.get('selection', '').replace('ناحیهای', 'ناحیه‌ای')    # ZWNJ در لایه‌ی متنی نیست
        rows.append(row)
    return rows

def sel_key(s):
    s = s.replace('ي', 'ی').replace('ك', 'ک').replace('ة', 'ه')
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u064b-\u065f\u0670\u200c]', '', s)
    s = re.sub(r'\([^)]*\)', '', s)
    s = re.sub(r'[\s\-–—.]+', '', s)
    return s.replace('دکترای', 'دکتری')

def canonicalize_provinces(records):
    """املای نام استان در بخش‌های مختلف دفترچه یکسان نیست (مثلاً «چهار محال» / «چهارمحال»)؛ پرتکرارترین املا انتخاب می‌شود."""
    for field in ('province', 'nativeProvince'):
        cnt = collections.defaultdict(collections.Counter)
        for r in records:
            v = r.get(field, '')
            if v: cnt[re.sub(r'[\s\u200c]+', '', v)][v] += 1
        best = {k: c.most_common(1)[0][0] for k, c in cnt.items()}
        changed = 0
        for r in records:
            v = r.get(field, '')
            if v:
                n = best[re.sub(r'[\s\u200c]+', '', v)]
                if n != v: r[field] = n; changed += 1
    return

def apply_selection(records, sel_rows, problems):
    """نوع گزینش فقط برای «دوره‌های روزانه» تعریف شده؛ فقط ردیف‌های با دوره‌ی «روزانه» و نام رشته‌ی دقیقاً مطابق جدول مقدار می‌گیرند."""
    orders = [r['order'] for r in sel_rows]
    if sel_rows and orders != list(range(1, len(orders) + 1)):
        problems.append(f'selection table: order column is not 1..{len(orders)} (got {orders[:5]}...{orders[-3:]})')
    tab = {}
    for r in sel_rows:
        k = sel_key(r['name'])
        if k in tab and tab[k] != r['selection']:
            problems.append(f'selection table: conflicting types for {r["name"]!r}')
        tab[k] = r['selection']
    n = 0
    for rec in records:
        if rec['degree'] == 'روزانه':
            t = tab.get(sel_key(rec['major']))
            if t:
                rec['selection'] = t; n += 1
    return n

# =====================================================================
# اعتبارسنجی مستقل با poppler (pdftotext) — یک موتور دوم، با ترتیب منطقی متن
# =====================================================================
def pdftotext_page(pdf_path, p):
    return subprocess.run(['pdftotext', '-f', str(p), '-l', str(p), pdf_path, '-'], capture_output=True).stdout.decode('utf8', 'ignore')

def _n(s, ours=False):
    # poppler گلیف لیگاتور «لا» را وارونه («ال») بیرون می‌دهد؛ فقط سمت خودمان را برای مقایسه وارونه می‌کنیم.
    if ours: s = s.replace('\u0644\u0627', '\u0627\u0644')
    s = s.replace('\u064a', '\u06cc').replace('\u0643', '\u06a9')
    s = re.sub(r'[\u200b-\u200f\u202a-\u202e\u2066-\u2069\u200c\u064b-\u065f\u0670]', '', s)
    s = s.translate(DIGIT_MAP)
    # فاصله/خط‌تیره/ستاره/پرانتز/گیومه/دونقطه/اسلش نادیده؛ جای آن‌ها در دو موتور bidi فرق می‌کند
    return re.sub(r'[\s\-–—*()\[\]«»:،/.]+', '', s)

def validate_with_poppler(pdf_path, records, pages, progress=None):
    lines, ok_all = [], True
    ptxt = {}
    for i, p in enumerate(pages):
        ptxt[p] = pdftotext_page(pdf_path, p)
        if progress and (i + 1) % 100 == 0: progress(f'poppler {i + 1}/{len(pages)}')
    pn = {p: _n(t) for p, t in ptxt.items()}
    by_page = collections.defaultdict(list)
    for r in records: by_page[int(r['page'])].append(r)
    bad = []                                                      # (1) مجموعه‌ی کدهای هر صفحه
    for p in pages:
        cand = set(re.findall(r'(?<!\d)\d{5}(?!\d)', ptxt[p].translate(DIGIT_MAP)))
        ours = {r['code'] for r in by_page.get(p, [])}
        if cand != ours: bad.append((p, sorted(cand - ours)[:4], sorted(ours - cand)[:4]))
    lines.append(f'[{"PASS" if not bad else "FAIL"}] code set per page vs poppler: {len(bad)} mismatching pages {bad[:5]}')
    ok_all &= not bad
    miss = []                                                      # (2) هر خطِ عنوان رشته/توضیحات
    for r in records:
        for key in ('_majorLines', '_notesLines'):
            for ln in r.get(key, []):
                v, v2 = _n(ln, ours=True), _n(ln)
                if v and v not in pn[int(r['page'])] and v2 not in pn[int(r['page'])]:
                    miss.append((r['page'], r['code'], key, ln[:60]))
    lines.append(f'[{"PASS" if not miss else "REVIEW"}] major/notes lines vs poppler text: {len(miss)} mismatches')
    for m in miss[:15]: lines.append('      ' + repr(m))
    # ناهمخوانی‌های معدود معمولاً «کلمه‌ی جداافتاده در دو خط» است؛ آستانه‌ی خطا: بیش از ۰٫۱٪ ردیف‌ها
    ok_all &= len(miss) <= max(3, len(records) // 1000)
    seen, tb = set(), []                                           # (3) عنوان بلوک‌ها
    for r in records:
        k = (r['page'], r['blockRaw'])
        if k in seen: continue
        seen.add(k)
        if _n(r['blockRaw'], True) not in pn[int(r['page'])] and _n(r['blockRaw']) not in pn[int(r['page'])]: tb.append(k)
    lines.append(f'[{"PASS" if not tb else "FAIL"}] block titles vs poppler: {len(tb)} mismatches {tb[:5]}')
    ok_all &= not tb
    c = collections.Counter(r['code'] for r in records)            # (4) یکتایی کد
    dups = {k: v for k, v in c.items() if v > 1}
    lines.append(f'[{"PASS" if not dups else "FAIL"}] duplicate codes: {len(dups)} {list(dups)[:5]}')
    ok_all &= not dups
    return ok_all, lines

# =====================================================================
# خروجی
# =====================================================================
CSV_HEAD = ['کد رشته', 'عنوان رشته', 'دانشگاه', 'استان', 'نحوه پذیرش', 'دوره تحصیلی', 'جنس پذیرش', 'ظرفیت اول', 'ظرفیت دوم',
            'ظرفیت زن', 'ظرفیت مرد', 'توضیحات', 'ستاره‌دار (مصاحبه/گزینش)', 'صفحه دفترچه', 'نوع ردیف',
            'بخش دفترچه', 'نوع گزینش (روزانه)', 'استان بومی', 'دامنه پذیرش', 'محل خدمت']
KEYS = ['code', 'major', 'university', 'province', 'admission', 'degree', 'gender', 'capFirst', 'capSecond', 'capFemale',
        'capMale', 'notes', 'star', 'page', 'blockRaw', 'section', 'selection', 'src', 'nativeProvince', 'domain', 'service']

def write_outputs(records, sel_rows, outdir, title='', year=None, group=''):
    bank = []
    for r in records:
        o = {k: r[k] for k in KEYS if k in r and r[k] not in ('', None)}
        for k in ('code', 'major', 'university', 'province', 'admission', 'degree', 'gender', 'capFirst', 'capSecond',
                  'capFemale', 'capMale', 'notes', 'page'):
            o.setdefault(k, '')
        o['star'] = bool(r.get('star'))
        bank.append(o)
    sections = collections.Counter(r['section'] for r in records)
    meta = {'title': title or 'بانک استخراج‌شده', 'group': group, 'year': year, 'rows': len(bank),
            'source': 'extract_booklet.py', 'sections': dict(sections),
            'selectionTable': [{'order': s['order'], 'name': s['name'], 'selection': s['selection'], 'level': s.get('level', '')} for s in sel_rows]}
    data = {'format': 'entekhab-reshte-bank', 'version': 1, 'meta': meta, 'bank': bank}
    with open(os.path.join(outdir, 'bank.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    with open(os.path.join(outdir, 'bank.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(CSV_HEAD)
        for r in records:
            w.writerow([r['code'], r['major'], r['university'], r['province'], r['admission'], r['degree'], r['gender'],
                        r['capFirst'], r['capSecond'], r['capFemale'], r['capMale'], r['notes'], 'بله' if r.get('star') else '',
                        r['page'], 'بومی' if r.get('src') == 'native' else 'اصلی', r.get('section', ''), r.get('selection', ''),
                        r.get('nativeProvince', ''), r.get('domain', ''), r.get('service', '')])

def summary_lines(records):
    L = [f'records={len(records)}  pages={len({r["page"] for r in records})}  provinces={len({r["province"] for r in records if r["province"]})}  blocks={len({r["blockRaw"] for r in records})}']
    for sec, n in collections.Counter(r['section'] for r in records).most_common():
        L.append(f'  section: {sec or "(none)"}: {n}')
    for k in ('admission', 'degree', 'gender'):
        L.append(f'{k}: ' + str(collections.Counter(r[k] for r in records).most_common(8)))
    L.append('code length: ' + str(collections.Counter(len(r['code']) for r in records)))
    return L

# =====================================================================
# --calibrate : کمک برای دفترچه‌ی جدید
# =====================================================================
def calibrate(src, page_no):
    page = src.get(page_no)
    print(f'page {page_no}: chars={len(page.chars)} thin-rules={len(page.thin)} header-cells={len(page.fills)}')
    bands = header_bands(page)
    print(f'\n-- {len(bands)} header band(s) (خانه‌های رنگیِ هدر جدول)')
    for b in bands:
        probs = []
        lay, err = build_layout(page, b, probs)
        print(f'  band y={b["top"]:.0f}-{b["bottom"]:.0f}:', 'layout =', lay.sig if lay else f'NOT RECOGNISED ({err})')
        if lay:
            for f, a, c in lay.cols: print(f'      {f:10} x={a}-{c}   header: {lay.labels.get(f, "")!r}')
        else:
            print('      text:', band_text(page, b))
    print('\n-- تیترهای بلوک (قرمز تیره، فونت عنوان):')
    for t in title_lines(page): print('   ', t['text'], '=>', {k: v for k, v in parse_title(t['text']).items() if v and k != 'raw'})
    print('\n-- سرتیترهای بزرگ:', heading_lines(page))
    fonts = collections.Counter((c['font'], c['size'], c['color']) for c in page.chars)
    print('\n-- فونت/اندازه/رنگ (برای تنظیم TITLE_FONT / TITLE_SIZE / TITLE_COLOR):')
    for k, v in fonts.most_common(8): print('   ', k, v)

def parse_range(s, total):
    if not s: return list(range(1, total + 1))
    a, _, b = s.partition('-')
    return list(range(int(a), int(b or a) + 1))

def main():
    ap = argparse.ArgumentParser(description='استخراج بانک رشته‌محل‌ها از PDF دفترچه انتخاب رشته')
    ap.add_argument('pdf')
    ap.add_argument('--out', default='out')
    ap.add_argument('--calibrate', type=int, metavar='PAGE', help='نمایش ساختار ستون‌ها و تیترهای یک صفحه')
    ap.add_argument('--pages', default='', help='محدوده‌ی صفحه‌ها، مثل 1-120 (پیش‌فرض: همه)')
    ap.add_argument('--cache', default='', help='پوشه‌ی کش صفحه‌ها (پارس مجدد را حذف می‌کند و کار را قابل ادامه می‌کند)')
    ap.add_argument('--skip-poppler', action='store_true')
    ap.add_argument('--title', default=''); ap.add_argument('--year', type=int, default=None); ap.add_argument('--group', default='')
    a = ap.parse_args()
    src = PageSource(a.pdf, a.cache or None)
    if a.calibrate:
        calibrate(src, a.calibrate); return
    os.makedirs(a.out, exist_ok=True)
    total = src.count() if not a.cache or not os.path.exists(os.path.join(a.cache, 'p0001.pkl')) else None
    if total is None:
        total = len([f for f in os.listdir(a.cache) if f.endswith('.pkl')])
    pages = parse_range(a.pages, total)
    t0 = time.time()
    def gen():
        for n in pages:
            if n % 25 == 0: print(f'  page {n}/{pages[-1]}  ({time.time() - t0:.0f}s)', file=sys.stderr, flush=True)
            yield src.get(n)
    records, problems, stats, sel_rows = extract(gen())
    canonicalize_provinces(records)
    n_sel = apply_selection(records, sel_rows, problems) if sel_rows else 0
    data_pages = [p for p, n in stats if n > 0]
    rep = [f'PDF: {a.pdf}', f'PROBLEMS = {len(problems)}'] + ['  ' + p[:220] for p in problems[:80]]
    rep += summary_lines(records)
    rep.append(f'selection table: {len(sel_rows)} majors ({dict(collections.Counter(s["selection"] for s in sel_rows))}); '
               f'matched daily rows: {n_sel}')
    ok = not problems
    if not a.skip_poppler:
        try:
            ok2, lines = validate_with_poppler(a.pdf, records, data_pages, progress=lambda m: print(' ', m, file=sys.stderr, flush=True))
            rep += lines; ok &= ok2
        except FileNotFoundError:
            rep.append('[SKIP] pdftotext (poppler-utils) not installed — independent validation skipped'); ok = False
    write_outputs(records, sel_rows, a.out, a.title, a.year, a.group)
    rep.append('RESULT: ' + ('ALL CHECKS PASSED' if ok else 'CHECKS FAILED — do not use the data before fixing'))
    open(os.path.join(a.out, 'validation_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
    print('\n'.join(rep))

if __name__ == '__main__':
    main()
