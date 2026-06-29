from flask import Flask, request, jsonify
import base64, zipfile, os, tempfile, json
from io import BytesIO
from collections import defaultdict

from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image
import reportlab

app = Flask(__name__)

FONT_MAP = {}
fonts_loaded = False

_REPORTLAB_FONTS = os.path.join(os.path.dirname(reportlab.__file__), 'fonts')
_FONT_CANDIDATES = [
    # Vera — ReportLab built-in, Türkçe karakter desteği var, her ortamda çalışır
    ('Sans',     os.path.join(_REPORTLAB_FONTS, 'Vera.ttf')),
    ('SansBold', os.path.join(_REPORTLAB_FONTS, 'VeraBd.ttf')),
    ('SansObl',  os.path.join(_REPORTLAB_FONTS, 'VeraIt.ttf')),
    # DejaVu fallback
    ('Sans',     '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),
    ('SansBold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
    ('SansObl',  '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf'),
]

def load_fonts():
    global fonts_loaded, FONT_MAP
    if fonts_loaded: return
    for name, path in _FONT_CANDIDATES:
        if name not in FONT_MAP and os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                FONT_MAP[name] = path
            except: pass
    fonts_loaded = True

def F():  return 'Sans'     if 'Sans'     in FONT_MAP else 'Helvetica'
def FB(): return 'SansBold' if 'SansBold' in FONT_MAP else 'Helvetica-Bold'
def FO(): return 'SansObl'  if 'SansObl'  in FONT_MAP else 'Helvetica-Oblique'

W, H  = A4
RED   = colors.HexColor('#E8281A')
DARK  = colors.HexColor('#1C1C1C')
LGRAY = colors.HexColor('#F5F5F5')
MGRAY = colors.HexColor('#CCCCCC')
DGRAY = colors.HexColor('#555555')
WHITE = colors.white

def tw(cv, t, f, s): return cv.stringWidth(t, f, s)

def wrap(cv, text, font, size, max_w):
    words = str(text).split(); lines, line = [], ''
    for w in words:
        t = (line+' '+w).strip()
        if tw(cv,t,font,size) <= max_w: line = t
        else:
            if line: lines.append(line)
            line = w
    if line: lines.append(line)
    return lines

def parse_xlsm(xlsm_bytes):
    buf = BytesIO(xlsm_bytes)
    wb  = load_workbook(buf, data_only=True)
    rt  = wb['RANGE TABLE']
    def get(label):
        for row in rt.iter_rows(values_only=True):
            if row[0] and str(row[0]).strip() == label:
                return str(row[1]).strip() if row[1] else ''
        return ''
    product = {
        'ref':      get('Product Reference'),
        'brand':    get('Brand'),
        'name':     get('Commercial Name'),
        'claim':    get('Key claim'),
        'category': get('PL') or get('Family L1') or 'GENEL',
        'benefits': [],
    }
    seen = set()
    for i in ['1 (USP)','2','3','4','5','6','7','8']:
        t = get(f'Benefit title {i}')
        d = get(f'Benefit detail {i}')
        if t and t.lower() not in ('none','') and t not in seen:
            seen.add(t); product['benefits'].append((t,d))
    # Görseli hızlı çek — upscale yok
    product['images_b64'] = extract_images_b64(buf)
    return product

def extract_images_b64(xlsm_buf):
    """Sadece ilk uygun görseli al, işleme yapma"""
    imgs = []
    xlsm_buf.seek(0)
    try:
        with zipfile.ZipFile(xlsm_buf,'r') as z:
            for path in [f for f in z.namelist() if f.startswith('xl/media/')]:
                data = z.read(path)
                try:
                    im = Image.open(BytesIO(data))
                    if min(im.size) >= 150:
                        # Sadece RGB'ye çevir, boyut değiştirme
                        if im.mode == 'P': im = im.convert('RGBA')
                        if im.mode == 'RGBA':
                            bg = Image.new('RGB', im.size, (255,255,255))
                            bg.paste(im, mask=im.split()[3]); im = bg
                        else: im = im.convert('RGB')
                        out = BytesIO()
                        im.save(out, 'JPEG', quality=85)
                        imgs.append(base64.b64encode(out.getvalue()).decode())
                        break  # sadece ilk görsel
                except: pass
    except: pass
    return imgs

def b64_to_reader(b64):
    try:
        out = BytesIO(base64.b64decode(b64))
        return ImageReader(out)
    except: return None

def draw_page_chrome(cv, page_num, category):
    HDR = 13*mm
    cv.setFillColor(DARK); cv.rect(0, H-HDR, W, HDR, fill=1, stroke=0)
    cv.setFillColor(RED);  cv.rect(0, H-HDR, 3.5*mm, HDR, fill=1, stroke=0)
    cv.setFillColor(WHITE); cv.setFont(FB(), 9)
    cv.drawString(8*mm, H-HDR+4.5*mm, str(category).upper())
    cv.setFont(F(), 7.5); cv.setFillColor(MGRAY)
    cv.drawRightString(W-8*mm, H-HDR+4.5*mm, 'Urun Katalogu 2024')
    FTR = 10*mm
    cv.setFillColor(LGRAY); cv.rect(0, 0, W, FTR, fill=1, stroke=0)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3); cv.line(0, FTR, W, FTR)
    cv.setFillColor(DGRAY); cv.setFont(F(), 6.5)
    cv.drawString(8*mm, 3.5*mm, '2024 TEFAL')
    cv.drawRightString(W-8*mm, 3.5*mm, 'tefal.com.tr')
    cv.setFillColor(DARK); cv.setFont(FB(), 7)
    cv.drawCentredString(W/2, 3.5*mm, f'{page_num} | TEFAL')

def draw_card(cv, x, y, cw, ch, product):
    """3-column grid card: clean white image, bold name, wrapped bullet features."""
    imgs = [b64_to_reader(b) for b in product.get('images_b64', [])]
    imgs = [i for i in imgs if i]

    IMG_H = cw * 0.76

    # Image area — white, no gray box
    img_bot = y - IMG_H
    cv.setFillColor(WHITE)
    cv.rect(x, img_bot, cw, IMG_H, fill=1, stroke=0)
    if imgs:
        try:
            cv.drawImage(imgs[0], x, img_bot, cw, IMG_H,
                         preserveAspectRatio=True, anchor='c', mask='auto')
        except: pass

    # Thin red accent under image
    cv.setStrokeColor(RED); cv.setLineWidth(1.5)
    cv.line(x, img_bot, x + cw * 0.38, img_bot)
    cv.setStrokeColor(colors.HexColor('#E8E8E8')); cv.setLineWidth(0.4)
    cv.line(x + cw * 0.38 + 1.5*mm, img_bot, x + cw, img_bot)

    # Product name — strip trailing punctuation if truncated at 2 lines
    ty = img_bot - 4*mm
    cv.setFillColor(DARK); cv.setFont(FB(), 8)
    name_lines = wrap(cv, product.get('name', ''), FB(), 8, cw)
    for i, ln in enumerate(name_lines[:2]):
        if i == 1 and len(name_lines) > 2:
            ln = ln.rstrip('.,;: ')
        cv.drawString(x, ty, ln)
        ty -= 4.8*mm

    # Ref code
    ty -= 0.5*mm
    cv.setFillColor(DGRAY); cv.setFont(F(), 6)
    cv.drawString(x, ty, product.get('ref', ''))
    ty -= 4.5*mm

    # Separator
    cv.setStrokeColor(colors.HexColor('#DEDEDE')); cv.setLineWidth(0.3)
    cv.line(x, ty, x + cw, ty)
    ty -= 3.5*mm

    # Bullet features — single line, tight spacing
    BUL_X = x + 4*mm
    BUL_W = cw - 4.5*mm
    bottom_limit = y - ch + 3*mm

    for title, _ in product.get('benefits', [])[:6]:
        if ty - 3.8*mm < bottom_limit: break
        text = str(title)
        while tw(cv, text, F(), 6.5) > BUL_W and len(text) > 5:
            text = text[:-2] + '.'
        cv.setFillColor(RED)
        cv.circle(x + 1.8*mm, ty - 1.3*mm, 1.1*mm, fill=1, stroke=0)
        cv.setFillColor(DARK); cv.setFont(F(), 6.5)
        cv.drawString(BUL_X, ty, text)
        ty -= 3.8*mm

def build_pdf(products, output_path, category):
    load_fonts()
    cv = canvas.Canvas(output_path, pagesize=A4)
    cv.setTitle(f'TEFAL {category} Katalogu 2024')

    MARGIN  = 12*mm
    HDR_H   = 13*mm
    FTR_H   = 10*mm
    COLS    = 3
    COL_GAP = 6*mm
    ROW_GAP = 8*mm

    card_w = (W - 2*MARGIN - (COLS - 1)*COL_GAP) / COLS
    card_h = 95*mm

    usable_h = H - HDR_H - FTR_H - 2*MARGIN
    rows_pp  = max(1, int((usable_h + ROW_GAP) / (card_h + ROW_GAP)))
    per_page = COLS * rows_pp

    page_num = 1
    draw_page_chrome(cv, page_num, category)

    for pi in range(0, len(products), per_page):
        if pi > 0:
            cv.showPage(); page_num += 1
            draw_page_chrome(cv, page_num, category)
        for i, product in enumerate(products[pi:pi + per_page]):
            col = i % COLS
            row = i // COLS
            x = MARGIN + col * (card_w + COL_GAP)
            y = H - HDR_H - MARGIN - row * (card_h + ROW_GAP)
            draw_card(cv, x, y, card_w, card_h, product)

    cv.showPage()
    cv.save()

# In-memory state: {category: {ref: product}}
CATALOG_STATE = defaultdict(dict)

@app.route('/health', methods=['GET'])
def health():
    load_fonts()
    summary = {cat: len(prods) for cat, prods in CATALOG_STATE.items()}
    return jsonify({'status': 'ok', 'fonts': list(FONT_MAP.keys()), 'state': summary})

@app.route('/add_product', methods=['POST'])
@app.route('/upload', methods=['POST'])
def add_product():
    try:
        load_fonts()
        if 'file' not in request.files:
            return jsonify({'error': 'file eksik'}), 400

        # Pre-populate state from previous call so server restarts don't break accumulation
        if 'prev_state_json' in request.form:
            try:
                incoming = json.loads(request.form['prev_state_json'])
                for cat, prods in incoming.items():
                    for ref, prod in prods.items():
                        if ref not in CATALOG_STATE[cat]:
                            CATALOG_STATE[cat][ref] = prod
            except Exception:
                pass

        xlsm_bytes = request.files['file'].read()
        product = parse_xlsm(xlsm_bytes)
        category = product['category'] or 'GENEL'

        CATALOG_STATE[category][product['ref']] = product

        products_list = list(CATALOG_STATE[category].values())

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            out_path = tmp.name

        build_pdf(products_list, out_path, category)

        with open(out_path,'rb') as fh:
            pdf_data = fh.read()
        os.unlink(out_path)

        safe_cat = category.replace(' ','_').replace('/','_').replace('&','and')
        filename  = f'katalog_{safe_cat}.pdf'

        # Include images_b64 so watcher can pass full state back on next call
        state_snapshot = {
            cat: dict(prods)
            for cat, prods in CATALOG_STATE.items()
        }

        return jsonify({
            'pdf_base64':    base64.b64encode(pdf_data).decode(),
            'filename':      filename,
            'category':      category,
            'product_ref':   product['ref'],
            'product_count': len(products_list),
            'state_json':    state_snapshot,
        })

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)


@app.route('/find_fonts', methods=['GET'])
def find_fonts():
    import subprocess, glob
    # Farklı yerlerde ara
    results = []
    for pattern in [
        '/usr/share/fonts/**/*.ttf',
        '/nix/store/**/dejavu*/*.ttf',
        '/nix/store/**/*.ttf',
        '/home/**/*.ttf',
        '/opt/**/*.ttf',
    ]:
        found = glob.glob(pattern, recursive=True)
        results.extend(found[:5])
    return jsonify({'found': results[:30]})
