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
    ('Sans',     os.path.join(_REPORTLAB_FONTS, 'Vera.ttf')),
    ('SansBold', os.path.join(_REPORTLAB_FONTS, 'VeraBd.ttf')),
    ('SansObl',  os.path.join(_REPORTLAB_FONTS, 'VeraIt.ttf')),
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
    name = get('Commercial Name')
    product = {
        'ref':      get('Product Reference'),
        'brand':    get('Brand'),
        'name':     name,
        'claim':    get('Key claim'),
        'category': get('PL') or get('Family L1') or 'GENEL',
        'series':   get('Family L2') or get('Range name') or get('Range') or get('Series') or ' '.join(name.split()[:2]),
        'benefits': [],
    }
    seen = set()
    for i in ['1 (USP)','2','3','4','5','6','7','8']:
        t = get(f'Benefit title {i}')
        d = get(f'Benefit detail {i}')
        if t and t.lower() not in ('none','') and t not in seen:
            seen.add(t); product['benefits'].append((t,d))
    product['images_b64'] = extract_images_b64(buf)
    return product

def extract_images_b64(xlsm_buf):
    imgs = []
    xlsm_buf.seek(0)
    try:
        with zipfile.ZipFile(xlsm_buf,'r') as z:
            for path in [f for f in z.namelist() if f.startswith('xl/media/')]:
                data = z.read(path)
                try:
                    im = Image.open(BytesIO(data))
                    if min(im.size) >= 150:
                        if im.mode == 'P': im = im.convert('RGBA')
                        if im.mode == 'RGBA':
                            bg = Image.new('RGB', im.size, (255,255,255))
                            bg.paste(im, mask=im.split()[3]); im = bg
                        else: im = im.convert('RGB')
                        out = BytesIO()
                        im.save(out, 'JPEG', quality=85)
                        imgs.append(base64.b64encode(out.getvalue()).decode())
                        break
                except: pass
    except: pass
    return imgs

def b64_to_reader(b64):
    try:
        return ImageReader(BytesIO(base64.b64decode(b64)))
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

def draw_lifestyle_banner(cv, x, y_top, w, h, lifestyle_b64, category):
    """Full-width lifestyle image banner on page 1. y_top = top edge in PDF coords."""
    img_reader = b64_to_reader(lifestyle_b64)

    # Background fill in case image doesn't cover fully
    cv.setFillColor(LGRAY)
    cv.rect(x, y_top - h, w, h, fill=1, stroke=0)

    if img_reader:
        try:
            cv.drawImage(img_reader, x, y_top - h, w, h,
                         preserveAspectRatio=True, anchor='c', mask='auto')
        except: pass

    # Dark text bar at bottom of banner (brand consistent with header)
    bar_h = 10*mm
    cv.setFillColor(DARK)
    cv.rect(x, y_top - h, w, bar_h, fill=1, stroke=0)
    # Red left accent stripe
    cv.setFillColor(RED)
    cv.rect(x, y_top - h, 3.5*mm, bar_h, fill=1, stroke=0)
    # Category label in white
    cv.setFillColor(WHITE); cv.setFont(FB(), 9)
    cv.drawString(x + 6.5*mm, y_top - h + 3.5*mm, category.upper())

    # Separator line at top of banner
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3)
    cv.line(x, y_top, x + w, y_top)

def draw_card(cv, x, y, cw, ch, product):
    """3-column card: white image, bold name, tightly packed bullet features."""
    imgs = [b64_to_reader(b) for b in product.get('images_b64', [])]
    imgs = [i for i in imgs if i]

    IMG_H = cw * 0.76

    # Image area
    img_bot = y - IMG_H
    cv.setFillColor(WHITE)
    cv.rect(x, img_bot, cw, IMG_H, fill=1, stroke=0)
    if imgs:
        try:
            cv.drawImage(imgs[0], x, img_bot, cw, IMG_H,
                         preserveAspectRatio=True, anchor='c', mask='auto')
        except: pass

    # Red accent line under image
    cv.setStrokeColor(RED); cv.setLineWidth(1.5)
    cv.line(x, img_bot, x + cw * 0.38, img_bot)
    cv.setStrokeColor(colors.HexColor('#E8E8E8')); cv.setLineWidth(0.4)
    cv.line(x + cw * 0.38 + 1.5*mm, img_bot, x + cw, img_bot)

    # Product name — tight gap, compact line spacing
    ty = img_bot - 2.5*mm
    cv.setFillColor(DARK); cv.setFont(FB(), 8)
    name_lines = wrap(cv, product.get('name', ''), FB(), 8, cw)
    for i, ln in enumerate(name_lines[:2]):
        if i == 1 and len(name_lines) > 2:
            ln = ln.rstrip('.,;: ')
        cv.drawString(x, ty, ln)
        ty -= 4*mm

    # Ref code (no extra pre-gap)
    cv.setFillColor(DGRAY); cv.setFont(F(), 6)
    cv.drawString(x, ty, product.get('ref', ''))
    ty -= 3.5*mm

    # Separator
    cv.setStrokeColor(colors.HexColor('#DEDEDE')); cv.setLineWidth(0.3)
    cv.line(x, ty, x + cw, ty)
    ty -= 2.5*mm

    # Bullet features — tight spacing so they start close to separator
    BUL_X = x + 4*mm
    BUL_W = cw - 4.5*mm
    bottom_limit = y - ch + 3*mm

    for title, _ in product.get('benefits', [])[:6]:
        if ty - 3.5*mm < bottom_limit: break
        text = str(title)
        while tw(cv, text, F(), 6.5) > BUL_W and len(text) > 5:
            text = text[:-2] + '.'
        cv.setFillColor(RED)
        cv.circle(x + 1.8*mm, ty + 0.7*mm, 1.1*mm, fill=1, stroke=0)
        cv.setFillColor(DARK); cv.setFont(F(), 6.5)
        cv.drawString(BUL_X, ty, text)
        ty -= 3.5*mm

def build_pdf(products, output_path, category, lifestyle_image=None):
    # Group same-series products together so they appear side by side
    products = sorted(products, key=lambda p: (p.get('series', ''), p.get('name', '')))

    load_fonts()
    cv = canvas.Canvas(output_path, pagesize=A4)
    cv.setTitle(f'TEFAL {category} Katalogu 2024')

    MARGIN   = 12*mm
    HDR_H    = 13*mm
    FTR_H    = 10*mm
    COLS     = 3
    COL_GAP  = 6*mm
    ROW_GAP  = 8*mm
    BANNER_H = 52*mm

    card_w = (W - 2*MARGIN - (COLS - 1)*COL_GAP) / COLS
    card_h = 95*mm

    usable_h = H - HDR_H - FTR_H - 2*MARGIN
    rows_pp  = max(1, int((usable_h + ROW_GAP) / (card_h + ROW_GAP)))
    per_page = COLS * rows_pp

    # Page 1 has fewer product rows when lifestyle banner is present
    if lifestyle_image:
        usable_h_p1 = usable_h - BANNER_H - ROW_GAP
        rows_pp1    = max(1, int((usable_h_p1 + ROW_GAP) / (card_h + ROW_GAP)))
        per_page_1  = COLS * rows_pp1
    else:
        per_page_1 = per_page

    page_num = 1
    draw_page_chrome(cv, page_num, category)

    # Lifestyle banner on page 1
    banner_top = H - HDR_H - MARGIN
    if lifestyle_image:
        draw_lifestyle_banner(cv, MARGIN, banner_top, W - 2*MARGIN, BANNER_H, lifestyle_image, category)
        product_start_y = banner_top - BANNER_H - ROW_GAP
    else:
        product_start_y = banner_top

    page_i = 0
    current_start_y   = product_start_y
    current_per_page  = per_page_1

    for product in products:
        if page_i >= current_per_page:
            cv.showPage(); page_num += 1
            draw_page_chrome(cv, page_num, category)
            page_i = 0
            current_start_y  = banner_top
            current_per_page = per_page

        col = page_i % COLS
        row = page_i // COLS
        x = MARGIN + col * (card_w + COL_GAP)
        y = current_start_y - row * (card_h + ROW_GAP)
        draw_card(cv, x, y, card_w, card_h, product)
        page_i += 1

    cv.showPage()
    cv.save()

# ─── In-memory catalog state ───────────────────────────────────────────────────
CATALOG_STATE = defaultdict(dict)

# ─── Claude Vision matching ────────────────────────────────────────────────────

def match_with_claude(img_b64, media_type, categories):
    """Use Claude Haiku Vision to pick the best-matching category for a lifestyle image."""
    try:
        import anthropic
        client = anthropic.Anthropic()

        if categories:
            cat_list = ', '.join(sorted(categories))
            prompt = (
                f"Bu yasam tarzi mutfak gorseli hangi urun kategorisine en uygun? "
                f"Kategoriler: {cat_list}. "
                f"Sadece kategori adini yaz, baska hicbir sey yazma."
            )
        else:
            prompt = "Bu gorseli 1-3 kelimeyle tanimla (mutfak urunu kategorisi)."

        message = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=50,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )

        result = message.content[0].text.strip()
        for cat in sorted(categories, key=len, reverse=True):
            if cat.lower() in result.lower() or result.lower() in cat.lower():
                return cat
        return list(categories)[0] if categories else result

    except Exception as e:
        print(f"Claude match error: {e}")
        return list(categories)[0] if categories else 'GENEL'

# ─── Endpoints ─────────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    load_fonts()
    summary = {cat: len(prods) for cat, prods in CATALOG_STATE.items()}
    return jsonify({'status': 'ok', 'fonts': list(FONT_MAP.keys()), 'state': summary})

@app.route('/match_lifestyle', methods=['POST'])
def match_lifestyle():
    """Receive a lifestyle image, match it to a category using Claude Vision."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'file eksik'}), 400

        categories = set(json.loads(request.form.get('categories_json', '[]')))

        img_data = request.files['file'].read()

        # Resize to max 1200px wide before sending to Claude
        try:
            im = Image.open(BytesIO(img_data))
            if im.mode != 'RGB':
                if im.mode == 'RGBA':
                    bg = Image.new('RGB', im.size, (255, 255, 255))
                    bg.paste(im, mask=im.split()[3]); im = bg
                else:
                    im = im.convert('RGB')
            if im.width > 1200:
                ratio = 1200 / im.width
                im = im.resize((1200, int(im.height * ratio)), Image.LANCZOS)
            out = BytesIO()
            im.save(out, 'JPEG', quality=85)
            img_data = out.getvalue()
        except Exception as e:
            print(f"Lifestyle image resize error: {e}")

        img_b64 = base64.b64encode(img_data).decode()
        matched = match_with_claude(img_b64, 'image/jpeg', categories)

        return jsonify({'category': matched, 'image_b64': img_b64})

    except Exception as e:
        import traceback; print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/finalize', methods=['POST'])
def finalize():
    """Rebuild PDF for a category using provided state JSON + optional lifestyle image."""
    try:
        category    = request.form.get('category')
        state_json  = request.form.get('state_json')
        lifestyle_b64 = request.form.get('lifestyle_b64') or None

        if not category or not state_json:
            return jsonify({'error': 'category ve state_json gerekli'}), 400

        state = json.loads(state_json)
        products_list = list(state.get(category, {}).values())

        if not products_list:
            return jsonify({'error': f'{category} icin urun yok'}), 400

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            out_path = tmp.name

        build_pdf(products_list, out_path, category, lifestyle_image=lifestyle_b64)

        with open(out_path, 'rb') as fh:
            pdf_data = fh.read()
        os.unlink(out_path)

        safe_cat = category.replace(' ', '_').replace('/', '_').replace('&', 'and')
        return jsonify({
            'pdf_base64':    base64.b64encode(pdf_data).decode(),
            'filename':      f'katalog_{safe_cat}.pdf',
            'category':      category,
            'product_count': len(products_list),
        })

    except Exception as e:
        import traceback; print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/add_product', methods=['POST'])
@app.route('/upload', methods=['POST'])
def add_product():
    try:
        load_fonts()
        if 'file' not in request.files:
            return jsonify({'error': 'file eksik'}), 400

        # Restore product state from previous call (Railway restart resilience)
        if 'prev_state_json' in request.form:
            try:
                incoming = json.loads(request.form['prev_state_json'])
                for cat, prods in incoming.items():
                    for ref, prod in prods.items():
                        if ref not in CATALOG_STATE[cat]:
                            CATALOG_STATE[cat][ref] = prod
            except Exception:
                pass

        # Restore lifestyle images map
        lifestyle_map = {}
        if 'lifestyle_images_json' in request.form:
            try:
                lifestyle_map = json.loads(request.form['lifestyle_images_json'])
            except Exception:
                pass

        xlsm_bytes = request.files['file'].read()
        product    = parse_xlsm(xlsm_bytes)
        category   = product['category'] or 'GENEL'

        CATALOG_STATE[category][product['ref']] = product
        products_list = list(CATALOG_STATE[category].values())

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            out_path = tmp.name

        build_pdf(products_list, out_path, category,
                  lifestyle_image=lifestyle_map.get(category))

        with open(out_path, 'rb') as fh:
            pdf_data = fh.read()
        os.unlink(out_path)

        safe_cat = category.replace(' ', '_').replace('/', '_').replace('&', 'and')

        state_snapshot = {cat: dict(prods) for cat, prods in CATALOG_STATE.items()}

        return jsonify({
            'pdf_base64':    base64.b64encode(pdf_data).decode(),
            'filename':      f'katalog_{safe_cat}.pdf',
            'category':      category,
            'product_ref':   product['ref'],
            'product_count': len(products_list),
            'state_json':    state_snapshot,
        })

    except Exception as e:
        import traceback; print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)


@app.route('/find_fonts', methods=['GET'])
def find_fonts():
    import glob
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
