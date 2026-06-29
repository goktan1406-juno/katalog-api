from flask import Flask, request, jsonify, Response
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

app = Flask(__name__)

FONT_MAP = {}
fonts_loaded = False

# Vera fontları ReportLab ile birlikte gelir — her ortamda çalışır
import reportlab
REPORTLAB_FONTS = os.path.join(os.path.dirname(reportlab.__file__), 'fonts')

FONT_CANDIDATES = [
    # Vera (ReportLab built-in) — Türkçe karakter desteği var
    ('Sans',     os.path.join(REPORTLAB_FONTS, 'Vera.ttf')),
    ('SansBold', os.path.join(REPORTLAB_FONTS, 'VeraBd.ttf')),
    ('SansObl',  os.path.join(REPORTLAB_FONTS, 'VeraIt.ttf')),
    # DejaVu fallback
    ('Sans',     '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),
    ('SansBold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
    ('SansObl',  '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf'),
]

def load_fonts():
    global fonts_loaded, FONT_MAP
    if fonts_loaded: return
    for name, path in FONT_CANDIDATES:
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
    """Sadece ilk kaliteli görseli al — ana packshot"""
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
                        w,h = im.size; s = max(w,h)
                        sq = Image.new('RGB',(s,s),(255,255,255))
                        sq.paste(im,((s-w)//2,(s-h)//2))
                        out = BytesIO()
                        sq.save(out, 'JPEG', quality=88)
                        return [base64.b64encode(out.getvalue()).decode()]
                except: pass
    except: pass
    return []

def b64_to_reader(b64):
    try:
        out = BytesIO(base64.b64decode(b64))
        return ImageReader(out)
    except: return None

def calc_card_h(product):
    """Kart yüksekliği — 2-3 ürün sayfaya sığsın diye max 120mm"""
    n = len(product.get('benefits', []))
    # Her benefit: başlık(4.5mm) + detay(4.2mm) + gap(1.5mm) = 10.2mm
    h = 8*mm + 5*mm + 4.5*mm + 4*mm + n*10.2*mm + 10*mm
    return max(55*mm, min(h, 120*mm))

def draw_page_chrome(cv, page_num, category):
    HDR=13*mm
    cv.setFillColor(DARK); cv.rect(0,H-HDR,W,HDR,fill=1,stroke=0)
    cv.setFillColor(RED);  cv.rect(0,H-HDR,3.5*mm,HDR,fill=1,stroke=0)
    cv.setFillColor(WHITE); cv.setFont(FB(),9)
    cv.drawString(8*mm,H-HDR+4.5*mm,str(category).upper())
    cv.setFont(F(),7.5); cv.setFillColor(MGRAY)
    cv.drawRightString(W-8*mm,H-HDR+4.5*mm,'Urun Katalogu 2024')
    FTR=10*mm
    cv.setFillColor(LGRAY); cv.rect(0,0,W,FTR,fill=1,stroke=0)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3); cv.line(0,FTR,W,FTR)
    cv.setFillColor(DGRAY); cv.setFont(F(),6.5)
    cv.drawString(8*mm,3.5*mm,'2024 TEFAL')
    cv.drawRightString(W-8*mm,3.5*mm,'tefal.com.tr')
    cv.setFillColor(DARK); cv.setFont(FB(),7)
    cv.drawCentredString(W/2,3.5*mm,f'{page_num} | TEFAL')

def draw_card(cv, x, y, cw, ch, product):
    imgs = [b64_to_reader(b) for b in product.get('images_b64',[])]
    imgs = [i for i in imgs if i]
    n_imgs = len(imgs)

    PAD=4*mm; GAP=5*mm
    LW=(cw-2*PAD-GAP)*0.48 if n_imgs>0 else cw-2*PAD
    RW=(cw-2*PAD-GAP)*0.52 if n_imgs>0 else 0
    LX=x+PAD; RX=x+PAD+LW+GAP; BOT=y-ch

    cv.setFillColor(WHITE); cv.roundRect(x,BOT,cw,ch,2*mm,fill=1,stroke=0)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.5); cv.roundRect(x,BOT,cw,ch,2*mm,fill=0,stroke=1)
    cv.setFillColor(RED); cv.roundRect(x,y-5*mm,3*mm,5*mm,1*mm,fill=1,stroke=0)
    cv.rect(x+1.5*mm,y-5*mm,1.5*mm,5*mm,fill=1,stroke=0)

    cy=y-PAD
    cv.setFillColor(DARK); cv.setFont(FB(),9)
    for ln in wrap(cv,product['name'],FB(),9,LW)[:2]:
        cv.drawString(LX,cy,ln); cy-=5.5*mm
    cy-=1*mm

    cv.setFillColor(LGRAY)
    bw=tw(cv,product['ref'],F(),6.5)+5*mm
    cv.roundRect(LX,cy-5*mm,bw,5*mm,1.5*mm,fill=1,stroke=0)
    cv.setFillColor(DGRAY); cv.setFont(F(),6.5)
    cv.drawString(LX+2.5*mm,cy-3.5*mm,product['ref']); cy-=8*mm

    cv.setStrokeColor(RED); cv.setLineWidth(1.5); cv.line(LX,cy,LX+LW*0.45,cy)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3); cv.line(LX+LW*0.45+2*mm,cy,LX+LW,cy)
    cy-=4*mm

    if product.get('claim'):
        cv.setFillColor(RED); cv.setFont(FO(),7.5)
        claim=str(product['claim'])
        while tw(cv,claim,FO(),7.5)>LW and len(claim)>5: claim=claim[:-4]+'...'
        cv.drawString(LX,cy,claim); cy-=4.5*mm
        cv.setStrokeColor(colors.HexColor('#EBEBEB')); cv.setLineWidth(0.25)
        cv.line(LX,cy,LX+LW,cy); cy-=2.5*mm

    for title,detail in product.get('benefits',[]):
        if cy-4.5*mm<BOT+PAD+4*mm: break
        cv.setFillColor(RED); cv.setFont(FB(),9)
        cv.drawString(LX,cy-4.5*mm+1.5*mm,'*')
        cv.setFillColor(DARK); cv.setFont(FB(),7.5)
        t=str(title)
        while tw(cv,t,FB(),7.5)>LW-5*mm and len(t)>5: t=t[:-2]+'.'
        cv.drawString(LX+4.5*mm,cy-4.5*mm+1.5*mm,t); cy-=4.5*mm
        if detail:
            cv.setFillColor(DGRAY); cv.setFont(F(),7.0)
            for dl in wrap(cv,str(detail),F(),7.0,LW-4.5*mm)[:2]:
                if cy-4.2*mm<BOT+PAD+4*mm: break
                cv.drawString(LX+4.5*mm,cy-4.2*mm+1.5*mm,dl); cy-=4.2*mm
        cv.setStrokeColor(colors.HexColor('#F0F0F0')); cv.setLineWidth(0.2)
        cv.line(LX,cy-1*mm,LX+LW,cy-1*mm); cy-=1.5*mm

    cv.setFillColor(LGRAY); cv.roundRect(LX,BOT+2*mm,LW,6*mm,1*mm,fill=1,stroke=0)
    cv.setFillColor(DGRAY); cv.setFont(F(),6)
    cv.drawString(LX+3*mm,BOT+4.5*mm,
        f"Kutu Icerigi: {str(product['name']).split(',')[0]} - {product['ref']}")

    if n_imgs==0: return
    # Sadece ana fotoğraf — tam sağ sütun boyunca
    area_h = ch - 2*PAD
    px, py, pw, ph = RX, y-PAD-area_h, RW, area_h
    cv.setFillColor(LGRAY); cv.roundRect(px,py,pw,ph,2*mm,fill=1,stroke=0)
    try: cv.drawImage(imgs[0],px+2*mm,py+2*mm,pw-4*mm,ph-4*mm,
                      preserveAspectRatio=True,anchor='c',mask='auto')
    except: pass
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.4)
    cv.roundRect(px,py,pw,ph,2*mm,fill=0,stroke=1)

def build_pdf(products, output_path, category):
    load_fonts()
    cv = canvas.Canvas(output_path, pagesize=A4)
    cv.setTitle(f'TEFAL {category} Katalogu 2024')
    MARGIN=8*mm; HDR_H=13*mm; FTR_H=10*mm; CARD_GAP=4*mm; CARD_W=W-2*MARGIN
    page_num=1; current_y=H-HDR_H-5*mm
    draw_page_chrome(cv,page_num,category)
    for product in products:
        ch=calc_card_h(product)
        if current_y-ch<FTR_H+4*mm:
            cv.showPage(); page_num+=1; current_y=H-HDR_H-5*mm
            draw_page_chrome(cv,page_num,category)
        draw_card(cv,MARGIN,current_y,CARD_W,ch,product)
        current_y-=ch+CARD_GAP
    cv.showPage(); cv.save()

@app.route('/health', methods=['GET'])
def health():
    load_fonts()
    return jsonify({'status': 'ok', 'fonts': list(FONT_MAP.keys())})

@app.route('/add_product', methods=['POST'])
@app.route('/upload', methods=['POST'])
def add_product():
    """
    Form fields:
    - file: xlsm binary (zorunlu)
    - state: mevcut state JSON string (opsiyonel, Drive'dan gelir)
    
    Response:
    - PDF binary
    - Header X-State: güncel state JSON (Drive'a yazılacak)
    - Header X-Filename: pdf dosya adı
    - Header X-State-Filename: state dosya adı
    """
    try:
        load_fonts()
        if 'file' not in request.files:
            return jsonify({'error': 'file eksik'}), 400

        # xlsm parse et
        xlsm_bytes = request.files['file'].read()
        product = parse_xlsm(xlsm_bytes)
        category = product['category'] or 'GENEL'

        # Mevcut state'i oku (Drive'dan gelen)
        state = {}
        if 'state' in request.form and request.form['state']:
            try:
                state = json.loads(request.form['state'])
            except: pass

        # Yeni ürünü ekle
        state[product['ref']] = product

        # Tüm ürünlerden PDF oluştur
        products_list = list(state.values())

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            out_path = tmp.name

        build_pdf(products_list, out_path, category)

        with open(out_path,'rb') as fh:
            pdf_data = fh.read()
        os.unlink(out_path)

        safe_cat = category.replace(' ','_').replace('/','_').replace('&','and')
        filename       = f'katalog_{safe_cat}.pdf'
        state_filename = f'state_{safe_cat}.json'
        state_json     = json.dumps(state, ensure_ascii=False)

        return Response(
            pdf_data,
            mimetype='application/pdf',
            headers={
                'Content-Disposition':  f'attachment; filename="{filename}"',
                'X-Filename':           filename,
                'X-State-Filename':     state_filename,
                'X-Category':           category,
                'X-Product-Ref':        product['ref'],
                'X-Product-Count':      str(len(products_list)),
                'X-State':              base64.b64encode(state_json.encode()).decode(),
            }
        )

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
