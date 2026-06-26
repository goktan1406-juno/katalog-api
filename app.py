from flask import Flask, request, jsonify
import base64, zipfile, os, tempfile
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
from PIL import Image, ImageFilter, ImageEnhance

app = Flask(__name__)

# Font yolları — birden fazla yer dene
FONT_CANDIDATES = {
    'Sans': [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/DejaVuSans.ttf',
    ],
    'SansBold': [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
        '/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf',
        '/usr/share/fonts/truetype/DejaVuSans-Bold.ttf',
    ],
    'SansObl': [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf',
        '/usr/share/fonts/dejavu/DejaVuSans-Oblique.ttf',
        '/usr/share/fonts/truetype/DejaVuSans-Oblique.ttf',
    ],
}

fonts_loaded = False
FONT_MAP = {}  # hangi fontların yüklendiği

def load_fonts():
    global fonts_loaded, FONT_MAP
    if fonts_loaded: return
    for name, paths in FONT_CANDIDATES.items():
        for path in paths:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(name, path))
                    FONT_MAP[name] = path
                    break
                except: pass
    fonts_loaded = True

def get_font(name, fallback='Helvetica'):
    """Font yüklendiyse onu kullan, yoksa Helvetica"""
    return name if name in FONT_MAP else fallback

def get_font_bold(fallback='Helvetica-Bold'):
    return 'SansBold' if 'SansBold' in FONT_MAP else fallback

def get_font_obl(fallback='Helvetica-Oblique'):
    return 'SansObl' if 'SansObl' in FONT_MAP else fallback

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
    product['images'] = extract_images(buf)
    return product

def extract_images(xlsm_buf):
    imgs = []
    xlsm_buf.seek(0)
    try:
        with zipfile.ZipFile(xlsm_buf,'r') as z:
            for path in [f for f in z.namelist() if f.startswith('xl/media/')]:
                data = z.read(path)
                try:
                    im = Image.open(BytesIO(data))
                    if min(im.size) >= 200:
                        imgs.append(prep_image(data))
                        if len(imgs)==3: break
                except: pass
    except: pass
    return imgs

def prep_image(data, target=900):
    im = Image.open(BytesIO(data))
    if im.mode=='P': im=im.convert('RGBA')
    if im.mode=='RGBA':
        bg=Image.new('RGB',im.size,(255,255,255)); bg.paste(im,mask=im.split()[3]); im=bg
    else: im=im.convert('RGB')
    w,h=im.size; s=max(w,h)
    sq=Image.new('RGB',(s,s),(255,255,255)); sq.paste(im,((s-w)//2,(s-h)//2)); im=sq
    if im.size[0]<target: im=im.resize((target,target),Image.LANCZOS)
    im=im.filter(ImageFilter.UnsharpMask(radius=1.5,percent=150,threshold=2))
    im=ImageEnhance.Contrast(im).enhance(1.05)
    out=BytesIO(); im.save(out,'JPEG',quality=97); out.seek(0)
    return ImageReader(out)

def calc_card_h(product):
    n=len(product['benefits'])
    return max(65*mm, min(8*mm+5*mm+4.5*mm+4*mm + n*(4.5*mm+2*4.2*mm+1.5*mm) + 10*mm, 150*mm))

def draw_page_chrome(cv, page_num, category):
    F  = get_font('Sans')
    FB = get_font_bold()
    HDR=13*mm
    cv.setFillColor(DARK); cv.rect(0,H-HDR,W,HDR,fill=1,stroke=0)
    cv.setFillColor(RED);  cv.rect(0,H-HDR,3.5*mm,HDR,fill=1,stroke=0)
    cv.setFillColor(WHITE); cv.setFont(FB,9)
    cv.drawString(8*mm,H-HDR+4.5*mm,str(category).upper())
    cv.setFont(F,7.5); cv.setFillColor(MGRAY)
    cv.drawRightString(W-8*mm,H-HDR+4.5*mm,'Urun Katalogu 2024')
    FTR=10*mm
    cv.setFillColor(LGRAY); cv.rect(0,0,W,FTR,fill=1,stroke=0)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3); cv.line(0,FTR,W,FTR)
    cv.setFillColor(DGRAY); cv.setFont(F,6.5)
    cv.drawString(8*mm,3.5*mm,'2024 TEFAL')
    cv.drawRightString(W-8*mm,3.5*mm,'tefal.com.tr')
    cv.setFillColor(DARK); cv.setFont(FB,7)
    cv.drawCentredString(W/2,3.5*mm,f'{page_num} | TEFAL')

def draw_card(cv, x, y, cw, ch, product):
    F  = get_font('Sans')
    FB = get_font_bold()
    FO = get_font_obl()
    PAD=4*mm; GAP=5*mm
    imgs=product.get('images',[]); n_imgs=len(imgs)
    LW=(cw-2*PAD-GAP)*0.48 if n_imgs>0 else cw-2*PAD
    RW=(cw-2*PAD-GAP)*0.52 if n_imgs>0 else 0
    LX=x+PAD; RX=x+PAD+LW+GAP; BOT=y-ch

    cv.setFillColor(WHITE); cv.roundRect(x,BOT,cw,ch,2*mm,fill=1,stroke=0)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.5); cv.roundRect(x,BOT,cw,ch,2*mm,fill=0,stroke=1)
    cv.setFillColor(RED); cv.roundRect(x,y-5*mm,3*mm,5*mm,1*mm,fill=1,stroke=0)
    cv.rect(x+1.5*mm,y-5*mm,1.5*mm,5*mm,fill=1,stroke=0)

    cy=y-PAD
    cv.setFillColor(DARK); cv.setFont(FB,9)
    for ln in wrap(cv,product['name'],FB,9,LW)[:2]:
        cv.drawString(LX,cy,ln); cy-=5.5*mm
    cy-=1*mm

    cv.setFillColor(LGRAY)
    bw=tw(cv,product['ref'],F,6.5)+5*mm
    cv.roundRect(LX,cy-5*mm,bw,5*mm,1.5*mm,fill=1,stroke=0)
    cv.setFillColor(DGRAY); cv.setFont(F,6.5)
    cv.drawString(LX+2.5*mm,cy-3.5*mm,product['ref']); cy-=8*mm

    cv.setStrokeColor(RED); cv.setLineWidth(1.5); cv.line(LX,cy,LX+LW*0.45,cy)
    cv.setStrokeColor(MGRAY); cv.setLineWidth(0.3); cv.line(LX+LW*0.45+2*mm,cy,LX+LW,cy)
    cy-=4*mm

    if product.get('claim'):
        cv.setFillColor(RED); cv.setFont(FO,7.5)
        claim=str(product['claim'])
        while tw(cv,claim,FO,7.5)>LW and len(claim)>5: claim=claim[:-4]+'...'
        cv.drawString(LX,cy,claim); cy-=4.5*mm
        cv.setStrokeColor(colors.HexColor('#EBEBEB')); cv.setLineWidth(0.25)
        cv.line(LX,cy,LX+LW,cy); cy-=2.5*mm

    for title,detail in product['benefits']:
        if cy-4.5*mm<BOT+PAD+4*mm: break
        cv.setFillColor(RED); cv.setFont(FB,9); cv.drawString(LX,cy-4.5*mm+1.5*mm,'*')
        cv.setFillColor(DARK); cv.setFont(FB,7.5)
        t=str(title)
        while tw(cv,t,FB,7.5)>LW-5*mm and len(t)>5: t=t[:-2]+'.'
        cv.drawString(LX+4.5*mm,cy-4.5*mm+1.5*mm,t); cy-=4.5*mm
        if detail:
            cv.setFillColor(DGRAY); cv.setFont(F,7.0)
            for dl in wrap(cv,str(detail),F,7.0,LW-4.5*mm)[:2]:
                if cy-4.2*mm<BOT+PAD+4*mm: break
                cv.drawString(LX+4.5*mm,cy-4.2*mm+1.5*mm,dl); cy-=4.2*mm
        cv.setStrokeColor(colors.HexColor('#F0F0F0')); cv.setLineWidth(0.2)
        cv.line(LX,cy-1*mm,LX+LW,cy-1*mm); cy-=1.5*mm

    cv.setFillColor(LGRAY); cv.roundRect(LX,BOT+2*mm,LW,6*mm,1*mm,fill=1,stroke=0)
    cv.setFillColor(DGRAY); cv.setFont(F,6)
    cv.drawString(LX+3*mm,BOT+4.5*mm,f"Kutu Icerigi: {str(product['name']).split(',')[0]} - {product['ref']}")

    if n_imgs==0: return
    area_h=ch-2*PAD
    if n_imgs==1: slots=[(RX,y-PAD-area_h,RW,area_h)]
    elif n_imgs==2:
        h1=area_h*0.62; h2=area_h-h1-3*mm
        slots=[(RX,y-PAD-h1,RW,h1),(RX,y-PAD-h1-3*mm-h2,RW,h2)]
    else:
        h1=area_h*0.55; hr=(area_h-h1-6*mm)/2
        slots=[(RX,y-PAD-h1,RW,h1),(RX,y-PAD-h1-3*mm-hr,RW,hr),(RX,y-PAD-h1-6*mm-2*hr,RW,hr)]

    for i,(px,py,pw,ph) in enumerate(slots):
        if i>=len(imgs): break
        cv.setFillColor(LGRAY); cv.roundRect(px,py,pw,ph,2*mm,fill=1,stroke=0)
        try: cv.drawImage(imgs[i],px+1.5*mm,py+1.5*mm,pw-3*mm,ph-3*mm,preserveAspectRatio=True,anchor='c',mask='auto')
        except: pass
        cv.setStrokeColor(MGRAY); cv.setLineWidth(0.4); cv.roundRect(px,py,pw,ph,2*mm,fill=0,stroke=1)

def build_catalog(products, output_path, category='GENEL'):
    load_fonts()
    cv=canvas.Canvas(output_path,pagesize=A4)
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
    return output_path

@app.route('/health', methods=['GET'])
def health():
    load_fonts()
    return jsonify({'status': 'ok', 'fonts': FONT_MAP})

@app.route('/upload', methods=['POST'])
def upload():
    try:
        load_fonts()
        if 'file' not in request.files:
            return jsonify({'error': 'file yok'}), 400

        categories = defaultdict(list)
        for f in request.files.getlist('file'):
            try:
                xlsm_bytes = f.read()
                p = parse_xlsm(xlsm_bytes)
                categories[p['category'] or 'GENEL'].append(p)
            except Exception as e:
                print(f"Hata: {f.filename} -> {e}")

        results = []
        for cat, products in categories.items():
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                out_path = tmp.name
            build_catalog(products, out_path, cat)
            with open(out_path,'rb') as fh:
                pdf_b64 = base64.b64encode(fh.read()).decode()
            os.unlink(out_path)
            results.append({
                'category':      cat,
                'filename':      f'katalog_{cat.replace(" ","_")}.pdf',
                'pdf_base64':    pdf_b64,
                'product_count': len(products),
                'products':      [p['ref'] for p in products],
            })
        return jsonify({'catalogs': results})

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/katalog', methods=['POST'])
def katalog():
    try:
        load_fonts()
        payload = request.get_json()
        if not payload or 'files' not in payload:
            return jsonify({'error': 'files alani eksik'}), 400

        categories = defaultdict(list)
        for f in payload['files']:
            try:
                p = parse_xlsm(base64.b64decode(f['data']))
                categories[p['category'] or 'GENEL'].append(p)
            except Exception as e:
                print(f"Hata: {f.get('name')} -> {e}")

        results = []
        for cat, products in categories.items():
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                out_path = tmp.name
            build_catalog(products, out_path, cat)
            with open(out_path,'rb') as fh:
                pdf_b64 = base64.b64encode(fh.read()).decode()
            os.unlink(out_path)
            results.append({
                'category':      cat,
                'filename':      f'katalog_{cat.replace(" ","_")}.pdf',
                'pdf_base64':    pdf_b64,
                'product_count': len(products),
                'products':      [p['ref'] for p in products],
            })
        return jsonify({'catalogs': results})
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
