from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import os
import subprocess
import tempfile
import shutil
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/"output"; OUT.mkdir(exist_ok=True)
NAVY="0F1B3D"; ORANGE="FF6B35"; GOLD="F5C451"; WHITE="F7F8FC"; GREY="6B7280"; LIGHT="E9ECF3"
def build_excel():
    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE")
    modules = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE_MODULES")
    if not node or not modules:
        raise RuntimeError("Use the Codex primary runtime: CODEX_PRIMARY_RUNTIME_NODE and CODEX_PRIMARY_RUNTIME_NODE_MODULES are required.")
    with tempfile.TemporaryDirectory(prefix="mexay-workbook-") as directory:
        work = Path(directory)
        (work / "node_modules").symlink_to(modules, target_is_directory=True)
        builder = work / "build_workbook.mjs"
        shutil.copyfile(ROOT / "generator" / "build_workbook.mjs", builder)
        subprocess.run([node, str(builder), str(OUT)], check=True)

def build_pdf():
    p=OUT/"MexAy_Business_Excel_Kit_Guide.pdf"
    doc=SimpleDocTemplate(str(p),pagesize=A4,rightMargin=42,leftMargin=42,topMargin=42,bottomMargin=42)
    font_root = Path(os.environ.get("MEXAY_FONT_DIR", "/usr/share/fonts/truetype/dejavu"))
    for name, filename in [("MexSans", "DejaVuSans.ttf"), ("MexSansBold", "DejaVuSans-Bold.ttf")]:
        pdfmetrics.registerFont(TTFont(name, str(font_root / filename)))
    pdfmetrics.registerFontFamily("MexSans", normal="MexSans", bold="MexSansBold", italic="MexSans", boldItalic="MexSansBold")
    s=getSampleStyleSheet()
    for style in s.byName.values():
        style.fontName="MexSans"
    s["Heading2"].fontName="MexSansBold"
    s["Heading2"].fontSize=13
    s["Heading2"].leading=17
    s["Heading2"].spaceBefore=8
    s["Heading2"].spaceAfter=6
    s["BodyText"].fontSize=10
    s["BodyText"].leading=15
    title=ParagraphStyle("MexTitle",parent=s["Title"],textColor=colors.HexColor("#0F1B3D"),fontSize=22,spaceAfter=10)
    story=[Paragraph("MexAy Business Excel Kit",title),Paragraph("Küçük işletmeler için randevu, müşteri ve finans takibi.",s["Heading2"]),Spacer(1,10)]
    rows=[["Bölüm","Ne yapar?"],["Dashboard","Randevu, müşteri, gelir, gider ve net kazanç özetini gösterir."],["Randevular","Randevuları ve durumlarını takip eder."],["Müşteriler","Müşteri bilgilerini ve harcama geçmişini tutar."],["Finans","Gelir ve giderleri kaydeder."],["Ayarlar","İşletme bilgileri ve çalışma saatlerini içerir."]]
    table_header=ParagraphStyle("TableHeader",parent=s["BodyText"],fontName="MexSansBold",textColor=colors.white)
    rows=[[Paragraph(cell,table_header if index==0 else s["BodyText"]) for cell in row] for index,row in enumerate(rows)]
    t=Table(rows,colWidths=[112,399]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0F1B3D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#D9DDE7")),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)])); story += [t,Spacer(1,12)]
    for h,b in [("1. Başlangıç","Ayarlar sayfasını doldurun. Dosya boş başlar; örnek müşteri veya gelir içermez."),("2. Randevular","Tarih, saat, müşteri, hizmet, ücret ve durum alanlarını kullanın."),("3. Müşteriler","Müşteri iletişim bilgilerini ve ziyaret özetlerini saklayın."),("4. Finans","Gerçekleşen gelir ve gider tutarlarını Finans sayfasına girin. Randevu ücreti gelir toplamına otomatik eklenmez."),("5. Dashboard","Günlük işletme görünümünü tek sayfadan takip edin."),("6. Yedekleme","Önemli değişikliklerden sonra Excel dosyanızın bir kopyasını alın.")]:
        story += [Paragraph(h,s["Heading2"]),Paragraph(b,s["BodyText"]),Spacer(1,6)]
    story += [Paragraph("Kapasite ve kullanım",s["Heading2"]),Paragraph("Her giriş sayfası 500 kayıt içerir (2-501. satırlar). Daha fazla kayıt için tablolar, özet formülleri ve doğrulama aralıkları birlikte uzatılmalıdır. Tarih ve saati Excel tarih/saat değeri olarak girin. Müşteri ziyaret ve harcama özetleri elle güncellenir. Para tutarları TRY olarak hesaplanır. Çalışma saatleri bilgi amaçlıdır; otomatik randevu engelleme sağlamaz.",s["BodyText"])]
    story += [Spacer(1,12),Paragraph("Lisans: Tek işletme kullanımı. Yeniden satış, paylaşım veya dağıtım yasaktır.",s["BodyText"])]
    doc.build(story)

def visual(name,headline,sub,accent):
    img=Image.new("RGB",(1600,900),f"#{NAVY}"); d=ImageDraw.Draw(img)
    try:
        f=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",70); sm=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",34)
    except: f=sm=ImageFont.load_default()
    d.text((100,150),headline,font=f,fill=f"#{WHITE}"); d.text((105,285),sub,font=sm,fill=f"#{accent}")
    d.rounded_rectangle((100,410,1500,760),30,fill=f"#{WHITE}")
    d.text((150,500),"DASHBOARD  •  RANDEVULAR  •  MÜŞTERİLER  •  FİNANS",font=sm,fill=f"#{NAVY}")
    d.text((150,610),"MexAy Business Excel Kit",font=f,fill=f"#{ORANGE}")
    img.save(OUT/name)

def main():
    build_excel(); build_pdf()
    visual("Sales_Visual_01.png","Run your business from one file","Appointments, customers and finance",GOLD)
    visual("Sales_Visual_02.png","A clear dashboard","See what matters at a glance",ORANGE)
    visual("Sales_Visual_03.png","Launch your workflow today","MexAy Business Excel Kit • $39 launch",GOLD)
    license_text="MexAy Business Excel Kit\n\nLicense: single-business use only. Redistribution, resale, sublicensing, or public sharing of the included files is prohibited.\n\n© 2026 MexAy"
    (OUT/"LICENSE.txt").write_text(license_text,encoding="utf-8")
    with ZipFile(OUT/"MexAy_Business_Excel_Kit.zip","w",ZIP_DEFLATED) as z:
        for name in ["MexAy_Business_Management_Kit.xlsx", "MexAy_Business_Excel_Kit_Guide.pdf", "Sales_Visual_01.png", "Sales_Visual_02.png", "Sales_Visual_03.png", "LICENSE.txt"]:
            z.write(OUT/name, name)
    print("Product build complete")

if __name__=="__main__": main()
