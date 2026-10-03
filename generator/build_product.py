from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[1]; OUT=ROOT/"output"; OUT.mkdir(exist_ok=True)
NAVY="0F1B3D"; ORANGE="FF6B35"; GOLD="F5C451"; WHITE="F7F8FC"; GREY="6B7280"
thin=Side(style="thin",color="D9DDE7")

def style_sheet(ws):
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for c in ws[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(horizontal="center"); c.border=Border(bottom=thin)
    for col in ws.columns:
        letter=col[0].column_letter; ws.column_dimensions[letter].width=min(max(max(len(str(x.value or "")) for x in col)+2,12),28)

def build_excel():
    wb=Workbook(); dash=wb.active; dash.title="Dashboard"; dash.sheet_view.showGridLines=False
    dash["A1"]="MexAy Business Excel Kit"; dash["A1"].font=Font(size=22,bold=True,color=NAVY)
    dash["A2"]="Zamanını planla, randevularını tek yerden yönet."; dash["A2"].font=Font(italic=True,color=GREY)
    metrics=[("A4","Toplam Randevu","=COUNTA('Randevular'!A2:A501)"),("A5","Bekleyen","=COUNTIF('Randevular'!G2:G501,\"Bekliyor\")"),("A6","Onaylanan","=COUNTIF('Randevular'!G2:G501,\"Onaylandı\")"),("A7","Tamamlanan","=COUNTIF('Randevular'!G2:G501,\"Tamamlandı\")"),("D4","Toplam Müşteri","=COUNTA('Müşteriler'!A2:A501)"),("D5","Toplam Gelir","=SUM('Finans'!D2:D501)"),("D6","Toplam Gider","=SUM('Finans'!H2:H501)"),("D7","Net Kazanç","=D5-D6")]
    for pos,label,formula in metrics:
        dash[pos]=label; dash[pos].font=Font(bold=True,color=GREY); v=dash.cell(dash[pos].row,dash[pos].column+1); v.value=formula; v.font=Font(size=16,bold=True,color=NAVY)
    for col,w in {"A":22,"B":18,"D":22,"E":18}.items(): dash.column_dimensions[col].width=w
    ap=wb.create_sheet("Randevular"); ap.append(["Tarih","Saat","Müşteri","Telefon","Hizmet","Ücret","Durum","Not"])
    for _ in range(20): ap.append(["","","","","",0,"Bekliyor",""])
    dv=DataValidation(type="list",formula1='"Bekliyor,Onaylandı,Tamamlandı,İptal"'); ap.add_data_validation(dv); dv.add("G2:G501"); style_sheet(ap)
    cu=wb.create_sheet("Müşteriler"); cu.append(["Müşteri","Telefon","E-posta","Son Ziyaret","Toplam Ziyaret","Toplam Harcama","Not"])
    for _ in range(20): cu.append(["","","","",0,0,""]); style_sheet(cu)
    fi=wb.create_sheet("Finans"); fi.append(["Gelir Tarihi","Gelir Açıklaması","Gelir Kategorisi","Gelir Tutarı","Gider Tarihi","Gider Açıklaması","Gider Kategorisi","Gider Tutarı"])
    for _ in range(20): fi.append(["","","",0,"","","",0]); style_sheet(fi)
    st=wb.create_sheet("Ayarlar"); st.append(["Alan","Değer"])
    for row in [["İşletme Adı",""],["Telefon",""],["E-posta",""],["Adres",""],["Çalışma Başlangıcı","09:00"],["Çalışma Bitişi","18:00"],["Para Birimi","TRY"]]: st.append(row)
    style_sheet(st)
    hi=wb.create_sheet("Hızlı Başlangıç"); hi.append(["MexAy Business Excel Kit"])
    for x in ["1. Ayarlar sayfasını doldurun.","2. İşletme bilgilerinizi belirleyin.","3. Müşterilerinizi ekleyin.","4. Randevularınızı girin.","5. Gelir ve giderlerinizi kaydedin.","6. Dashboard üzerinden genel durumu takip edin."]: hi.append([x])
    hi["A1"].font=Font(size=18,bold=True,color=NAVY); hi.column_dimensions["A"].width=80
    wb.save(OUT/"MexAy_Business_Management_Kit.xlsx")

def build_pdf():
    p=OUT/"MexAy_Business_Excel_Kit_Guide.pdf"; doc=SimpleDocTemplate(str(p),pagesize=A4,rightMargin=45,leftMargin=45,topMargin=45,bottomMargin=45)
    s=getSampleStyleSheet(); title=ParagraphStyle("MexTitle",parent=s["Title"],textColor=colors.HexColor("#0F1B3D"),fontSize=24)
    story=[Paragraph("MexAy Business Excel Kit",title),Paragraph("Zamanını planla, randevularını tek yerden yönet.",s["Heading2"]),Spacer(1,20)]
    for h,b in [("1. Başlangıç","Önce Ayarlar sayfasındaki işletme bilgilerinizi doldurun."),("2. Randevular","Tarih, saat, müşteri, hizmet, ücret ve durum alanlarını doldurun."),("3. Müşteriler","Müşteri ve ziyaret bilgilerinizi tek listede tutun."),("4. Finans","Gelir ve giderleri kaydedin; Dashboard net kazancı özetler."),("5. Dashboard","Randevu, müşteri, gelir, gider ve net kazanç özetini takip edin."),("6. Yedek","Düzenli olarak dosyanızın bir kopyasını alın.")]:
        story += [Paragraph(h,s["Heading2"]),Paragraph(b,s["BodyText"]),Spacer(1,12)]
    doc.build(story)

def visual(name,headline,sub):
    img=Image.new("RGB",(1600,900),f"#{NAVY}"); d=ImageDraw.Draw(img)
    try:
        f=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",72); sm=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",34)
    except: f=sm=ImageFont.load_default()
    d.text((100,180),headline,font=f,fill=f"#{WHITE}"); d.text((105,300),sub,font=sm,fill=f"#{GOLD}")
    d.rounded_rectangle((100,420,1500,760),30,fill=f"#{WHITE}"); d.text((150,500),"DASHBOARD  •  RANDEVULAR  •  MÜŞTERİLER  •  FİNANS",font=sm,fill=f"#{NAVY}"); d.text((150,600),"MexAy Business Excel Kit",font=f,fill=f"#{ORANGE}"); img.save(OUT/name)

def main():
    build_excel(); build_pdf(); visual("Sales_Visual_01.png","MexAy Business Excel Kit","One dashboard for your business"); visual("Sales_Visual_02.png","Everything in one file","Appointments • Customers • Finance"); visual("Sales_Visual_03.png","Save time. Track more.","Launch price: $39")
    with ZipFile(OUT/"MexAy_Business_Excel_Kit.zip","w",ZIP_DEFLATED) as z:
        for p in OUT.iterdir():
            if p.is_file() and p.name!="MexAy_Business_Excel_Kit.zip": z.write(p,p.name)
    print("Product build complete")

if __name__=="__main__": main()
