from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/"output"; OUT.mkdir(exist_ok=True)
NAVY="0F1B3D"; ORANGE="FF6B35"; GOLD="F5C451"; WHITE="F7F8FC"; GREY="6B7280"; LIGHT="E9ECF3"
thin=Side(style="thin",color="D9DDE7")

def header(ws, row=1):
    for c in ws[row]:
        c.font=Font(bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=NAVY)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c.border=Border(bottom=thin)
    ws.row_dimensions[row].height=28

def table_sheet(ws, widths):
    ws.freeze_panes="A2"
    ws.auto_filter.ref=f"A1:{get_column_letter(ws.max_column)}501"
    header(ws)
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    for r in range(2,502):
        ws.row_dimensions[r].height=21
        for c in range(1,ws.max_column+1):
            ws.cell(r,c).alignment=Alignment(vertical="center")
            ws.cell(r,c).border=Border(bottom=thin)
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_view.showGridLines=False

def build_excel():
    wb=Workbook()
    dash=wb.active; dash.title="Dashboard"; dash.sheet_view.showGridLines=False
    dash["A1"]="MexAy Business Excel Kit"; dash["A1"].font=Font(size=24,bold=True,color=NAVY)
    dash["A2"]="Zamanını planla, randevularını tek yerden yönet."; dash["A2"].font=Font(size=11,italic=True,color=GREY)
    dash.merge_cells("A1:H1"); dash.merge_cells("A2:H2")
    cards=[("A4","Toplam Randevu",'=COUNTIF(Randevular!A2:A501,"<>")'),
           ("C4","Bekleyen",'=COUNTIF(Randevular!G2:G501,"Bekliyor")'),
           ("E4","Onaylanan",'=COUNTIF(Randevular!G2:G501,"Onaylandı")'),
           ("G4","Tamamlanan",'=COUNTIF(Randevular!G2:G501,"Tamamlandı")'),
           ("A7","Toplam Müşteri",'=COUNTIF(Müşteriler!A2:A501,"<>")'),
           ("C7","Toplam Gelir",'=SUM(Finans!D2:D501)'),
           ("E7","Toplam Gider",'=SUM(Finans!H2:H501)'),
           ("G7","Net Kazanç",'=C8-E8')]
    for pos,label,formula in cards:
        c=dash[pos]; c.value=label; c.font=Font(bold=True,color=GREY); c.fill=PatternFill("solid",fgColor=LIGHT)
        v=dash.cell(c.row+1,c.column); v.value=formula; v.font=Font(size=16,bold=True,color=NAVY); v.fill=PatternFill("solid",fgColor=WHITE)
        dash.merge_cells(start_row=c.row,start_column=c.column,end_row=c.row,end_column=c.column+1)
        dash.merge_cells(start_row=c.row+1,start_column=c.column,end_row=c.row+1,end_column=c.column+1)
    dash["A11"]="Durum"; dash["B11"]="Adet"
    for cell in dash[11][0:2]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor=NAVY)
    for r,(s,fm) in enumerate([("Bekliyor",'=COUNTIF(Randevular!G2:G501,"Bekliyor")'),("Onaylandı",'=COUNTIF(Randevular!G2:G501,"Onaylandı")'),("Tamamlandı",'=COUNTIF(Randevular!G2:G501,"Tamamlandı")'),("İptal",'=COUNTIF(Randevular!G2:G501,"İptal")')],12):
        dash.cell(r,1,s); dash.cell(r,2,fm)
    chart=BarChart(); chart.title="Randevu Durumu"; chart.y_axis.title="Adet"; chart.x_axis.title="Durum"
    chart.add_data(Reference(dash,min_col=2,min_row=11,max_row=15),titles_from_data=True)
    chart.set_categories(Reference(dash,min_col=1,min_row=12,max_row=15)); dash.add_chart(chart,"D11")
    for col in range(1,9): dash.column_dimensions[get_column_letter(col)].width=18
    dash.sheet_properties.pageSetUpPr.fitToPage=True; dash.page_setup.fitToWidth=1

    ap=wb.create_sheet("Randevular")
    ap.append(["Tarih","Saat","Müşteri","Telefon","Hizmet","Ücret","Durum","Not"])
    sample=[["2026-10-05","10:00","Örnek Müşteri","05xx xxx xx xx","Saç + Sakal",750,"Bekliyor","Örnek satır — silebilirsiniz"]]
    ap.append(sample[0])
    for _ in range(499): ap.append(["","","","", "",0,"Bekliyor",""])
    table_sheet(ap,[14,10,24,18,22,14,16,36])
    ap["A2"].number_format="dd.mm.yyyy"; ap["F2"].number_format='#,##0.00'
    for r in range(3,502): ap[f"F{r}"].number_format='#,##0.00'
    dv=DataValidation(type="list",formula1='"Bekliyor,Onaylandı,Tamamlandı,İptal"',allow_blank=True)
    ap.add_data_validation(dv); dv.add("G2:G501")
    ap.conditional_formatting.add("G2:G501",FormulaRule(formula=['G2="İptal"'],fill=PatternFill("solid",fgColor="FDE2E1")))

    cu=wb.create_sheet("Müşteriler")
    cu.append(["Müşteri","Telefon","E-posta","Son Ziyaret","Toplam Ziyaret","Toplam Harcama","Not"])
    cu.append(["Örnek Müşteri","05xx xxx xx xx","ornek@example.com","05.10.2026",1,750,"Örnek satır — silebilirsiniz"])
    for _ in range(499): cu.append(["","","","",0,0,""])
    table_sheet(cu,[24,18,30,16,18,18,36])
    for r in range(2,502): cu[f"F{r}"].number_format='#,##0.00'

    fi=wb.create_sheet("Finans")
    fi.append(["Gelir Tarihi","Gelir Açıklaması","Gelir Kategorisi","Gelir Tutarı","Gider Tarihi","Gider Açıklaması","Gider Kategorisi","Gider Tutarı"])
    fi.append(["05.10.2026","Örnek hizmet","Hizmet",750,"05.10.2026","Örnek gider","Genel",100])
    for _ in range(499): fi.append(["","","",0,"","","",0])
    table_sheet(fi,[15,26,20,16,15,26,20,16])
    for r in range(2,502):
        fi[f"D{r}"].number_format='#,##0.00'; fi[f"H{r}"].number_format='#,##0.00'

    st=wb.create_sheet("Ayarlar")
    st.append(["Alan","Değer"])
    for row in [["İşletme Adı","Demo İşletme"],["Telefon",""],["E-posta",""],["Adres",""],["Çalışma Başlangıcı","09:00"],["Çalışma Bitişi","18:00"],["Para Birimi","TRY"],["Pazar","Kapalı"]]: st.append(row)
    table_sheet(st,[28,32])

    hi=wb.create_sheet("Hızlı Başlangıç"); hi.sheet_view.showGridLines=False
    hi.append(["MexAy Business Excel Kit — Hızlı Başlangıç"])
    steps=["Ayarlar: işletme bilgilerinizi doldurun.","Müşteriler: müşteri kayıtlarınızı ekleyin.","Randevular: tarih, saat, hizmet, ücret ve durumu girin.","Finans: tüm gelir ve giderleri kaydedin.","Dashboard: genel performansı takip edin.","Örnek satırlar: sarı/örnek kayıtları kendi verilerinizle değiştirin veya silin.","Yedek: düzenli olarak dosyanın kopyasını alın."]
    for s in steps: hi.append(["• "+s])
    hi["A1"].font=Font(size=20,bold=True,color=NAVY); hi.column_dimensions["A"].width=105
    hi.row_dimensions[1].height=34
    for r in range(2,9): hi.cell(r,1).alignment=Alignment(wrap_text=True,vertical="top"); hi.row_dimensions[r].height=32
    wb.save(OUT/"MexAy_Business_Management_Kit.xlsx")

def build_pdf():
    p=OUT/"MexAy_Business_Excel_Kit_Guide.pdf"
    doc=SimpleDocTemplate(str(p),pagesize=A4,rightMargin=42,leftMargin=42,topMargin=42,bottomMargin=42)
    s=getSampleStyleSheet()
    title=ParagraphStyle("MexTitle",parent=s["Title"],textColor=colors.HexColor("#0F1B3D"),fontSize=24,spaceAfter=12)
    story=[Paragraph("MexAy Business Excel Kit",title),Paragraph("Küçük işletmeler için randevu, müşteri ve finans takibi.",s["Heading2"]),Spacer(1,16)]
    rows=[["Bölüm","Ne yapar?"],["Dashboard","Randevu, müşteri, gelir, gider ve net kazanç özetini gösterir."],["Randevular","Randevuları ve durumlarını takip eder."],["Müşteriler","Müşteri bilgilerini ve harcama geçmişini tutar."],["Finans","Gelir ve giderleri kaydeder."],["Ayarlar","İşletme bilgileri ve çalışma saatlerini içerir."]]
    t=Table(rows,colWidths=[120,360]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0F1B3D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#D9DDE7")),("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),7)])); story += [t,Spacer(1,18)]
    for h,b in [("1. Başlangıç","Ayarlar sayfasını doldurun ve örnek kayıtları kendi verilerinizle değiştirin veya silin."),("2. Randevular","Tarih, saat, müşteri, hizmet, ücret ve durum alanlarını kullanın."),("3. Müşteriler","Müşteri iletişim bilgilerini ve ziyaret özetlerini saklayın."),("4. Finans","Gelir ve gider tutarlarını girin. Dashboard net kazancı otomatik özetler."),("5. Dashboard","Günlük işletme görünümünü tek sayfadan takip edin."),("6. Yedekleme","Önemli değişikliklerden sonra Excel dosyanızın bir kopyasını alın.")]:
        story += [Paragraph(h,s["Heading2"]),Paragraph(b,s["BodyText"]),Spacer(1,9)]
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
    for p in OUT.iterdir():
        if p.is_file(): p.unlink()
    build_excel(); build_pdf()
    visual("Sales_Visual_01.png","Run your business from one file","Appointments, customers and finance",GOLD)
    visual("Sales_Visual_02.png","A clear dashboard","See what matters at a glance",ORANGE)
    visual("Sales_Visual_03.png","Launch your workflow today","MexAy Business Excel Kit • $39 launch",GOLD)
    license_text="MexAy Business Excel Kit\n\nLicense: single-business use only. Redistribution, resale, sublicensing, or public sharing of the included files is prohibited.\n\n© 2026 MexAy"
    (OUT/"LICENSE.txt").write_text(license_text,encoding="utf-8")
    with ZipFile(OUT/"MexAy_Business_Excel_Kit.zip","w",ZIP_DEFLATED) as z:
        for p in OUT.iterdir():
            if p.is_file() and p.name!="MexAy_Business_Excel_Kit.zip": z.write(p,p.name)
    print("Product build complete")

if __name__=="__main__": main()
