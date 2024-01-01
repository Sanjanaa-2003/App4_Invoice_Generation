from pathlib import Path
from fpdf import FPDF
import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")

    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()

    filename = Path(filepath).stem  #imports the filename and results in '1000-2023.1.18
    invoice_nr,date = filename.split("-")#[0]     #splits as ["10001','2023.1.18']

    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50, h=8,txt=f"Invoice nr.{invoice_nr}",ln=1) #ln= adds new line

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}",ln=1)

    df = pd.read_excel(filepath,sheet_name="Sheet 1")

    #Add a header
    columns = df.columns
    columns=[item.replace("-"," ").title() for item in columns]
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0],border=1)  # AttributeError: 'int' object has no attribute 'replace' is raised if str() is not used
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1,ln=1)

    #Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)  #AttributeError: 'int' object has no attribute 'replace' is raised if str() is not used
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1,ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="",border=1)  # AttributeError: 'int' object has no attribute 'replace' is raised if str() is not used
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    #Add total sum sentence
    pdf.set_font(family="Times",size=10)
    pdf.cell(w=30,h=8,txt=f"The total price is {total_sum}",ln=1)

    #Add company name and logo
    pdf.set_font(family="Times", size=10,style="B")
    pdf.cell(w=25, h=8, txt=f"Pythonhow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")  #Creates a pdf in PDFs directory for each excel file

#openone of the pdf files to see date added in new line
