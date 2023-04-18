import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
  

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    data_frame = pd.read_excel(filepath, sheet_name="Sheet 1")
    header_list = list(data_frame.columns)
    header_list_refined = [header.replace("_", " ").title() for header in header_list]
    # for header in header_list:
    #     header_name = header.split("_")
    #     header_name = " ".join(header_name).capitalize()
    #     header_list_refined.append(header_name)
    
    # print(header_list_refined)
    #Add header for the table
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, border=1, txt=header_list_refined[0])
    pdf.cell(w=60, h=8, border=1, txt=header_list_refined[1])
    pdf.cell(w=35, h=8, border=1, txt=header_list_refined[2])
    pdf.cell(w=30, h=8, border=1, txt=header_list_refined[3]) 
    pdf.cell(w=30, h=8, border=1, txt=header_list_refined[4], ln=1)

    #Add rows for the table
    for index, row in data_frame.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, border=1, txt=str(row["product_id"]))
        pdf.cell(w=60, h=8, border=1, txt=str(row["product_name"]))
        pdf.cell(w=35, h=8, border=1, txt=str(row["amount_purchased"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["price_per_unit"])) 
        pdf.cell(w=30, h=8, border=1, txt=str(row["total_price"]), ln=1)


    pdf.output(f"PDFs/{filename}.pdf")