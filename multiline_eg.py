from fpdf import FPDF

pdf = FPDF(orientation="P", unit="mm", format="A4")
pdf.add_page()

content = """
Lorem ipsum dolor sit amet, consectetur adipiscing 
elit, sed do eiusmod tempor incididunt ut labore 
et dolore magna aliqua. Ut enim ad minim veniam, 
quis nostrud exercitation ullamco.
"""

pdf.set_font(family="Times", size=12)
pdf.multi_cell(w=0, h=6, txt=content)
pdf.output("output.pdf")

#
# So far, we have only added single-line cells to a PDF using the pdf.cell() method. However, that method does not work
# for bigger pieces of text that expand across multiple lines. For such text, you should use the multi_cell method:
# pdf.multi_cell(w, h, txt). The code above will generate an output.pdf file with multiline text inside.
