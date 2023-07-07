from PyPDF2 import PdfWriter, PdfReader
import io
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor

certemplate = "template.pdf"
excelfile = "data.xlsx"
varname = "Name"
# print("> Enter the pixel dimensions for the text to be printed on the certificate:")
horz = 145
vert = 270
varfont = "Vollkorn-VariableFont_wght.ttf"
fontsize = 41
fontcolor = "#000000"
box_width = 500  # Width of the rectangular box
box_height = 50  # Height of the rectangular box

# create the certificate directory
os.makedirs("certificates", exist_ok=True)

# register the necessary font
pdfmetrics.registerFont(TTFont('myFont', varfont))

# provide the excel file that contains the participant names (in column 'Name')
data = pd.read_excel(excelfile)

name_list = data[varname].tolist()
names = [x.title().strip() for x in name_list]  # Change names to capitalized format
for i in names:
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Draw the rectangular box
    # can.rect(horz, vert, box_width, box_height)

    # the registered font is used, provide font size and color
    can.setFont("myFont", fontsize)
    can.setFillColor(HexColor(fontcolor))  # Set font color

    # Calculate the center position for the text
    text_width = pdfmetrics.stringWidth(i, "myFont", fontsize)
    text_height = fontsize
    text_x = horz + (box_width - text_width) / 2
    text_y = vert + (box_height - text_height) / 2

    # provide the text location in pixels
    can.drawString(text_x, text_y, i)

    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # provide the certificate template
    existing_pdf = PdfReader(open(certemplate, "rb"))

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    destination = "certificates" + os.sep + i + ".pdf"
    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    print("created " + i + ".pdf")
