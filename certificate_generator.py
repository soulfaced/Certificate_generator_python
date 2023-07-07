import os
# from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from flask import Flask, render_template, request, send_file, send_from_directory


app = Flask(__name__)

# Certificate template path
certemplate = "template.pdf"

# Get the absolute path of the current directory
current_directory = os.path.dirname(os.path.abspath(__file__))

# Specify the path to the templates directory
templates_directory = os.path.join(current_directory, 'templates')


@app.route('/')
def index():
    error = request.args.get('error')
    return render_template('index.html', error=error)


@app.route('/generate_certificate', methods=['POST'])
def generate_certificate():
    # Get the name from the form submission
    name = request.form['name'].strip().title()

    # Validate the name length
    if len(name.split()) > 20:
        return render_template('index.html', error='Name exceeds word limit')

    # Create the certificates directory if it doesn't exist
    os.makedirs("certificates", exist_ok=True)

    # Register the necessary font
    pdfmetrics.registerFont(TTFont('myFont', "Vollkorn-VariableFont_wght.ttf"))

    # Generate the certificate for the given name
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Read the certificate template
    existing_pdf = PdfReader(open(certemplate, "rb"))

    # Get the first page of the template
    page = existing_pdf.pages[0]

    # Set the font, font size, and font color
    can.setFont("myFont", 41)
    can.setFillColor(HexColor("#000000"))

    # Calculate the text position
    text_width = pdfmetrics.stringWidth(name, "myFont", 41)
    text_height = 41
    text_x = 145 + (500 - text_width) / 2
    text_y = 270 + (50 - text_height) / 2

    # Draw the name on the certificate
    can.drawString(text_x, text_y, name)
    can.save()

    # Merge the generated certificate with the template
    new_pdf = PdfReader(packet)
    output = PdfWriter()
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    # Save the generated certificate
    certificate_path = os.path.join("certificates", f"{name}.pdf")
    with open(certificate_path, "wb") as f:
        output.write(f)

    return render_template('certificate.html', name=name)


@app.route('/download_certificate/<name>', methods=['GET'])
def download_certificate(name):
    # Generate the certificate path
    certificate_path = os.path.join("certificates", f"{name}.pdf")

    # Check if the certificate file exists
    if os.path.exists(certificate_path):
        return send_from_directory("certificates", f"{name}.pdf", as_attachment=True)
    else:
        return render_template('index.html', error='Certificate not found')


if __name__ == '__main__':
    app.run()
