import os
import shutil
from io import BytesIO
from flask import Flask, request, send_file, render_template_string, Response
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image  # Pillow library

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Allow up to 100 MB files
app.secret_key = 'secret-key'

@app.errorhandler(413)
def too_large(e):
    return "File is too large. The maximum allowed size is 100 MB.", 413

# Set the default header and bottom image paths (inside the container)
DEFAULT_HEADER_IMAGE = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_HEADER_IMAGE", "header_image.png")
DEFAULT_BOTTOM_IMAGE = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_BOTTOM_IMAGE", "bottom_image.png")

def create_docx_with_images_header_footer(folder_path, header_image_path, bottom_image_path,
                                           output_docx="output.docx", output_folder="OUTPUT",
                                           image_width=Inches(2), images_per_row=2):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
    # --- Set Vertical Alignment for the Section ---
    sectPr = section._sectPr
    vAlign = sectPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        sectPr.append(vAlign)
    vAlign.set(qn('w:val'), 'center')
    
    # --- Header Section ---
    header_section = section.header
    header_para = header_section.paragraphs[0] if header_section.paragraphs else header_section.add_paragraph()
    header_para.text = ""
    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if os.path.exists(header_image_path):
        try:
            run = header_para.add_run()
            run.add_picture(header_image_path, width=Inches(2.5))
        except Exception as e:
            run = header_para.add_run("Error inserting header image")
            run.font.size = Pt(10)
            print("Error inserting header image:", e)
    else:
        run = header_para.add_run("Header Image Not Found")
        run.font.size = Pt(10)
        print("Header image not found:", header_image_path)
    
    # --- Main Body: Recursively scan for JPEG images and create the table ---
    valid_images = []
    for root, dirs, files in os.walk(folder_path):
        for f in sorted(files):
            if f.lower().endswith(".jpg"):
                image_path = os.path.join(root, f)
                try:
                    with Image.open(image_path) as img:
                        stream = BytesIO()
                        img.convert('RGB').save(stream, format='JPEG')
                        stream.seek(0)
                    valid_images.append((stream, f))
                except Exception as e:
                    print(f"Skipping '{f}' due to error: {e}")
    
    table = doc.add_table(rows=0, cols=images_per_row)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i in range(0, len(valid_images), images_per_row):
        batch = valid_images[i:i+images_per_row]
        # --- Image Row ---
        image_row = table.add_row().cells
        for idx, (stream, img_name) in enumerate(batch):
            para = image_row[idx].paragraphs[0]
            run = para.add_run()
            try:
                run.add_picture(stream, width=image_width)
            except Exception as e:
                print(f"Error inserting image '{img_name}': {e}")
        # --- Label Row ---
        label_row = table.add_row().cells
        for idx, (stream, img_name) in enumerate(batch):
            para = label_row[idx].paragraphs[0]
            para.text = img_name
            para.paragraph_format.space_after = Pt(12)
    
    # --- Footer Section ---
    footer_section = section.footer
    for para in footer_section.paragraphs:
        p = para._element
        p.getparent().remove(p)
    footer_para = footer_section.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(bottom_image_path):
        try:
            run = footer_para.add_run()
            run.add_picture(bottom_image_path, height=Cm(1.08))
        except Exception as e:
            run = footer_para.add_run("Error inserting bottom image")
            run.font.size = Pt(10)
            print("Error inserting bottom image:", e)
    else:
        run = footer_para.add_run("Bottom Image Not Found")
        run.font.size = Pt(10)
        print("Bottom image not found:", bottom_image_path)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    output_path = os.path.join(output_folder, output_docx)
    doc.save(output_path)
    return output_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Clear the persistent OUTPUT folder
        output_dir = os.path.join(os.getcwd(), "OUTPUT")
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        
        # Clear / re-create USER_UPLOADED_FOLDER folder
        upload_dir = os.path.join(os.getcwd(), "USER_INPUT", "USER_UPLOADED_FOLDER")
        if os.path.exists(upload_dir):
            shutil.rmtree(upload_dir)
        os.makedirs(upload_dir, exist_ok=True)
        
        # Save uploaded folder files into USER_UPLOADED_FOLDER
        folder_files = request.files.getlist("folder_files")
        for file in folder_files:
            filename = file.filename  # may include subdirectory info
            safe_filename = os.path.normpath(filename)
            if ".." in safe_filename:
                continue
            dest_path = os.path.join(upload_dir, safe_filename)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            file.save(dest_path)
        
        # Generate the DOCX file using the persistent uploaded folder and default images
        header_image_path = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_HEADER_IMAGE", "header_image.png")
        bottom_image_path = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_BOTTOM_IMAGE", "bottom_image.png")
        output_docx = "output.docx"
        generated_docx_path = create_docx_with_images_header_footer(
            folder_path=upload_dir,
            header_image_path=header_image_path,
            bottom_image_path=bottom_image_path,
            output_docx=output_docx,
            output_folder=output_dir
        )
        
        # Send the generated DOCX for download
        with open(generated_docx_path, 'rb') as f:
            data = f.read()
        response = Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        response.headers["Content-Disposition"] = f"attachment; filename={output_docx}"
        return response
    else:
        html_form = '''
        <!doctype html>
        <html>
          <head>
            <title>DOCX Report Generator</title>
            <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
            <style>
              body { margin-top: 50px; }
              .container { max-width: 600px; }
              .dropzone {
                  border: 2px dashed #007bff;
                  padding: 20px;
                  text-align: center;
                  color: #007bff;
                  cursor: pointer;
              }
            </style>
          </head>
          <body>
            <div class="container">
              <h1 class="mb-4">Generate DOCX Report</h1>
              <form method="post" enctype="multipart/form-data">
                <div class="form-group">
                  <label for="folder_files">Upload Folder (JPEG images &ndash; drag & drop or click to select):</label>
                  <input type="file" class="form-control-file" id="folder_files" name="folder_files" webkitdirectory directory multiple required>
                </div>
                <button type="submit" class="btn btn-primary">Generate DOCX Report</button>
              </form>
              <p class="mt-3 text-muted">Note: The header and footer images are provided by default.</p>
            </div>
          </body>
        </html>
        '''
        return render_template_string(html_form)

if __name__ == '__main__':
    app.run(debug=True, port=11312, host="0.0.0.0")
