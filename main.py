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

# Demo presets configuration
DEMO_PRESETS = {
    "elevation_6_path_1": {
        "name": "Elevation 6 - Path 1",
        "description": "8 aerial drone photos from a building elevation survey",
        "folder": os.path.join(os.getcwd(), "Gary_Project_testset", "Elevation 6_Path 1"),
        "image_count": 8,
        "date": "August 2024",
        "output_filename": "Elevation_6_Path_1_Report.docx",
        "preview_images": [
            "DJI_20240807104359_0322_D.JPG",
            "DJI_20240807104407_0323_D.JPG",
            "DJI_20240807104415_0324_D.JPG",
            "DJI_20240807104425_0325_D.JPG"
        ]
    },
    "path_6_rthk": {
        "name": "Roof Inspection - Ancillary Building A",
        "description": "6 aerial roof inspection photos from a building survey",
        "folder": os.path.join(os.getcwd(), "Gary_Project_testset",
                  "Path 6_RTHK_BH_Ancillary Building A_Roof_Aerial Photos"),
        "image_count": 6,
        "date": "January 2025",
        "output_filename": "Roof_Inspection_Report.docx",
        "preview_images": [
            "DJI_20250117124043_0017_V_SMALL.JPG",
            "DJI_20250117124049_0018_V_SMALL.JPG",
            "DJI_20250117124054_0019_V_SMALL.JPG",
            "DJI_20250117124058_0020_V_SMALL.JPG"
        ]
    }
}


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


@app.route('/demo/thumbnail/<preset_id>/<filename>')
def demo_thumbnail(preset_id, filename):
    if preset_id not in DEMO_PRESETS:
        return "Not found", 404
    preset = DEMO_PRESETS[preset_id]
    # Sanitize filename
    safe_filename = os.path.basename(filename)
    image_path = os.path.join(preset["folder"], safe_filename)
    if not os.path.exists(image_path):
        return "Not found", 404

    # Generate thumbnail
    img = Image.open(image_path)
    img.thumbnail((300, 300))
    buf = BytesIO()
    img.convert('RGB').save(buf, format='JPEG', quality=70)
    buf.seek(0)

    response = Response(buf.getvalue(), mimetype='image/jpeg')
    response.headers['Cache-Control'] = 'public, max-age=3600'
    return response


@app.route('/demo/generate', methods=['POST'])
def demo_generate():
    preset_id = request.form.get('preset_id')
    if preset_id not in DEMO_PRESETS:
        return "Invalid preset", 400

    preset = DEMO_PRESETS[preset_id]
    output_dir = os.path.join(os.getcwd(), "OUTPUT")
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    generated_path = create_docx_with_images_header_footer(
        folder_path=preset["folder"],
        header_image_path=DEFAULT_HEADER_IMAGE,
        bottom_image_path=DEFAULT_BOTTOM_IMAGE,
        output_docx=preset["output_filename"],
        output_folder=output_dir
    )

    with open(generated_path, 'rb') as f:
        data = f.read()
    response = Response(data, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    response.headers["Content-Disposition"] = f"attachment; filename={preset['output_filename']}"
    return response


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
        # Build demo preset cards HTML dynamically
        demo_cards_html = ""
        for preset_id, preset in DEMO_PRESETS.items():
            thumbnails_html = ""
            for img_name in preset["preview_images"]:
                thumbnails_html += f'<img src="/demo/thumbnail/{preset_id}/{img_name}" alt="{img_name}" class="thumb-img">'

            demo_cards_html += f'''
            <div class="col-md-6 mb-4">
                <div class="demo-card">
                    <div class="thumb-grid">
                        {thumbnails_html}
                    </div>
                    <div class="card-body">
                        <h5 class="card-title">{preset["name"]}</h5>
                        <p class="card-desc">{preset["description"]}</p>
                        <div class="card-meta">
                            <span class="badge-custom">{preset["image_count"]} photos</span>
                            <span class="badge-custom">{preset["date"]}</span>
                        </div>
                        <form method="post" action="/demo/generate" class="demo-form">
                            <input type="hidden" name="preset_id" value="{preset_id}">
                            <button type="submit" class="btn-generate" onclick="handleDemoClick(this)">
                                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
                                Generate Demo Report
                            </button>
                        </form>
                    </div>
                </div>
            </div>
            '''

        html_form = '''
        <!doctype html>
        <html lang="en">
          <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>DOCX Report Generator</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            <style>
              :root {
                --primary: #2563eb;
                --primary-hover: #1d4ed8;
                --primary-light: #eff6ff;
                --secondary: #f8fafc;
                --accent: #10b981;
                --accent-hover: #059669;
                --text-dark: #1e293b;
                --text-muted: #64748b;
                --card-shadow: 0 4px 6px -1px rgba(0,0,0,0.07), 0 2px 4px -2px rgba(0,0,0,0.05);
                --card-shadow-hover: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -4px rgba(0,0,0,0.1);
                --radius: 16px;
                --radius-sm: 10px;
              }
              * { box-sizing: border-box; }
              body {
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                background: #f1f5f9;
                color: var(--text-dark);
                margin: 0;
                padding: 0;
              }
              .navbar-custom {
                background: white;
                border-bottom: 1px solid #e2e8f0;
                padding: 12px 0;
              }
              .navbar-brand {
                font-weight: 700;
                font-size: 1.1rem;
                color: var(--text-dark);
              }
              .hero {
                background: linear-gradient(135deg, var(--primary-light) 0%, #ffffff 100%);
                padding: 60px 0 40px;
                text-align: center;
              }
              .hero h1 {
                font-size: 2.2rem;
                font-weight: 800;
                color: var(--text-dark);
                margin-bottom: 12px;
              }
              .hero p {
                font-size: 1.1rem;
                color: var(--text-muted);
                max-width: 550px;
                margin: 0 auto;
              }
              .section-title {
                font-size: 1.4rem;
                font-weight: 700;
                margin-bottom: 8px;
                color: var(--text-dark);
              }
              .section-subtitle {
                color: var(--text-muted);
                margin-bottom: 24px;
              }
              .section {
                padding: 40px 0;
              }
              .demo-card {
                background: white;
                border-radius: var(--radius);
                box-shadow: var(--card-shadow);
                overflow: hidden;
                transition: all 0.2s ease;
                border: 1px solid #e2e8f0;
              }
              .demo-card:hover {
                box-shadow: var(--card-shadow-hover);
                transform: translateY(-2px);
              }
              .thumb-grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 2px;
                background: #f1f5f9;
                padding: 2px;
              }
              .thumb-img {
                width: 100%;
                aspect-ratio: 4/3;
                object-fit: cover;
                display: block;
              }
              .card-body {
                padding: 20px;
              }
              .card-title {
                font-weight: 700;
                font-size: 1rem;
                margin-bottom: 6px;
              }
              .card-desc {
                color: var(--text-muted);
                font-size: 0.875rem;
                margin-bottom: 12px;
              }
              .card-meta {
                display: flex;
                gap: 8px;
                margin-bottom: 16px;
              }
              .badge-custom {
                background: var(--primary-light);
                color: var(--primary);
                padding: 4px 10px;
                border-radius: 20px;
                font-size: 0.75rem;
                font-weight: 600;
              }
              .btn-generate {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: var(--primary);
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: var(--radius-sm);
                font-size: 0.9rem;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s ease;
                width: 100%;
                justify-content: center;
              }
              .btn-generate:hover {
                background: var(--primary-hover);
                transform: translateY(-1px);
              }
              .btn-generate:disabled {
                opacity: 0.7;
                cursor: not-allowed;
                transform: none;
              }
              .upload-section {
                background: white;
                border-radius: var(--radius);
                padding: 32px;
                box-shadow: var(--card-shadow);
                border: 1px solid #e2e8f0;
              }
              .dropzone {
                border: 2px dashed #cbd5e1;
                border-radius: var(--radius-sm);
                padding: 40px 20px;
                text-align: center;
                cursor: pointer;
                transition: all 0.2s ease;
                background: var(--secondary);
              }
              .dropzone:hover, .dropzone.dragover {
                border-color: var(--primary);
                background: var(--primary-light);
              }
              .dropzone-icon {
                margin-bottom: 12px;
              }
              .dropzone-text {
                font-size: 0.95rem;
                color: var(--text-muted);
                margin-bottom: 4px;
              }
              .dropzone-hint {
                font-size: 0.8rem;
                color: #94a3b8;
              }
              .file-count {
                margin-top: 12px;
                font-size: 0.875rem;
                color: var(--accent);
                font-weight: 600;
                display: none;
              }
              .file-count.visible {
                display: block;
              }
              .btn-upload {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: var(--accent);
                color: white;
                border: none;
                padding: 14px 32px;
                border-radius: var(--radius-sm);
                font-size: 1rem;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s ease;
                margin-top: 20px;
              }
              .btn-upload:hover:not(:disabled) {
                background: var(--accent-hover);
                transform: translateY(-1px);
              }
              .btn-upload:disabled {
                opacity: 0.5;
                cursor: not-allowed;
                transform: none;
              }
              .steps {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 24px;
                text-align: center;
              }
              .step {
                padding: 24px 16px;
              }
              .step-number {
                width: 40px;
                height: 40px;
                border-radius: 50%;
                background: var(--primary-light);
                color: var(--primary);
                display: inline-flex;
                align-items: center;
                justify-content: center;
                font-weight: 700;
                font-size: 1rem;
                margin-bottom: 12px;
              }
              .step-title {
                font-weight: 700;
                margin-bottom: 4px;
                font-size: 0.95rem;
              }
              .step-desc {
                font-size: 0.8rem;
                color: var(--text-muted);
              }
              .footer {
                background: white;
                border-top: 1px solid #e2e8f0;
                padding: 24px 0;
                text-align: center;
                color: var(--text-muted);
                font-size: 0.8rem;
              }
              .spinner {
                display: inline-block;
                width: 16px;
                height: 16px;
                border: 2px solid rgba(255,255,255,0.3);
                border-radius: 50%;
                border-top-color: white;
                animation: spin 0.6s linear infinite;
              }
              @keyframes spin {
                to { transform: rotate(360deg); }
              }
              @media (max-width: 768px) {
                .hero h1 { font-size: 1.6rem; }
                .steps { grid-template-columns: 1fr; gap: 12px; }
                .hero { padding: 40px 0 24px; }
              }
            </style>
          </head>
          <body>
            <!-- Navbar -->
            <nav class="navbar-custom">
              <div class="container">
                <span class="navbar-brand">DOCX Report Generator</span>
              </div>
            </nav>

            <!-- Hero -->
            <section class="hero">
              <div class="container">
                <h1>Turn Drone Photos into<br>Professional Reports</h1>
                <p>Upload a folder of inspection photos and instantly get a formatted Word document with branded headers and footers. No software to install.</p>
              </div>
            </section>

            <!-- Demo Presets -->
            <section class="section">
              <div class="container">
                <h2 class="section-title">Try It Now</h2>
                <p class="section-subtitle">Click a demo below to instantly generate a sample report - no upload needed.</p>
                <div class="row">
                  ''' + demo_cards_html + '''
                </div>
              </div>
            </section>

            <!-- Upload Section -->
            <section class="section" style="padding-top: 0;">
              <div class="container">
                <h2 class="section-title">Upload Your Own Photos</h2>
                <p class="section-subtitle">Select a folder of JPG images from your computer.</p>
                <div class="upload-section">
                  <form id="upload-form" method="post" enctype="multipart/form-data">
                    <div class="dropzone" id="dropzone" onclick="document.getElementById('folder_files').click()">
                      <div class="dropzone-icon">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><line x1="12" y1="11" x2="12" y2="17"/><polyline points="9 14 12 11 15 14"/></svg>
                      </div>
                      <div class="dropzone-text">Click to select a folder of JPG images</div>
                      <div class="dropzone-hint">or drag and drop a folder here</div>
                    </div>
                    <input type="file" id="folder_files" name="folder_files" webkitdirectory directory multiple required style="display:none;">
                    <div class="file-count" id="file-count"></div>
                    <div class="text-center">
                      <button type="submit" class="btn-upload" id="upload-btn" disabled>
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
                        Generate My Report
                      </button>
                    </div>
                  </form>
                </div>
              </div>
            </section>

            <!-- How It Works -->
            <section class="section" style="background: white; border-top: 1px solid #e2e8f0;">
              <div class="container">
                <h2 class="section-title text-center">How It Works</h2>
                <p class="section-subtitle text-center">Three simple steps to your formatted report.</p>
                <div class="steps">
                  <div class="step">
                    <div class="step-number">1</div>
                    <div class="step-title">Upload Photos</div>
                    <div class="step-desc">Select a folder containing your JPG inspection images</div>
                  </div>
                  <div class="step">
                    <div class="step-number">2</div>
                    <div class="step-title">Auto-Format</div>
                    <div class="step-desc">Images are arranged in a professional grid with labels</div>
                  </div>
                  <div class="step">
                    <div class="step-number">3</div>
                    <div class="step-title">Download Report</div>
                    <div class="step-desc">Get a branded Word document ready to share</div>
                  </div>
                </div>
              </div>
            </section>

            <!-- Footer -->
            <footer class="footer">
              <div class="container">
                <p>DOCX Report Generator &middot; Powered by Metapeller Limited</p>
              </div>
            </footer>

            <script>
              // File input handling
              const fileInput = document.getElementById('folder_files');
              const fileCount = document.getElementById('file-count');
              const uploadBtn = document.getElementById('upload-btn');
              const dropzone = document.getElementById('dropzone');

              fileInput.addEventListener('change', function() {
                const files = this.files;
                if (files.length > 0) {
                  const jpgCount = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.jpg')).length;
                  fileCount.textContent = jpgCount + ' JPG image' + (jpgCount !== 1 ? 's' : '') + ' ready to process';
                  fileCount.classList.add('visible');
                  uploadBtn.disabled = false;
                } else {
                  fileCount.classList.remove('visible');
                  uploadBtn.disabled = true;
                }
              });

              // Drag and drop visual feedback
              dropzone.addEventListener('dragover', function(e) {
                e.preventDefault();
                this.classList.add('dragover');
              });
              dropzone.addEventListener('dragleave', function() {
                this.classList.remove('dragover');
              });
              dropzone.addEventListener('drop', function(e) {
                e.preventDefault();
                this.classList.remove('dragover');
              });

              // Upload form loading state
              document.getElementById('upload-form').addEventListener('submit', function() {
                uploadBtn.disabled = true;
                uploadBtn.innerHTML = '<span class="spinner"></span> Generating Report...';
              });

              // Demo button loading state
              function handleDemoClick(btn) {
                btn.disabled = true;
                btn.innerHTML = '<span class="spinner"></span> Generating...';
              }
            </script>
          </body>
        </html>
        '''
        return render_template_string(html_form)

if __name__ == '__main__':
    app.run(debug=True, port=11312, host="0.0.0.0")
