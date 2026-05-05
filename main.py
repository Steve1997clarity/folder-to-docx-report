import os
import shutil
from io import BytesIO
from flask import Flask, request, send_file, render_template_string, Response
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # Demo limit: 20 MB max
app.secret_key = 'secret-key'

# --- Configuration ---
MAX_IMAGES = 20          # Max images per upload (demo limit)
MAX_IMAGE_SIZE_MB = 5    # Max per-image size
THUMB_SIZE = (300, 300)  # Thumbnail dimensions

@app.errorhandler(413)
def too_large(e):
    return "File too large. Demo limit is 20 MB total.", 413

# Default branding images (appear in generated reports only)
DEFAULT_HEADER_IMAGE = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_HEADER_IMAGE", "header_image.png")
DEFAULT_BOTTOM_IMAGE = os.path.join(os.getcwd(), "USER_INPUT", "DOCX_BOTTOM_IMAGE", "bottom_image.png")

# Demo presets
DEMO_PRESETS = {
    "elevation_6_path_1": {
        "name": "Building Elevation Survey",
        "description": "Exterior facade inspection photos captured by drone along elevation path 1.",
        "folder": os.path.join(os.getcwd(), "Gary_Project_testset", "Elevation 6_Path 1"),
        "image_count": 8,
        "date": "August 2024",
        "output_filename": "Elevation_Survey_Report.docx",
        "preview_images": [
            "DJI_20240807104359_0322_D.JPG",
            "DJI_20240807104407_0323_D.JPG",
            "DJI_20240807104415_0324_D.JPG",
            "DJI_20240807104425_0325_D.JPG"
        ]
    },
    "path_6_rthk": {
        "name": "Roof Condition Inspection",
        "description": "Aerial roof inspection photos for ancillary building condition assessment.",
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

    sectPr = section._sectPr
    vAlign = sectPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        sectPr.append(vAlign)
    vAlign.set(qn('w:val'), 'center')

    # Header
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
    else:
        run = header_para.add_run("Header Image Not Found")
        run.font.size = Pt(10)

    # Body: scan for JPEG images (resize for demo to keep output small)
    valid_images = []
    for root, dirs, files in os.walk(folder_path):
        for f in sorted(files):
            if f.lower().endswith((".jpg", ".jpeg")):
                image_path = os.path.join(root, f)
                try:
                    with Image.open(image_path) as img:
                        # Resize to max 600px wide for demo (smaller DOCX, faster download)
                        if img.width > 600:
                            ratio = 600 / img.width
                            new_size = (600, int(img.height * ratio))
                            img = img.resize(new_size, Image.LANCZOS)
                        stream = BytesIO()
                        img.convert('RGB').save(stream, format='JPEG', quality=55)
                        stream.seek(0)
                    valid_images.append((stream, f))
                except Exception as e:
                    print(f"Skipping '{f}': {e}")
                if len(valid_images) >= MAX_IMAGES:
                    break

    table = doc.add_table(rows=0, cols=images_per_row)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i in range(0, len(valid_images), images_per_row):
        batch = valid_images[i:i+images_per_row]
        image_row = table.add_row().cells
        for idx, (stream, img_name) in enumerate(batch):
            para = image_row[idx].paragraphs[0]
            run = para.add_run()
            try:
                run.add_picture(stream, width=image_width)
            except Exception as e:
                print(f"Error inserting '{img_name}': {e}")
        label_row = table.add_row().cells
        for idx, (stream, img_name) in enumerate(batch):
            para = label_row[idx].paragraphs[0]
            para.text = img_name
            para.paragraph_format.space_after = Pt(12)

    # Footer
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
    else:
        run = footer_para.add_run("Bottom Image Not Found")
        run.font.size = Pt(10)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    output_path = os.path.join(output_folder, output_docx)
    doc.save(output_path)
    return output_path


# In-memory caches (generated once)
_thumb_cache = {}
_demo_report_cache = {}


def _pregenerate_demos():
    """Pre-generate demo reports on startup for instant serving."""
    print("Pre-generating demo reports...")
    output_dir = os.path.join(os.getcwd(), "OUTPUT")
    os.makedirs(output_dir, exist_ok=True)
    for preset_id, preset in DEMO_PRESETS.items():
        path = create_docx_with_images_header_footer(
            folder_path=preset["folder"],
            header_image_path=DEFAULT_HEADER_IMAGE,
            bottom_image_path=DEFAULT_BOTTOM_IMAGE,
            output_docx=preset["output_filename"],
            output_folder=output_dir
        )
        with open(path, 'rb') as f:
            _demo_report_cache[preset_id] = f.read()
        print(f"  Cached: {preset['name']} ({len(_demo_report_cache[preset_id])//1024} KB)")
    # Also pre-warm thumbnail cache
    for preset_id, preset in DEMO_PRESETS.items():
        for img_name in preset["preview_images"]:
            cache_key = f"{preset_id}/{img_name}"
            image_path = os.path.join(preset["folder"], img_name)
            if os.path.exists(image_path):
                img = Image.open(image_path)
                img.thumbnail(THUMB_SIZE)
                buf = BytesIO()
                img.convert('RGB').save(buf, format='JPEG', quality=60)
                _thumb_cache[cache_key] = buf.getvalue()
    print("Demo reports and thumbnails cached.")

@app.route('/demo/thumbnail/<preset_id>/<filename>')
def demo_thumbnail(preset_id, filename):
    if preset_id not in DEMO_PRESETS:
        return "Not found", 404
    preset = DEMO_PRESETS[preset_id]
    safe_filename = os.path.basename(filename)
    cache_key = f"{preset_id}/{safe_filename}"

    if cache_key not in _thumb_cache:
        image_path = os.path.join(preset["folder"], safe_filename)
        if not os.path.exists(image_path):
            return "Not found", 404
        img = Image.open(image_path)
        img.thumbnail(THUMB_SIZE)
        buf = BytesIO()
        img.convert('RGB').save(buf, format='JPEG', quality=60)
        _thumb_cache[cache_key] = buf.getvalue()

    response = Response(_thumb_cache[cache_key], mimetype='image/jpeg')
    response.headers['Cache-Control'] = 'public, max-age=86400'
    return response


@app.route('/demo/download/<preset_id>')
def demo_download(preset_id):
    """GET-based download (works reliably through Codespace proxy)."""
    if preset_id not in DEMO_PRESETS:
        return "Invalid preset", 404

    preset = DEMO_PRESETS[preset_id]

    # Serve from pre-generated cache (instant response)
    if preset_id in _demo_report_cache:
        data = _demo_report_cache[preset_id]
    else:
        # Fallback: generate on the fly
        output_dir = os.path.join(os.getcwd(), "OUTPUT")
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
        output_dir = os.path.join(os.getcwd(), "OUTPUT")
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        upload_dir = os.path.join(os.getcwd(), "USER_INPUT", "USER_UPLOADED_FOLDER")
        if os.path.exists(upload_dir):
            shutil.rmtree(upload_dir)
        os.makedirs(upload_dir, exist_ok=True)

        folder_files = request.files.getlist("folder_files")
        saved_count = 0
        for file in folder_files:
            if saved_count >= MAX_IMAGES:
                break
            filename = file.filename
            safe_filename = os.path.normpath(filename)
            if ".." in safe_filename:
                continue
            if not safe_filename.lower().endswith((".jpg", ".jpeg")):
                continue
            dest_path = os.path.join(upload_dir, safe_filename)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            file.save(dest_path)
            saved_count += 1

        header_image_path = DEFAULT_HEADER_IMAGE
        bottom_image_path = DEFAULT_BOTTOM_IMAGE
        output_docx = "Survey_Report.docx"
        generated_docx_path = create_docx_with_images_header_footer(
            folder_path=upload_dir,
            header_image_path=header_image_path,
            bottom_image_path=bottom_image_path,
            output_docx=output_docx,
            output_folder=output_dir
        )

        with open(generated_docx_path, 'rb') as f:
            data = f.read()
        response = Response(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        response.headers["Content-Disposition"] = f"attachment; filename={output_docx}"
        return response
    else:
        # Build demo cards
        demo_cards_html = ""
        for preset_id, preset in DEMO_PRESETS.items():
            thumbnails_html = ""
            for img_name in preset["preview_images"]:
                thumbnails_html += f'<img src="/demo/thumbnail/{preset_id}/{img_name}" alt="Survey photo" class="thumb-img" loading="lazy">'

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
                            <span class="badge-pill">{preset["image_count"]} photos</span>
                            <span class="badge-pill">{preset["date"]}</span>
                        </div>
                        <a href="/demo/download/{preset_id}" class="btn-demo" onclick="handleDemoClick(this)">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
                            Generate Sample Report
                        </a>
                    </div>
                </div>
            </div>
            '''

        html_page = '''
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AI Document Automation for Building Surveyors</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
      :root {
        --primary: #1e40af;
        --primary-hover: #1e3a8a;
        --primary-light: #eff6ff;
        --accent: #059669;
        --accent-hover: #047857;
        --text-dark: #0f172a;
        --text-body: #334155;
        --text-muted: #64748b;
        --bg-light: #f8fafc;
        --border: #e2e8f0;
        --card-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
        --card-shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.08), 0 4px 6px -4px rgba(0,0,0,0.04);
        --radius: 16px;
        --radius-sm: 10px;
      }
      * { box-sizing: border-box; }
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
        color: var(--text-body);
        background: var(--bg-light);
        margin: 0;
        line-height: 1.6;
      }

      /* --- Navigation --- */
      .nav-bar {
        background: white;
        border-bottom: 1px solid var(--border);
        padding: 14px 0;
        position: sticky;
        top: 0;
        z-index: 100;
      }
      .nav-brand {
        font-weight: 800;
        font-size: 1rem;
        color: var(--primary);
        letter-spacing: -0.02em;
      }
      .nav-tag {
        font-size: 0.7rem;
        background: var(--primary-light);
        color: var(--primary);
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: 600;
        margin-left: 10px;
      }

      /* --- Hero --- */
      .hero {
        background: linear-gradient(160deg, #1e3a8a 0%, #1e40af 50%, #2563eb 100%);
        padding: 80px 0 60px;
        color: white;
        text-align: center;
      }
      .hero h1 {
        font-size: 2.5rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        margin-bottom: 16px;
        line-height: 1.2;
      }
      .hero .lead {
        font-size: 1.15rem;
        opacity: 0.9;
        max-width: 600px;
        margin: 0 auto 32px;
        line-height: 1.7;
      }
      .hero-cta {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: white;
        color: var(--primary);
        padding: 14px 28px;
        border-radius: var(--radius-sm);
        font-weight: 700;
        font-size: 0.95rem;
        text-decoration: none;
        transition: all 0.2s;
      }
      .hero-cta:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0,0,0,0.2); color: var(--primary); }

      /* --- Problem Section --- */
      .problem-section {
        padding: 60px 0;
        background: white;
        border-bottom: 1px solid var(--border);
      }
      .problem-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 28px;
        margin-top: 32px;
      }
      .problem-card {
        padding: 28px 24px;
        border-radius: var(--radius);
        background: var(--bg-light);
        border: 1px solid var(--border);
      }
      .problem-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: #fef2f2;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 14px;
        color: #dc2626;
      }
      .problem-card h4 {
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 6px;
        color: var(--text-dark);
      }
      .problem-card p {
        font-size: 0.85rem;
        color: var(--text-muted);
        margin: 0;
      }

      /* --- Solution Section --- */
      .solution-section {
        padding: 60px 0;
      }
      .solution-steps {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 24px;
        margin-top: 32px;
      }
      .sol-step {
        text-align: center;
        padding: 32px 20px;
        background: white;
        border-radius: var(--radius);
        border: 1px solid var(--border);
        box-shadow: var(--card-shadow);
      }
      .sol-num {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        background: var(--primary-light);
        color: var(--primary);
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        margin-bottom: 14px;
      }
      .sol-step h4 {
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 6px;
      }
      .sol-step p {
        font-size: 0.82rem;
        color: var(--text-muted);
        margin: 0;
      }

      /* --- Section Headers --- */
      .section-header {
        text-align: center;
        margin-bottom: 8px;
      }
      .section-label {
        display: inline-block;
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: var(--primary);
        background: var(--primary-light);
        padding: 4px 12px;
        border-radius: 20px;
        margin-bottom: 12px;
      }
      .section-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: var(--text-dark);
        letter-spacing: -0.02em;
        margin-bottom: 8px;
      }
      .section-desc {
        color: var(--text-muted);
        font-size: 0.95rem;
        max-width: 520px;
        margin: 0 auto;
      }

      /* --- Demo Section --- */
      .demo-section {
        padding: 60px 0;
        background: white;
        border-top: 1px solid var(--border);
        border-bottom: 1px solid var(--border);
      }
      .demo-card {
        background: white;
        border-radius: var(--radius);
        box-shadow: var(--card-shadow);
        overflow: hidden;
        transition: all 0.2s ease;
        border: 1px solid var(--border);
      }
      .demo-card:hover {
        box-shadow: var(--card-shadow-lg);
        transform: translateY(-3px);
      }
      .thumb-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 2px;
        background: #e2e8f0;
      }
      .thumb-img {
        width: 100%;
        aspect-ratio: 4/3;
        object-fit: cover;
        display: block;
      }
      .card-body { padding: 20px; }
      .card-title {
        font-weight: 700;
        font-size: 1rem;
        color: var(--text-dark);
        margin-bottom: 6px;
      }
      .card-desc {
        color: var(--text-muted);
        font-size: 0.82rem;
        margin-bottom: 14px;
      }
      .card-meta {
        display: flex;
        gap: 8px;
        margin-bottom: 16px;
        flex-wrap: wrap;
      }
      .badge-pill {
        background: var(--primary-light);
        color: var(--primary);
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.72rem;
        font-weight: 600;
      }
      .btn-demo {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: var(--primary);
        color: white;
        border: none;
        padding: 12px 20px;
        border-radius: var(--radius-sm);
        font-size: 0.85rem;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
        width: 100%;
        justify-content: center;
        text-decoration: none;
      }
      .btn-demo:hover { background: var(--primary-hover); transform: translateY(-1px); color: white; }
      a.btn-demo:visited { color: white; }

      /* --- Upload Section --- */
      .upload-section {
        padding: 60px 0;
      }
      .upload-box {
        background: white;
        border-radius: var(--radius);
        padding: 36px;
        box-shadow: var(--card-shadow);
        border: 1px solid var(--border);
        max-width: 640px;
        margin: 32px auto 0;
      }
      .dropzone {
        border: 2px dashed #cbd5e1;
        border-radius: var(--radius-sm);
        padding: 44px 20px;
        text-align: center;
        cursor: pointer;
        transition: all 0.2s;
        background: var(--bg-light);
      }
      .dropzone:hover, .dropzone.dragover {
        border-color: var(--primary);
        background: var(--primary-light);
      }
      .dropzone-icon { margin-bottom: 12px; color: #94a3b8; }
      .dropzone-text { font-size: 0.9rem; color: var(--text-body); margin-bottom: 4px; }
      .dropzone-hint { font-size: 0.78rem; color: #94a3b8; }
      .file-count {
        margin-top: 12px;
        font-size: 0.85rem;
        color: var(--accent);
        font-weight: 600;
        display: none;
      }
      .file-count.visible { display: block; }
      .btn-upload {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: var(--accent);
        color: white;
        border: none;
        padding: 14px 28px;
        border-radius: var(--radius-sm);
        font-size: 0.9rem;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
        margin-top: 20px;
      }
      .btn-upload:hover:not(:disabled) { background: var(--accent-hover); transform: translateY(-1px); }
      .btn-upload:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }

      /* --- Limits Notice --- */
      .limits-notice {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-radius: var(--radius-sm);
        padding: 14px 18px;
        margin-top: 20px;
        font-size: 0.78rem;
        color: #92400e;
      }
      .limits-notice strong { color: #78350f; }

      /* --- Footer --- */
      .footer {
        background: var(--text-dark);
        color: #94a3b8;
        padding: 40px 0;
        text-align: center;
        font-size: 0.8rem;
      }
      .footer p { margin: 4px 0; }

      /* --- Utilities --- */
      .spinner {
        display: inline-block;
        width: 14px;
        height: 14px;
        border: 2px solid rgba(255,255,255,0.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 0.6s linear infinite;
      }
      @keyframes spin { to { transform: rotate(360deg); } }

      @media (max-width: 768px) {
        .hero h1 { font-size: 1.7rem; }
        .hero { padding: 50px 0 40px; }
        .problem-grid, .solution-steps { grid-template-columns: 1fr; }
        .upload-box { padding: 24px 18px; }
      }
    </style>
</head>
<body>
    <!-- Nav -->
    <nav class="nav-bar">
      <div class="container d-flex align-items-center">
        <span class="nav-brand">SurveyDoc AI</span>
        <span class="nav-tag">Prototype Demo</span>
      </div>
    </nav>

    <!-- Hero -->
    <section class="hero">
      <div class="container">
        <h1>AI-Driven Document Automation<br>for Building Surveyors</h1>
        <p class="lead">Stop spending hours formatting inspection photos in Word. Upload your survey images and get a professionally branded report in seconds.</p>
        <a href="#demo" class="hero-cta">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>
          Try the Live Demo
        </a>
      </div>
    </section>

    <!-- Problem Statement -->
    <section class="problem-section">
      <div class="container">
        <div class="section-header">
          <span class="section-label">The Problem</span>
          <h2 class="section-title">Report Formatting is a Time Sink</h2>
          <p class="section-desc">Independent surveyors and small firms waste valuable hours on repetitive document formatting instead of focusing on professional expertise.</p>
        </div>
        <div class="problem-grid">
          <div class="problem-card">
            <div class="problem-icon">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            </div>
            <h4>Hours Lost to Formatting</h4>
            <p>Manually inserting, resizing, and labelling dozens of inspection photos in Word for every single report.</p>
          </div>
          <div class="problem-card">
            <div class="problem-icon">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
            </div>
            <h4>No IT Support</h4>
            <p>Independent practitioners and small firms operate without dedicated IT teams or enterprise software solutions.</p>
          </div>
          <div class="problem-card">
            <div class="problem-icon">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
            </div>
            <h4>Generic Tools Only</h4>
            <p>Relying on generic Word processing software that was never designed for photo-heavy survey reports.</p>
          </div>
        </div>
      </div>
    </section>

    <!-- Solution -->
    <section class="solution-section">
      <div class="container">
        <div class="section-header">
          <span class="section-label">The Solution</span>
          <h2 class="section-title">From Photos to Report in Seconds</h2>
          <p class="section-desc">AI-powered automation handles the tedious formatting so surveyors can focus on what matters.</p>
        </div>
        <div class="solution-steps">
          <div class="sol-step">
            <div class="sol-num">1</div>
            <h4>Upload Inspection Photos</h4>
            <p>Select a folder of JPG images from your drone or camera. The system reads them automatically.</p>
          </div>
          <div class="sol-step">
            <div class="sol-num">2</div>
            <h4>AI Formats Your Report</h4>
            <p>Photos are arranged in a professional grid with labels, branded headers, and company footers.</p>
          </div>
          <div class="sol-step">
            <div class="sol-num">3</div>
            <h4>Download &amp; Deliver</h4>
            <p>Get a ready-to-send Word document. No manual resizing, no formatting, no wasted time.</p>
          </div>
        </div>
      </div>
    </section>

    <!-- Live Demo -->
    <section class="demo-section" id="demo">
      <div class="container">
        <div class="section-header">
          <span class="section-label">Live Demo</span>
          <h2 class="section-title">Try It Yourself</h2>
          <p class="section-desc">Click a sample dataset below to instantly generate a real report. No sign-up required.</p>
        </div>
        <div class="row mt-4">
          ''' + demo_cards_html + '''
        </div>
      </div>
    </section>

    <!-- Upload Your Own -->
    <section class="upload-section" id="upload">
      <div class="container">
        <div class="section-header">
          <span class="section-label">Try Your Own</span>
          <h2 class="section-title">Upload Your Photos</h2>
          <p class="section-desc">Test with your own inspection images to see the output quality.</p>
        </div>
        <div class="upload-box">
          <form id="upload-form" method="post" enctype="multipart/form-data">
            <div class="dropzone" id="dropzone" onclick="document.getElementById('folder_files').click()">
              <div class="dropzone-icon">
                <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><line x1="12" y1="11" x2="12" y2="17"/><polyline points="9 14 12 11 15 14"/></svg>
              </div>
              <div class="dropzone-text">Click to select a folder of JPG inspection images</div>
              <div class="dropzone-hint">Supports folder upload with subfolders</div>
            </div>
            <input type="file" id="folder_files" name="folder_files" webkitdirectory directory multiple required style="display:none;">
            <div class="file-count" id="file-count"></div>
            <div class="limits-notice">
              <strong>Demo Limits:</strong> Max 20 images per report &middot; Max 5 MB per image &middot; 20 MB total upload &middot; JPG/JPEG only
            </div>
            <div class="text-center">
              <button type="submit" class="btn-upload" id="upload-btn" disabled>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                Generate Report
              </button>
            </div>
          </form>
        </div>
      </div>
    </section>

    <!-- Footer -->
    <footer class="footer">
      <div class="container">
        <p><strong style="color:#e2e8f0;">SurveyDoc AI</strong> &mdash; Prototype Demonstration</p>
        <p>AI-Driven Document Automation for Building Survey Professionals</p>
        <p style="margin-top: 12px; font-size: 0.72rem; color: #475569;">This is a technology demonstration. Generated reports use sample branding for illustration purposes.</p>
      </div>
    </footer>

    <script>
      const fileInput = document.getElementById('folder_files');
      const fileCount = document.getElementById('file-count');
      const uploadBtn = document.getElementById('upload-btn');
      const dropzone = document.getElementById('dropzone');

      fileInput.addEventListener('change', function() {
        const files = Array.from(this.files).filter(f => f.name.toLowerCase().match(/\\.jpe?g$/));
        if (files.length > 0) {
          const count = Math.min(files.length, 20);
          fileCount.textContent = count + ' JPG image' + (count !== 1 ? 's' : '') + ' ready' + (files.length > 20 ? ' (first 20 will be used)' : '');
          fileCount.classList.add('visible');
          uploadBtn.disabled = false;
        } else {
          fileCount.textContent = 'No JPG images found in selected folder';
          fileCount.classList.add('visible');
          fileCount.style.color = '#dc2626';
          uploadBtn.disabled = true;
        }
      });

      dropzone.addEventListener('dragover', function(e) { e.preventDefault(); this.classList.add('dragover'); });
      dropzone.addEventListener('dragleave', function() { this.classList.remove('dragover'); });
      dropzone.addEventListener('drop', function(e) { e.preventDefault(); this.classList.remove('dragover'); });

      document.getElementById('upload-form').addEventListener('submit', function() {
        uploadBtn.disabled = true;
        uploadBtn.innerHTML = '<span class="spinner"></span> Generating Report...';
      });

      function handleDemoClick(btn) {
        btn.style.pointerEvents = 'none';
        btn.style.opacity = '0.7';
        btn.innerHTML = '<span class="spinner"></span> Downloading...';
        // Re-enable after 8 seconds (file should have downloaded by then)
        setTimeout(function() {
          btn.style.pointerEvents = '';
          btn.style.opacity = '';
          btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> Generate Sample Report';
        }, 8000);
      }

      // Smooth scroll for CTA
      document.querySelector('.hero-cta').addEventListener('click', function(e) {
        e.preventDefault();
        document.getElementById('demo').scrollIntoView({ behavior: 'smooth' });
      });
    </script>
</body>
</html>
        '''
        return render_template_string(html_page)

if __name__ == '__main__':
    _pregenerate_demos()
    app.run(debug=False, port=11312, host="0.0.0.0")
