# DOCX Report Generator

AI-driven document automation for building survey professionals. Upload inspection photos, customise branding, and generate professional Word reports instantly.

**Live:** `http://100.50.38.143:11312`
**Progress Board:** [https://tinyurl.com/22jc6zag](https://tinyurl.com/22jc6zag)

## Features

### Report Generation
- **3x4 Photo Grid** -- exactly 3 columns x 4 rows (12 photos per page) with an automatic page break between pages (2.2" image width)
- **Sequential Photo Labels** -- Auto-numbered (Photo 1_01, Photo 1_02, ...) with inline editing before generation
- **Folder Metadata** -- Auto-extracts Master Folder / Folder / Sub-Folder / Path from upload directory structure, inserted as a table in the DOCX
- **Photo Preview Grid** -- 3-column thumbnail preview with drag-to-replace support. Drop a JPG onto any photo to swap it before generating
- **White-Label Branding** -- Choose from preset branding (HKPC, Metapeller) or upload custom header/footer images

### Feedback System
- **Floating Feedback Widget** -- Public-facing widget on every page for submitting bugs, feature requests, and UI/UX feedback with file attachments
- **Progress Board** (`/progress`) -- Public board showing all feedback items grouped by status (Completed, In Progress, Upcoming) with overall progress bar
- **Admin Dashboard** (`/feedback`) -- Login-protected admin panel to view, manage, and update feedback status
- **Admin Tick-Off** -- Logged-in admins see checkbox buttons on the progress board to toggle item status via AJAX

### Infrastructure
- **CI/CD** -- GitHub Actions auto-deploys to EC2 on every push to `main`
- **Docker Support** -- Dockerfile included for containerised deployment

### Security
- **Environment-driven secrets** -- Flask secret key and admin passwords are read from environment variables (no secrets hardcoded in source)
- **CSRF protection** -- Flask-WTF protects all authenticated form/AJAX actions (login, feedback status update, delete). The two public endpoints (report generation, feedback submission) are exempt by design

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Backend | Flask (Python 3.12) |
| DOCX Generation | python-docx |
| Image Processing | Pillow |
| Database | SQLite via Flask-SQLAlchemy |
| Auth | Flask-Login (sessions) |
| CSRF | Flask-WTF |
| Frontend | Jinja2 templates, Bootstrap Icons, vanilla JS |
| Deployment | AWS EC2 (t3.micro), systemd |
| CI/CD | GitHub Actions (`appleboy/ssh-action`) |

## Project Structure

```
main.py                  # Flask app, routes, DOCX generator
feedback.py              # Feedback Blueprint (models, routes, auth)
templates/
  base.html              # Shared layout
  index.html             # Upload page with preview grid + labels
  progress.html          # Public progress board
  login.html             # Admin login
  feedback_list.html     # Admin feedback list
  feedback_detail.html   # Admin feedback detail
BRANDING_PRESETS/        # Header/footer images per branding preset
.github/workflows/
  deploy.yml             # CI/CD pipeline
Dockerfile               # Container build
requirements.txt         # Python dependencies
```

## Quick Start

### Docker

```bash
docker build --no-cache -t docx_generator .
docker run -p 11312:11312 docx_generator
```

### Local Development

```bash
pip install -r requirements.txt
python main.py
```

App runs on `http://localhost:11312`.

### Admin Login

Two admin accounts (`eric`, `gary`) are seeded on the **first run only** (when the
user table is empty). Their passwords come from environment variables — see
Configuration below. If no variable is set, a fallback default is used; set the
variables before first run (or rotate later) so credentials are never hardcoded.

## Configuration

All secrets are read from environment variables. None are required to run locally,
but they should be set in production.

| Variable | Purpose | Default if unset |
|----------|---------|------------------|
| `FLASK_SECRET_KEY` | Session signing key | Random per process (admins re-login after each restart) |
| `ADMIN_PASSWORD` | Shared fallback password for seeded admins | `tdu7101` |
| `ERIC_PASSWORD` | Password for the `eric` admin (overrides `ADMIN_PASSWORD`) | falls back to `ADMIN_PASSWORD` |
| `GARY_PASSWORD` | Password for the `gary` admin (overrides `ADMIN_PASSWORD`) | falls back to `ADMIN_PASSWORD` |

> Admin passwords are only applied when accounts are first seeded (empty database).
> On an already-seeded deployment, rotate passwords separately — setting the variable
> alone will not change existing stored hashes.

Example (production):

```bash
export FLASK_SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_hex(32))')"
export ERIC_PASSWORD="..." GARY_PASSWORD="..."
python main.py
```

## CI/CD

Pushes to `main` trigger automatic deployment to EC2:

1. SSH to EC2 via deploy key
2. `git fetch origin main && git reset --hard origin/main`
3. `pip install -r requirements.txt`
4. Restart systemd service
5. Health check (`systemctl is-active`)

## API Endpoints

| Method | Route | Auth | Description |
|--------|-------|------|-------------|
| GET | `/` | No | Upload page |
| POST | `/` | No | Generate DOCX report |
| GET | `/demo/download/<preset>` | No | Download demo report |
| GET | `/progress` | No | Public progress board |
| POST | `/feedback/submit` | No | Submit feedback (JSON) |
| GET | `/feedback` | Admin | Feedback list |
| GET | `/feedback/<id>` | Admin | Feedback detail |
| POST | `/feedback/<id>/status` | Admin | Update status |
| POST | `/feedback/<id>/delete` | Admin | Delete feedback |
| GET | `/login` | No | Login page |
| GET | `/logout` | Admin | Logout |
