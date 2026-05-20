import os
from uuid import uuid4
from datetime import datetime
from functools import wraps

from flask import (
    Blueprint, request, render_template, redirect, url_for,
    flash, jsonify, send_from_directory, current_app, abort
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    UserMixin, login_user, logout_user, login_required, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

db = SQLAlchemy()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf', 'docx', 'xlsx', 'txt', 'md', 'pptx'}
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB

STATUS_LABELS = {
    'new': 'New',
    'in_progress': 'In Progress',
    'done': 'Done',
}

CATEGORY_LABELS = {
    'bug': 'Bug',
    'feature': 'Feature',
    'ui_ux': 'UI/UX',
    'general': 'General',
}

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

class User(UserMixin, db.Model):
    __tablename__ = 'user'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    display_name = db.Column(db.String(100), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='admin')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_admin(self):
        return self.role == 'admin'


class Feedback(db.Model):
    __tablename__ = 'feedback'

    id = db.Column(db.Integer, primary_key=True)
    author = db.Column(db.String(100), nullable=False)
    category = db.Column(db.String(20), nullable=False, default='general')
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default='new')
    page_context = db.Column(db.String(200), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    files = db.relationship('FeedbackFile', backref='feedback',
                            cascade='all, delete-orphan', order_by='FeedbackFile.id')

    @property
    def status_label(self):
        return STATUS_LABELS.get(self.status, self.status)

    @property
    def category_label(self):
        return CATEGORY_LABELS.get(self.category, self.category)


class FeedbackFile(db.Model):
    __tablename__ = 'feedback_file'

    id = db.Column(db.Integer, primary_key=True)
    feedback_id = db.Column(db.Integer, db.ForeignKey('feedback.id'), nullable=False)
    stored_name = db.Column(db.String(255), nullable=False)
    original_name = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.Integer, default=0)

    @property
    def is_image(self):
        ext = self.original_name.rsplit('.', 1)[-1].lower() if '.' in self.original_name else ''
        return ext in ('png', 'jpg', 'jpeg', 'gif')


# ---------------------------------------------------------------------------
# Seed
# ---------------------------------------------------------------------------

def seed_users():
    """Create default admin users if none exist."""
    if User.query.first() is not None:
        return
    default_users = [
        {'username': 'eric', 'display_name': 'Eric', 'role': 'admin', 'password': 'tdu7101'},
        {'username': 'gary', 'display_name': 'Gary', 'role': 'admin', 'password': 'tdu7101'},
    ]
    for u in default_users:
        user = User(username=u['username'], display_name=u['display_name'], role=u['role'])
        user.set_password(u['password'])
        db.session.add(user)
    db.session.commit()
    print("  Seeded default admin users: eric, gary")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def admin_required(f):
    """Decorator: must be logged in AND have admin role."""
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Blueprint
# ---------------------------------------------------------------------------
feedback_bp = Blueprint('feedback', __name__)


@feedback_bp.app_context_processor
def inject_feedback_globals():
    """Inject is_admin and new_feedback_count into every template."""
    is_admin = current_user.is_authenticated and current_user.is_admin
    new_count = 0
    if is_admin:
        new_count = Feedback.query.filter_by(status='new').count()
    return dict(is_admin=is_admin, new_feedback_count=new_count)


# --- Auth routes ---

@feedback_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('feedback.feedback_list'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for('feedback.feedback_list'))
        flash('Invalid username or password.', 'error')
        return redirect(url_for('feedback.login'))
    return render_template('login.html')


@feedback_bp.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))


# --- Public routes ---

@feedback_bp.route('/progress')
def progress():
    """Public progress page -- no login required."""
    feedbacks = Feedback.query.order_by(Feedback.created_at.desc()).all()
    stats = {
        'total': len(feedbacks),
        'new': sum(1 for f in feedbacks if f.status == 'new'),
        'in_progress': sum(1 for f in feedbacks if f.status == 'in_progress'),
        'done': sum(1 for f in feedbacks if f.status == 'done'),
    }
    return render_template('progress.html', feedbacks=feedbacks, stats=stats,
                           STATUS_LABELS=STATUS_LABELS, CATEGORY_LABELS=CATEGORY_LABELS)


@feedback_bp.route('/feedback/submit', methods=['POST'])
def feedback_submit():
    """Public submission endpoint (no login). Returns JSON."""
    name = request.form.get('name', '').strip()
    category = request.form.get('category', 'general')
    title = request.form.get('title', '').strip()
    description = request.form.get('description', '').strip()
    page_context = request.form.get('page_context', '').strip()

    if not name or not title:
        return jsonify({'success': False, 'error': 'Name and title are required.'}), 400

    if category not in CATEGORY_LABELS:
        category = 'general'

    fb = Feedback(
        author=name,
        category=category,
        title=title,
        description=description or None,
        page_context=page_context or None,
    )
    db.session.add(fb)
    db.session.flush()  # get fb.id before handling files

    upload_folder = current_app.config.get('UPLOAD_FOLDER', 'uploads')
    os.makedirs(upload_folder, exist_ok=True)

    uploaded_files = request.files.getlist('files')
    for f in uploaded_files:
        if f and f.filename and allowed_file(f.filename):
            original = secure_filename(f.filename)
            prefix = uuid4().hex[:8]
            stored = f"{prefix}_{original}"
            filepath = os.path.join(upload_folder, stored)
            f.save(filepath)
            fsize = os.path.getsize(filepath)
            if fsize > MAX_FILE_SIZE:
                os.remove(filepath)
                continue
            fb_file = FeedbackFile(
                feedback_id=fb.id,
                stored_name=stored,
                original_name=original,
                file_size=fsize,
            )
            db.session.add(fb_file)

    db.session.commit()
    return jsonify({'success': True, 'id': fb.id})


# --- Admin routes ---

@feedback_bp.route('/feedback')
@admin_required
def feedback_list():
    """Admin feedback list with optional status filter."""
    status_filter = request.args.get('status')
    query = Feedback.query
    if status_filter and status_filter in STATUS_LABELS:
        query = query.filter_by(status=status_filter)
    feedbacks = query.order_by(Feedback.created_at.desc()).all()
    return render_template('feedback_list.html', feedbacks=feedbacks,
                           status_filter=status_filter,
                           STATUS_LABELS=STATUS_LABELS,
                           CATEGORY_LABELS=CATEGORY_LABELS)


@feedback_bp.route('/feedback/<int:feedback_id>')
@admin_required
def feedback_detail(feedback_id):
    """Admin detail view for a single feedback item."""
    fb = Feedback.query.get_or_404(feedback_id)
    return render_template('feedback_detail.html', feedback=fb,
                           STATUS_LABELS=STATUS_LABELS,
                           CATEGORY_LABELS=CATEGORY_LABELS)


@feedback_bp.route('/feedback/<int:feedback_id>/status', methods=['POST'])
@admin_required
def feedback_update_status(feedback_id):
    """Admin: update feedback status."""
    fb = Feedback.query.get_or_404(feedback_id)
    new_status = request.form.get('status', '')
    if new_status in STATUS_LABELS:
        fb.status = new_status
        fb.updated_at = datetime.utcnow()
        db.session.commit()
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify({'success': True, 'status': fb.status})
    return redirect(url_for('feedback.feedback_detail', feedback_id=fb.id))


@feedback_bp.route('/feedback/<int:feedback_id>/delete', methods=['POST'])
@admin_required
def feedback_delete(feedback_id):
    """Admin: delete feedback and its uploaded files."""
    fb = Feedback.query.get_or_404(feedback_id)
    upload_folder = current_app.config.get('UPLOAD_FOLDER', 'uploads')
    for f in fb.files:
        filepath = os.path.join(upload_folder, f.stored_name)
        if os.path.exists(filepath):
            os.remove(filepath)
    db.session.delete(fb)
    db.session.commit()
    return redirect(url_for('feedback.feedback_list'))


@feedback_bp.route('/uploads/<filename>')
@admin_required
def uploaded_file(filename):
    """Serve an uploaded file (admin only)."""
    upload_folder = current_app.config.get('UPLOAD_FOLDER', 'uploads')
    return send_from_directory(upload_folder, filename)
