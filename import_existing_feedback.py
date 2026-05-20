"""One-time script to import existing feedback from feedback_pool/grace/ into the database."""
import os
import shutil
from datetime import datetime

# Must be run from within the project directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from main import app
from feedback import db, Feedback, FeedbackFile

FEEDBACK_POOL = os.path.expanduser(
    "~/Desktop/project_202605/feedback_pool/grace"
)

EXISTING_FEEDBACK = [
    {
        "folder": "202605051103",
        "timestamp": datetime(2026, 5, 5, 11, 3),
        "author": "Grace",
        "category": "bug",
        "title": "Header and Footer duplicated on word template",
        "description": (
            'Header and Footer duplicated on the word template. '
            'Should put "Powered by METAPELLER LIMITED" default at the footer (compulsory). '
            'Can user edit the header and footer by themselves, on top of the default one? '
            'Remove the SurveyDoc AI words.'
        ),
        "status": "done",
        "files": [
            "Screenshot 2026-05-05 at 12.02.17 PM.png",
            "Screenshot 2026-05-05 at 11.57.00 AM.png",
            "Screenshot 2026-05-05 at 11.55.22 AM.png",
        ],
    },
    {
        "folder": "202605051252",
        "timestamp": datetime(2026, 5, 5, 12, 52),
        "author": "Grace",
        "category": "general",
        "title": "Letterhead template reference file",
        "description": "Provided a sample letterhead DOCX for branding reference.",
        "status": "done",
        "files": ["Letter heads 1.docx"],
    },
    {
        "folder": "202605051600",
        "timestamp": datetime(2026, 5, 5, 16, 0),
        "author": "Grace",
        "category": "bug",
        "title": "Header/footer images included in content table",
        "description": (
            "Bug: header footer images also included in the content table. "
            "Need to exclude the header and footer images."
        ),
        "status": "done",
        "files": ["Screenshot 2026-05-05 at 3.34.37 PM.png"],
    },
]


def import_feedback():
    upload_dir = app.config["UPLOAD_FOLDER"]
    os.makedirs(upload_dir, exist_ok=True)

    with app.app_context():
        for item in EXISTING_FEEDBACK:
            # Check if already imported (by title match)
            existing = Feedback.query.filter_by(title=item["title"]).first()
            if existing:
                print(f"  Skipping (already exists): {item['title']}")
                continue

            fb = Feedback(
                author=item["author"],
                category=item["category"],
                title=item["title"],
                description=item["description"],
                status=item["status"],
                page_context="/",
                created_at=item["timestamp"],
                updated_at=item["timestamp"],
            )
            db.session.add(fb)
            db.session.flush()  # get the id

            folder_path = os.path.join(FEEDBACK_POOL, item["folder"])
            for fname in item["files"]:
                src = os.path.join(folder_path, fname)
                if not os.path.exists(src):
                    print(f"  WARNING: File not found: {src}")
                    continue
                stored_name = f"{fb.id}_{fname}"
                dest = os.path.join(upload_dir, stored_name)
                shutil.copy2(src, dest)
                file_size = os.path.getsize(dest)
                ff = FeedbackFile(
                    feedback_id=fb.id,
                    stored_name=stored_name,
                    original_name=fname,
                    file_size=file_size,
                )
                db.session.add(ff)
                print(f"  Imported file: {fname} ({file_size // 1024} KB)")

            db.session.commit()
            print(f"Imported: [{item['status'].upper()}] {item['title']}")

    print("\nDone. All existing feedback imported.")


if __name__ == "__main__":
    import_feedback()
