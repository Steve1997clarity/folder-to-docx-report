# Changelog

All notable changes to the DOCX Report Generator.

## [2026-05-20] Phase 5 -- Photo Preview Grid with Drag-to-Replace

- Added 3-column thumbnail preview grid after file selection
- Each preview card shows photo number, filename label, and drop overlay
- Drag a JPG onto any card to replace that photo in-place
- Form submission rebuilds FormData from managed file array with numeric prefixes to preserve order
- Backend label lookup strips numeric prefix for correct photo-label mapping

## [2026-05-20] Phase 4 -- Photo Metadata Dialogue Box

- Added metadata panel that auto-extracts folder structure from `webkitRelativePath`
- Four fields: Master Folder, Folder, Sub-Folder, Path (auto-populated, editable via Edit button)
- Backend inserts a 4-row x 2-col metadata table at the top of the DOCX (bold labels, 9pt font)
- Metadata fields submitted as regular form fields, parsed in POST handler

## [2026-05-20] Phase 3 -- Sequential Photo Numbering with Editable Labels

- Photos auto-numbered as Photo 1_01, Photo 1_02, etc. on file selection
- Editable label inputs for each photo, displayed in a scrollable panel
- Form submission converted from standard POST to AJAX `fetch()` with blob download
- Labels collected as JSON, sent via FormData, used in DOCX label rows
- Backend accepts `photo_labels` parameter, defaults to filename when no label provided

## [2026-05-20] Phase 2 -- 3x4 Grid Layout

- Changed photo grid from 2 columns to 3 columns (3x4 = 12 photos per page)
- Image width set to 2.2 inches (fits 3 across 7.5" usable width)
- Image resize reduced from 600px to 450px max width (sufficient for 2.2" print)
- Label font reduced to 8pt, row spacing to 6pt for tighter grid

## [2026-05-20] Phase 1 -- Footer Fix + Admin Tick-Off

- Removed "Powered by METAPELLER LIMITED" text from DOCX footer
- Removed "All reports include Powered by METAPELLER LIMITED" from upload page UI
- Added admin tick-off checkboxes to progress board (visible when logged in)
- Done items show filled green check; click to reopen
- Non-done items show empty circle; click to mark as done
- AJAX toggle calls `/feedback/<id>/status` endpoint, reloads page on success
- Backend returns JSON for AJAX requests (checks `X-Requested-With` header)

## [2026-05-19] Structured Feedback System

- Added feedback Blueprint with SQLAlchemy models (User, Feedback, FeedbackFile)
- Floating feedback widget on every page (public submission, no login required)
- Public progress board at `/progress` with status grouping and progress bar
- Admin dashboard at `/feedback` with login, detail views, status management
- File attachment support (images, PDFs, documents up to 20MB)
- Default admin accounts seeded on first run (eric, gary)
- CI/CD pipeline via GitHub Actions for auto-deploy to EC2 on push

## [2026-05-18] Branding and UI Improvements

- Added HKPC as default branding preset
- Custom header/footer upload support
- Metapeller branding preset
- Excluded custom branding images from report content table
- Fixed header/footer duplication bug

## [2026-05-17] Initial Release

- Flask DOCX Report Generator with photo grid layout
- Demo presets with pre-generated reports
- Image resizing and compression for smaller output
- Docker support
- Landing page with upload form and live demo section
