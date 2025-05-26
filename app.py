import os
import json
from reportlab.lib.units import inch
import shutil
import io
import base64
import re
import requests as http_requests
from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file, Response, session, abort
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload, MediaFileUpload
from datetime import datetime
import pdfkit
from werkzeug.utils import secure_filename
import tempfile
from io import BytesIO
from PyPDF2 import PdfMerger
import logging
from google.auth.transport.requests import Request
from jinja2 import Environment, FileSystemLoader

# Import PDF stamping libraries
try:
    from pdfrw import PdfReader, PdfWriter, PageMerge
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.colors import HexColor, black, grey
    from reportlab.lib.units import inch, cm
    from PIL import Image
except ImportError:
    PdfReader = PdfWriter = PageMerge = canvas = ImageReader = letter = Image = HexColor = black = grey = inch = cm = http_requests = None
    logging.error(
        "PDF stamping libraries (pdfrw, reportlab, Pillow, requests) not installed. Stamping feature unavailable.")

# Optional PDF generation library fallback
try:
    from weasyprint import HTML
except ImportError:
    HTML = None
    logging.warning("weasyprint not installed. PDF generation might fail if wkhtmltopdf is not found.")

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

app = Flask(__name__)
app.secret_key = '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'

# Configuration
GOOGLE_DRIVE_FOLDER_ID = "1B62aXBQwjH8-ZOwWfJt-d5YPiq7Z2exN"
GOOGLE_SHEETS_SPREADSHEET_ID = "1tGqhzBaEwGEh9Vq7GW-toGiO-hW25Sc4sjv4Nct9SqQ"

REQUESTER_SIGNATURE_IMAGE_URL = "https://i.ibb.co/B0t7zn7/handwritten-signature-high-quality.png"
CEO_SIGNATURE_IMAGE_URL = "https://i.ibb.co/YJwX7mm/images.jpg"
STANDARD_SIGNATURE_IMAGE_URL = "https://i.ibb.co/fVNbgBVD/make-signature-hero.jpg"
CEO_APPROVER_NAME = "Sir Qaiser"
STANDARD_APPROVER_NAME = "Amir Saddique"

# --- NEW: Separate Username/Passwords ---
CEO_USERNAME = "ceo_user"
CEO_PASSWORD = "ceo_password"

STANDARD_USERNAME = "standard_user"
STANDARD_PASSWORD = "standard_password"

DASHBOARD_USERNAME = "dashboard_user"
DASHBOARD_PASSWORD = "dashboard_password"

PORTAL_USERS = {
    DASHBOARD_USERNAME: DASHBOARD_PASSWORD,
    STANDARD_USERNAME: STANDARD_PASSWORD,
    CEO_USERNAME: CEO_PASSWORD
}
# --- END NEW ---

# Company Logos
COMPANY_LOGOS = {
    "Bpro": "https://i.ibb.co/R4739SnZ/bpro-ai-logo.jpg",
    "DS": "https://i.ibb.co/tMhLwFN8/images.jpg",
    "ML-1": "https://i.ibb.co/tPTmj2YB/machine-learning-1-logo.jpg"
}

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'gif', 'doc', 'docx'}
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets']

if letter:
    PAGE_WIDTH_POINTS, PAGE_HEIGHT_POINTS = letter
else:
    PAGE_WIDTH_POINTS, PAGE_HEIGHT_POINTS = 595, 842  # Default A4 width/height in points

HTML_REQUESTER_BOX_WIDTH_PERCENT = 42
HTML_APPROVER_BOX_WIDTH_PERCENT = 48
HTML_BOX_MIN_HEIGHT_PT = 125
HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT = 60
HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT = 150
APPROVER_HTML_PLACEHOLDER_HEIGHT_PT = HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT
APPROVER_HTML_PLACEHOLDER_WIDTH_PT = HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT

STAMP_PAGE_INDEX = 0

APPROVER_STAMP_SECTION_WIDTH_PT = 220
APPROVER_STAMP_SECTION_HEIGHT_PT = 110

if letter:
    APPROVER_STAMP_SECTION_X_PT = (PAGE_WIDTH_POINTS * 0.48) + 5
else:
    APPROVER_STAMP_SECTION_X_PT = 350

APPROVER_STAMP_SECTION_Y_PT = 2.5 * inch

SIGNATURE_IMAGE_HEIGHT_IN_STAMP_AREA = 0.6 * inch
TEXT_LINE_HEIGHT = 12
NAME_TEXT_OFFSET_Y_IN_STAMP = 10 + TEXT_LINE_HEIGHT
DATE_TEXT_OFFSET_Y_IN_STAMP = 10

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

import sys

if sys.platform == "win32":
    WKHTMLTOPDF_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
else:
    WKHTMLTOPDF_PATH = '/usr/local/bin/wkhtmltopdf'

PDFKIT_CONFIG = None
if WKHTMLTOPDF_PATH and os.path.exists(WKHTMLTOPDF_PATH) and os.access(WKHTMLTOPDF_PATH, os.X_OK):
    try:
        PDFKIT_CONFIG = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        logging.info(f"pdfkit configured successfully with {WKHTMLTOPDF_PATH}")
    except Exception as e:
        logging.error(f"Error configuring pdfkit: {e}. Falling back to weasyprint if available.", exc_info=True);
        PDFKIT_CONFIG = None
else:
    logging.warning(
        f"wkhtmltopdf not found or not executable at {WKHTMLTOPDF_PATH}. Falling back to weasyprint if available.")
    PDFKIT_CONFIG = None

if not PDFKIT_CONFIG and HTML is None: logging.error(
    "Neither wkhtmltopdf nor weasyprint are available. PDF generation will not work.")

PDF_STAMPING_AVAILABLE = True
if PdfReader is None or PdfWriter is None or PageMerge is None or canvas is None or ImageReader is None or letter is None or Image is None or http_requests is None:
    logging.error(
        "PDF stamping libraries (pdfrw, reportlab, Pillow, requests) not fully available. Stamping will not work.")
    PDF_STAMPING_AVAILABLE = False


# --- Authentication Functions ---
def check_user_auth(portal_type):
    if f'{portal_type}_authenticated' in session:
        return True
    return False


def authenticate_user(username, password, portal_type):
    if portal_type == 'ceo' and username == CEO_USERNAME and password == CEO_PASSWORD:
        session['ceo_authenticated'] = True
        session['current_username'] = CEO_USERNAME
        return True
    elif portal_type == 'standard' and username == STANDARD_USERNAME and password == STANDARD_PASSWORD:
        session['standard_authenticated'] = True
        session['current_username'] = STANDARD_USERNAME
        return True
    elif portal_type == 'dashboard' and username == DASHBOARD_USERNAME and password == DASHBOARD_PASSWORD:
        session['dashboard_authenticated'] = True
        session['current_username'] = DASHBOARD_USERNAME
        return True
    return False


def require_auth(portal_type):
    def decorator(f):
        from functools import wraps
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'credentials' not in session:
                return redirect(url_for('authorize'))
            if not check_user_auth(portal_type):
                if portal_type == 'dashboard':
                    return redirect(url_for('dashboard_login'))
                elif portal_type == 'standard':
                    return redirect(url_for('standard_login'))
                elif portal_type == 'ceo':
                    return redirect(url_for('ceo_login'))
            return f(*args, **kwargs)

        return decorated_function

    return decorator


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_google_auth_flow():
    client_secrets_path = 'credentials.json'
    if not os.path.exists(client_secrets_path):
        logging.error(f"{client_secrets_path} not found. OAuth flow cannot be initialized.")
    return Flow.from_client_secrets_file(
        client_secrets_path,
        scopes=SCOPES,
        redirect_uri=url_for('oauth2callback', _external=True)
    )


def get_credentials():
    if 'credentials' not in session: return None
    try:
        creds_data = json.loads(session['credentials'])
        creds = Credentials(token=creds_data.get('token'), refresh_token=creds_data.get('refresh_token'),
                            token_uri=creds_data.get('token_uri'), client_id=creds_data.get('client_id'),
                            client_secret=creds_data.get('client_secret'), scopes=creds_data.get('scopes'))
        if creds.expired and creds.refresh_token:
            logging.info("Credentials expired, attempting refresh...");
            creds.refresh(Request())
            session['credentials'] = json.dumps({
                'token': creds.token, 'refresh_token': creds.refresh_token, 'token_uri': creds.token_uri,
                'client_id': creds.client_id, 'client_secret': creds.client_secret, 'scopes': creds.scopes})
            logging.info("Credentials refreshed successfully.")
        elif creds.expired and not creds.refresh_token:
            logging.warning("Credentials expired, no refresh token. User needs re-authorization.");
            session.pop('credentials', None);
            return None
        return creds
    except Exception as e:
        logging.error(f"Error loading/refreshing credentials: {e}", exc_info=True);
        session.pop('credentials', None);
        return None


def upload_file_from_path(file_path, file_name, mime_type):
    creds = get_credentials()
    if not creds: logging.error("Cannot upload from path: No credentials."); return None
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': file_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
        logging.info(f"Uploading from path: {file_path} to Drive folder: {GOOGLE_DRIVE_FOLDER_ID}")
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media,
                                                     fields='id,webViewLink').execute()
        file_id = uploaded_file.get('id');
        web_view_link = uploaded_file.get('webViewLink')
        if not file_id: logging.error(f"Drive file creation failed (path): {file_name}"); return None
        try:
            permission = {'type': 'anyone', 'role': 'reader'}
            drive_service.permissions().create(fileId=file_id, body=permission, fields='id').execute()
            logging.info(f"Set public permission for file ID: {file_id}")
        except HttpError as error:
            logging.warning(f"Could not set public permission for {file_id}: {error}", exc_info=True)
        logging.info(f"Uploaded {file_name} from path. ID: {file_id}, Link: {web_view_link}")
        return web_view_link
    except Exception as e:
        logging.error(f"Error uploading from path: {e}", exc_info=True);
        return None


def upload_file_from_bytes(file_content, file_name, mime_type, file_id_to_update=None):
    creds = get_credentials()
    if not creds: logging.error("Cannot upload from bytes: No credentials."); return None
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        media = MediaIoBaseUpload(io.BytesIO(file_content), mimetype=mime_type, resumable=True)
        if file_id_to_update:
            logging.info(f"Updating Drive file ID: {file_id_to_update}...")
            updated_file = drive_service.files().update(fileId=file_id_to_update, media_body=media,
                                                        fields='id,webViewLink,name').execute()
            file_id = updated_file.get('id');
            web_view_link = updated_file.get('webViewLink');
            updated_name = updated_file.get('name')
            logging.info(f"Drive file ID {file_id_to_update} updated. Name: {updated_name}, Link: {web_view_link}")
        else:
            file_metadata = {'name': file_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
            logging.info(
                f"Creating new Drive file: {file_name} (MIME: {mime_type}) in folder: {GOOGLE_DRIVE_FOLDER_ID}")
            uploaded_file = drive_service.files().create(body=file_metadata, media_body=media,
                                                         fields='id,webViewLink').execute()
            file_id = uploaded_file.get('id');
            web_view_link = uploaded_file.get('webViewLink')
            if not file_id: logging.error(f"Drive file creation failed (bytes): {file_name}"); return None
            try:
                permission = {'type': 'anyone', 'role': 'reader'}
                drive_service.permissions().create(fileId=file_id, body=permission, fields='id').execute()
                logging.info(f"Set public permission for new file ID: {file_id}")
            except HttpError as error:
                logging.warning(f"Could not set public permission for {file_id}: {error}", exc_info=True)
        logging.info(f"Uploaded/Updated file from bytes. ID: {file_id}, Link: {web_view_link}")
        return web_view_link
    except HttpError as e:
        logging.error(f"Google API HttpError during file upload/update: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        return None
    except Exception as e:
        logging.error(f"Error uploading/updating from bytes: {e}", exc_info=True);
        return None


def get_signature_data_from_url(image_url):
    if http_requests is None or Image is None:
        return None, None, "Libraries for downloading/processing images not available (requests/Pillow)."
    if not image_url:
        return None, None, "Signature image URL is missing."
    try:
        logging.info(f"Downloading signature image from URL: {image_url}")
        response = http_requests.get(image_url, stream=True, timeout=10)
        response.raise_for_status()
        image_stream = io.BytesIO(response.content)
        image_stream.seek(0)
        try:
            img = Image.open(image_stream)
            mime_type = f"image/{img.format.lower()}" if img.format else 'image/png'  # Default to png
            image_stream.seek(0)  # Reset stream for reading again
        except Exception as img_id_e:
            logging.warning(f"Could not identify image format from URL {image_url} ({img_id_e}). Guessing mime type.",
                            exc_info=True)
            mime_type = response.headers.get('Content-Type', 'application/octet-stream')
            if 'image/' not in mime_type:  # Fallback to extension if content-type isn't specific
                ext = os.path.splitext(image_url)[1].lower()
                mime_map = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif',
                            '.webp': 'image/webp'}
                mime_type = mime_map.get(ext, 'application/octet-stream')
            if 'image/' not in mime_type:
                raise ValueError(f"Could not determine valid image mime type for {image_url}. Guessed: {mime_type}")

        image_bytes = image_stream.read()
        base64_string = base64.b64encode(image_bytes).decode('utf-8')
        logging.info(f"Successfully fetched and encoded signature data from URL {image_url}.")
        return base64_string, mime_type, None
    except http_requests.exceptions.RequestException as req_e:
        error_msg = f"HTTP request error fetching signature from URL {image_url}: {req_e}"
        logging.error(error_msg, exc_info=True)
        return None, None, error_msg
    except Exception as e:
        error_msg = f"An unexpected error occurred fetching signature from URL {image_url}: {e}"
        logging.error(error_msg, exc_info=True)
        return None, None, error_msg


def stamp_pdf_with_signature(original_pdf_bytes, signature_base64, signature_mime_type, approver_name,
                             approval_date_str, approval_type_heading, page=0):
    if not PDF_STAMPING_AVAILABLE:
        logging.error("PDF stamping libraries not available. Cannot stamp PDF.")
        return None, "PDF stamping feature unavailable (libraries missing)."
    if not original_pdf_bytes: return None, "Missing original PDF bytes."
    if signature_base64 and (not signature_mime_type or not signature_mime_type.startswith('image/')):
        return None, f"Signature file is not an image ({signature_mime_type})"

    try:
        logging.info(f"Starting PDF stamping process for page {page} with approval type '{approval_type_heading}'...")
        original_pdf = PdfReader(io.BytesIO(original_pdf_bytes))
        if page >= len(original_pdf.pages):
            return None, f"Invalid page index {page} (PDF has {len(original_pdf.pages)} pages)."

        target_page = original_pdf.pages[page]
        try:
            # MediaBox is [lower_left_x, lower_left_y, upper_right_x, upper_right_y]
            page_width = float(target_page.MediaBox[2]) - float(target_page.MediaBox[0])
            page_height = float(target_page.MediaBox[3]) - float(target_page.MediaBox[1])
        except:
            page_width, page_height = letter  # Fallback

        section_x = APPROVER_STAMP_SECTION_X_PT
        section_y = APPROVER_STAMP_SECTION_Y_PT
        section_w = APPROVER_STAMP_SECTION_WIDTH_PT
        section_h = APPROVER_STAMP_SECTION_HEIGHT_PT

        overlay_bytes = io.BytesIO()
        c = canvas.Canvas(overlay_bytes, pagesize=(page_width, page_height))

        c.setStrokeColor(grey)
        c.setFillColor(HexColor("#f9f9f9"))
        c.rect(section_x, section_y, section_w, section_h, stroke=1, fill=1)
        c.setFillColor(black)

        heading_y_pos = section_y + section_h - (0.25 * inch)
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(section_x + section_w / 2, heading_y_pos, approval_type_heading)

        text_base_y = heading_y_pos - (0.3 * inch)  # Initial position for text below heading
        if signature_base64 and signature_mime_type:
            try:
                image_data_stream = io.BytesIO(base64.b64decode(signature_base64))
                img = ImageReader(image_data_stream)
                img_width_orig, img_height_orig = img.getSize()

                scaled_h = SIGNATURE_IMAGE_HEIGHT_IN_STAMP_AREA
                scaled_w = img_width_orig * (scaled_h / img_height_orig) if img_height_orig > 0 else 0
                if scaled_w > section_w * 0.8:  # Cap width to 80% of section
                    scaled_w = section_w * 0.8
                    scaled_h = img_height_orig * (scaled_w / img_width_orig) if img_width_orig > 0 else 0

                img_x_draw = section_x + (section_w - scaled_w) / 2
                img_y_draw = heading_y_pos - scaled_h - (0.1 * inch)  # Place below heading
                c.drawImage(img, img_x_draw, img_y_draw, width=scaled_w, height=scaled_h, mask='auto')
                text_base_y = img_y_draw - (0.15 * inch)  # Adjust text base below image
            except Exception as img_e:
                logging.error(f"Error drawing signature image: {img_e}", exc_info=True)
                c.drawString(section_x + 0.1 * inch, section_y + section_h / 2, "Signature Error")
                text_base_y = section_y + section_h / 2 - (0.2 * inch)  # Fallback text position
        else:
            logging.info("No signature image data provided for stamping approver, only text will be added.")
            # text_base_y remains as initialized if no image

        c.setFont("Helvetica", 9)
        # Ensure text fits within the box, adjust Y positions carefully
        name_y_pos = text_base_y - TEXT_LINE_HEIGHT
        date_y_pos = name_y_pos - TEXT_LINE_HEIGHT

        # Prevent text from going below the stamp box bottom
        name_y_pos = max(section_y + DATE_TEXT_OFFSET_Y_IN_STAMP + TEXT_LINE_HEIGHT, name_y_pos)
        date_y_pos = max(section_y + DATE_TEXT_OFFSET_Y_IN_STAMP, date_y_pos)

        c.drawCentredString(section_x + section_w / 2, name_y_pos, f"Name: {approver_name}")
        c.drawCentredString(section_x + section_w / 2, date_y_pos,
                            f"Date: {approval_date_str.split(' ')[0]}")  # Show only date part

        c.save()
        overlay_bytes.seek(0)
        overlay_pdf = PdfReader(overlay_bytes)
        if not overlay_pdf.pages: return None, "Failed to create overlay PDF for stamping."

        PageMerge(target_page).add(overlay_pdf.pages[0]).render()

        writer = PdfWriter();
        writer.addpages(original_pdf.pages)
        output_pdf_stream = io.BytesIO();
        writer.write(output_pdf_stream);
        output_pdf_stream.seek(0)
        logging.info(f"PDF stamped successfully on page {page}.")
        return output_pdf_stream.getvalue(), None
    except Exception as e:
        logging.error(f"Unexpected error during PDF stamping: {e}", exc_info=True)
        return None, f"Unexpected error during PDF stamping: {e}"


def ensure_sheet_headers(sheets_service, spreadsheet_id):
    try:
        expected_headers = [
            "Request ID", "Timestamp", "Name", "Email", "Company Name",
            "Account Title", "Account Number", "IBAN Number", "Bank Name",
            "Payment Type", "Description", "Quantity", "Amount", "Currency",
            "Supporting Document Link", "Request PDF Link", "Status",
            "Approval Type", "Approval Date", "Rejection Reason",
            "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason"
        ]
        range_to_read = f"Sheet1!A1:{chr(ord('A') + len(expected_headers) - 1)}1"

        result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_to_read).execute()
        values = result.get('values', [])

        current_headers = values[0] if values else []

        headers_match = True
        if len(current_headers) != len(expected_headers):
            headers_match = False
        else:
            for i, header in enumerate(expected_headers):
                if i >= len(current_headers) or current_headers[i] != header:
                    headers_match = False
                    break

        if not headers_match:
            logging.info("Sheet headers missing or incorrect. Adding/Updating headers.")
            body = {'values': [expected_headers]}
            sheets_service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range="Sheet1!A1",
                                                          valueInputOption="RAW", body=body).execute()
            logging.info(f"Sheet headers updated to: {expected_headers}")
        else:
            logging.info("Sheet headers are already correct.")

    except HttpError as e:
        logging.error(f"Google API HttpError ensuring sheet headers: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
    except Exception as e:
        logging.error(f"Error ensuring sheet headers: {e}", exc_info=True)


def get_next_request_id(sheets_service, spreadsheet_id):
    try:
        result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Sheet1!A:A").execute()
        ids = result.get('values', [])
        num_requests = 0
        if len(ids) > 1:
            for row_val in ids[1:]:
                if row_val and row_val[0]:
                    num_requests += 1

        next_id_number = num_requests + 1
        next_request_id = f"{next_id_number:05d}"
        logging.info(f"Generated next Request ID: {next_request_id} (based on {num_requests} existing requests)")
        return next_request_id, None
    except HttpError as e:
        logging.error(f"Google API HttpError generating next Request ID: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        return None, f"API Error: {e._get_reason()}"
    except Exception as e:
        logging.error(f"Error generating next Request ID: {e}", exc_info=True);
        return None, "Failed to generate Request ID"


def add_to_sheet(data, pdf_link, attachment_link, status, approval_type):
    creds = get_credentials()
    if not creds: return False, "Authentication failed"
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID
        ensure_sheet_headers(sheets_service, spreadsheet_id)

        request_id = data.get('requestId')
        expected_columns = [
            "Request ID", "Timestamp", "Name", "Email", "Company Name",
            "Account Title", "Account Number", "IBAN Number", "Bank Name",
            "Payment Type", "Description", "Quantity", "Amount", "Currency",
            "Supporting Document Link", "Request PDF Link", "Status",
            "Approval Type", "Approval Date", "Rejection Reason",
            "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason"
        ]

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        data_dict = {
            "Request ID": request_id,
            "Timestamp": timestamp,
            "Name": data.get('name', ''),
            "Email": data.get('email', ''),
            "Company Name": data.get('companyName', ''),
            "Account Title": data.get('accountTitle', ''),
            "Account Number": data.get('accountNumber', ''),
            "IBAN Number": data.get('ibanNumber', ''),
            "Bank Name": data.get('bankName', ''),
            "Payment Type": data.get('paymentType', ''),
            "Description": data.get('description', ''),
            "Quantity": data.get('quantity', ''),
            "Amount": data.get('amount', ''),
            "Currency": data.get('currency', ''),
            "Supporting Document Link": attachment_link if attachment_link else "No attachment",
            "Request PDF Link": pdf_link if pdf_link else "Error generating PDF",
            "Status": status,
            "Approval Type": approval_type,
            "Approval Date": "",
            "Rejection Reason": "",
            "Voucher PDF Link": "",
            "Voucher Generated At": "",
            "Voucher Approved By": "",
            "Voucher Rejection Reason": ""
        }
        row = [data_dict.get(col, "") for col in expected_columns]
        body = {'values': [row]}

        logging.info(f"Appending new row to sheet {spreadsheet_id} with Request ID {request_id}...")
        result = sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range="Sheet1!A:A",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS", body=body
        ).execute()
        logging.info(
            f"Data added to sheet for Request ID: {request_id}. Cells updated: {result.get('updates', {}).get('updatedCells')}")
        return True, request_id
    except HttpError as e:
        logging.error(f"Google API HttpError adding to sheet: {e.resp.status} - {e._get_reason()}", exc_info=True)
        return False, f"API Error: {e._get_reason()}"
    except Exception as e:
        logging.error(f"Error adding to Google Sheet: {e}", exc_info=True);
        return False, str(e)


def update_sheet_status(request_id, status=None, approval_date=None, rejection_reason=None,
                        pdf_link=None, voucher_link=None, voucher_generated_at=None,
                        voucher_approved_by=None, voucher_rejection_reason=None):
    creds = get_credentials()
    if not creds: return False, "Authentication failed"
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID

        header_result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                   range="Sheet1!1:1").execute()
        headers = header_result.get('values', [[]])[0]
        if not headers:
            ensure_sheet_headers(sheets_service, spreadsheet_id)
            header_result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                       range="Sheet1!1:1").execute()
            headers = header_result.get('values', [[]])[0]  # FIX: This line was missing re-assignment of headers
            if not headers: return False, "Sheet headers not found even after attempting fix."

        id_column_data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                    range="Sheet1!A:A").execute()
        ids = id_column_data.get('values', [])
        row_index_to_update = -1
        if ids and len(ids) > 1:
            for i, row_val in enumerate(ids):
                if row_val and row_val[0] == request_id:
                    row_index_to_update = i
                    break

        if row_index_to_update == -1:
            return False, "Request ID not found in sheet."

        sheet_row_num = row_index_to_update + 1

        update_data = []

        def get_col_idx(header_name):
            try:
                return headers.index(header_name)
            except ValueError:
                logging.warning(f"Header '{header_name}' not found in sheet. Update will be skipped for this column.")
                return None

        def col_letter_from_0_idx(n_idx):
            string = ""
            n = n_idx + 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string

        if status is not None:
            col_idx = get_col_idx("Status")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[status]]
                })
        if approval_date is not None:
            col_idx = get_col_idx("Approval Date")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[approval_date]]
                })
        if rejection_reason is not None:
            col_idx = get_col_idx("Rejection Reason")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[rejection_reason]]
                })
        if pdf_link is not None:
            col_idx = get_col_idx("Request PDF Link")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[pdf_link]]
                })

        if voucher_link is not None:
            col_idx = get_col_idx("Voucher PDF Link")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[voucher_link]]
                })
        if voucher_generated_at is not None:
            col_idx = get_col_idx("Voucher Generated At")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[voucher_generated_at]]
                })
        if voucher_approved_by is not None:
            col_idx = get_col_idx("Voucher Approved By")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[voucher_approved_by]]
                })
        if voucher_rejection_reason is not None:
            col_idx = get_col_idx("Voucher Rejection Reason")
            if col_idx is not None:
                update_data.append({
                    'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                    'values': [[voucher_rejection_reason]]
                })

        if not update_data:
            return True, "No valid data provided for sheet update."

        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': update_data
        }

        logging.info(f"Batch updating sheet for Request ID {request_id} at sheet row {sheet_row_num}...")
        result = sheets_service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        logging.info(f"Sheet updated for Request ID {request_id}. Responses: {result.get('responses')}")
        return True, "Sheet updated successfully"

    except HttpError as e:
        logging.error(
            f"Google API HttpError updating sheet status for {request_id}: {e.resp.status} - {e._get_reason()}",
            exc_info=True)
        return False, f"API Error: {e._get_reason()}"
    except Exception as e:
        logging.error(f"Error updating sheet status for Request ID {request_id}: {e}", exc_info=True);
        return False, str(e)


# --- NEW: Function to get approver signatures/names from Sheet2 ---
def get_approver_signatures_from_sheet(sheets_service, spreadsheet_id):
    signatures_data = {
        'prepared_by_names': [],  # List of names for dropdown
        'finance_review_name_default': 'N/A',
        'finance_review_signature_url': '',
        'approved_by_name_default': 'N/A',
        'approved_by_signature_url': '',
        'prepared_by_signature_urls_map': {}  # To map selected name to its signature URL
    }
    try:
        # Assuming Sheet2 for signatures
        result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Sheet2!A:F").execute()
        values = result.get('values', [])

        if not values or len(values) < 2:
            logging.warning("Sheet2 for signatures is empty or missing headers. Please populate it correctly.")
            return signatures_data  # Return empty defaults

        headers = [h.strip() for h in values[0]]  # Clean headers by stripping whitespace
        header_map = {h: i for i, h in enumerate(headers)}

        # Define all possible column names with variations
        column_mappings = {
            'prepared_by_name': ['Prepared by Name', 'Prepared By Name', 'Prepared by', 'Prepared By'],
            'prepared_by_sig': ['Prepared by Signature URL', 'Prepared By Signature URL', 'Prepared by Signature', 'Prepared By Signature'],
            'finance_name': ['Finance Review Name', 'Finance Reviewer Name', 'Finance Name'],
            'finance_sig': ['Finance Review Signature', 'Finance Reviewer Signature', 'Finance Signature'],
            'approved_by_name': ['Approved By Name', 'Approver Name', 'Approved by Name'],
            'approved_by_sig': ['Approved By Signature', 'Approver Signature', 'Approved by Signature']
        }

        # Find actual column indices based on header variations
        col_indices = {}
        for key, possible_names in column_mappings.items():
            for name in possible_names:
                if name in header_map:
                    col_indices[key] = header_map[name]
                    break

        for row_data in values[1:]:  # Skip header row
            if not row_data:
                continue

            # Collect all Prepared By names and their signature URLs
            if 'prepared_by_name' in col_indices and col_indices['prepared_by_name'] < len(row_data):
                name = row_data[col_indices['prepared_by_name']].strip()
                if name and name not in signatures_data['prepared_by_names']:
                    signatures_data['prepared_by_names'].append(name)

                    if 'prepared_by_sig' in col_indices and col_indices['prepared_by_sig'] < len(row_data):
                        sig_url = row_data[col_indices['prepared_by_sig']].strip()
                        if sig_url:
                            signatures_data['prepared_by_signature_urls_map'][name] = sig_url

            # Get Finance Reviewer data (use first valid entry)
            if ('finance_name' in col_indices and
                col_indices['finance_name'] < len(row_data) and
                signatures_data['finance_review_name_default'] == 'N/A'):
                finance_name = row_data[col_indices['finance_name']].strip()
                if finance_name:
                    signatures_data['finance_review_name_default'] = finance_name
                    if ('finance_sig' in col_indices and
                        col_indices['finance_sig'] < len(row_data)):
                        finance_sig = row_data[col_indices['finance_sig']].strip()
                        if finance_sig:
                            signatures_data['finance_review_signature_url'] = finance_sig

            # Get Approved By data (use first valid entry)
            if ('approved_by_name' in col_indices and
                col_indices['approved_by_name'] < len(row_data) and
                signatures_data['approved_by_name_default'] == 'N/A'):
                approved_name = row_data[col_indices['approved_by_name']].strip()
                if approved_name:
                    signatures_data['approved_by_name_default'] = approved_name
                    if ('approved_by_sig' in col_indices and
                        col_indices['approved_by_sig'] < len(row_data)):
                        approved_sig = row_data[col_indices['approved_by_sig']].strip()
                        if approved_sig:
                            signatures_data['approved_by_signature_url'] = approved_sig

        logging.info(f"Fetched signatures data: {signatures_data}")
        return signatures_data

    except HttpError as e:
        logging.error(f"Google API HttpError fetching signatures from Sheet2: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        return signatures_data  # Return what we have even if error
    except Exception as e:
        logging.error(f"Error fetching signatures from Sheet2: {e}", exc_info=True)
        return signatures_data  # Return what we have even if error

    except HttpError as e:
        logging.error(f"Google API HttpError fetching signatures from Sheet2: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        return signatures_data  # Return what we have even if error
    except Exception as e:
        logging.error(f"Error fetching signatures from Sheet2: {e}", exc_info=True)
        return signatures_data  # Return what we have even if error# Return what we have even if error


# --- END NEW ---


def get_requests_from_sheet(status_filter=None):
    creds = get_credentials()
    if not creds: return None, "Authentication failed"
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID

        result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Sheet1!A:Z").execute()
        values = result.get('values', [])

        if not values: return [], None

        headers = values[0]
        header_map = {h.strip(): i for i, h in enumerate(headers)}

        expected_headers_for_read = [
            "Request ID", "Timestamp", "Name", "Email", "Company Name",
            "Account Title", "Account Number", "IBAN Number", "Bank Name",
            "Payment Type", "Description", "Quantity", "Amount", "Currency",
            "Supporting Document Link", "Request PDF Link", "Status",
            "Approval Type", "Approval Date", "Rejection Reason",
            "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason"
        ]

        requests_list = []
        for i, row_data in enumerate(values[1:]):
            if not row_data or not any(cell.strip() for cell in row_data):
                continue

            request_data = {}
            for header in expected_headers_for_read:
                index = header_map.get(header)
                request_data[header] = row_data[index].strip() if index is not None and index < len(row_data) else ""

            if not request_data.get("Request ID"):
                logging.warning(f"Skipping row {i + 2} due to missing Request ID: {row_data}")
                continue

            if status_filter is None or request_data.get("Status", "") == status_filter:
                requests_list.append(request_data)

        logging.info(
            f"Fetched {len(values[1:])} data rows. Filtered to {len(requests_list)} requests (Status filter: '{status_filter}').")
        return requests_list, None
    except HttpError as e:
        logging.error(f"Google API HttpError fetching from sheet: {e.resp.status} - {e._get_reason()}", exc_info=True)
        return None, f"API Error: {e._get_reason()}"
    except Exception as e:
        logging.error(f"Error fetching data from Google Sheet: {e}", exc_info=True);
        return None, str(e)


def get_request_by_id(request_id):
    logging.info(f"Attempting to fetch request with ID: {request_id}")
    requests_list, error = get_requests_from_sheet(status_filter=None)
    if error: return None, error
    if requests_list is None: return None, "Could not retrieve requests list from sheet."

    for req in requests_list:
        if req.get("Request ID") == request_id:
            logging.info(f"Found request with ID: {request_id}")
            return req, None

    logging.warning(f"Request ID {request_id} not found in sheet data.")
    return None, "Request not found"


def generate_pdf(data, approval_type, attachment_path=None):
    request_id_display = data.get('requestId', 'N/A')
    requester_name_from_form = data.get('name', 'N/A')
    request_date = datetime.now().strftime("%Y-%m-%d")

    requester_signature_html = ""
    if REQUESTER_SIGNATURE_IMAGE_URL and http_requests and Image:
        try:
            base64_sig, mime_type, err = get_signature_data_from_url(REQUESTER_SIGNATURE_IMAGE_URL)
            if base64_sig and mime_type:
                requester_signature_html = f'<img src="data:{mime_type};base64,{base64_sig}" alt="Requester Signature" style="max-height:{HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT}pt; max-width:{HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT}pt; object-fit:contain;">'
            else:
                logging.error(f"Failed to load requester signature image for initial PDF: {err}")
                requester_signature_html = "(Requester Signature Error)"
        except Exception as e:
            logging.error(f"Exception fetching requester signature for initial PDF: {e}", exc_info=True)
            requester_signature_html = "(Requester Signature Load Error)"
    else:
        requester_signature_html = "(Requester Signature Setup Incomplete)"

    # This part handles embedding of non-PDF attachments for preview purposes in HTML.
    attachment_html_content = ""
    if attachment_path and os.path.exists(attachment_path):
        if not attachment_path.lower().endswith('.pdf'):  # Only for non-PDF attachments for direct embedding
            try:
                with open(attachment_path, 'rb') as f:
                    encoded_string = base64.b64encode(f.read()).decode('utf-8')
                file_ext = os.path.splitext(attachment_path)[1].lower()
                mime_map = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif'}
                mime_type = mime_map.get(file_ext)
                if mime_type:
                    attachment_html_content = f"""<div class="page-break"><h3>Supporting Document Preview</h3><img class="attachment-preview" src="data:{mime_type};base64,{encoded_string}" alt="Attachment Preview"></div>"""
                else:
                    attachment_html_content = f"""<div class="page-break"><h3>Supporting Document</h3><p>A non-PDF, non-image supporting document ({file_ext.upper()}) is attached (preview not available here).</p></div>"""
            except Exception as e:
                logging.error(f"Error embedding attachment {attachment_path} in HTML: {e}", exc_info=True)
                attachment_html_content = """<div class="page-break"><h3>Supporting Document</h3><p>Error generating attachment preview.</p></div>"""

    html_template_str = """
 <!DOCTYPE html>
   <html>
   <head>
       <meta charset="UTF-8">
       <title>Request Form - {{ request_id_display }}</title>
  <style>
      body { font-family: 'Helvetica', Arial, sans-serif; margin: 20px; font-size: 10pt;}
      h1 { text-align: center; color: #333; font-size: 18pt; }
      .header { border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 20px; text-align:center; }
      .header p { margin: 2px 0; font-size: 10pt;}
      .info-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size:10pt; }
      .info-table th, .info-table td { border: 1px solid #ddd; padding: 6px; text-align: left; }
      .info-table th { background-color: #f2f2f2; font-weight: bold; width: 30%; }


      .overall-approval-area { margin-top: 99mm; page-break-inside: avoid;} /* Adjusted margin */
      .approval-title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 5px; color: black; }
      .approval-separator { border-bottom: 1px solid black; margin-bottom: 15px; }


      .approval-container { display: flex; justify-content: space-between; page-break-inside: avoid; width: 109%; }
      .approval-box {
          height: 200px;
          width: {{ HTML_REQUESTER_BOX_WIDTH_PERCENT }}%; /* Use Jinja variable */
          padding: 8pt; /* Increased padding */
          border: 1px solid black;   
          border-radius: 3px; /* Subtle rounding */
          background-color: #f9f9f9;
          text-align: center;
          min-height: {{ HTML_BOX_MIN_HEIGHT_PT }}pt; /* Use Jinja variable */
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
          justify-content: space-around; /* Better distribution of space */
      }
      .requester-approval { margin-right: 2%; } /* Space between boxes */
      .approver-placeholder-box {
          width: {{ HTML_APPROVER_BOX_WIDTH_PERCENT }}%; /* Use Jinja variable */
          padding: 8pt;
          border: 1px solid black;
          border-radius: 3px;
          background-color: #f9f9f9;
          text-align: center;
          min-height: {{ HTML_BOX_MIN_HEIGHT_PT }}pt;
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
          justify-content: space-around;
          align-items:center;
      }
      .approval-box h3, .approver-placeholder-box h3 {
          margin-top: 5px; margin-bottom: 8px; color: black;
          text-align: center; font-size: 11pt; font-family: 'Helvetica-Bold';
      }
      .signature-area-html { /* For requester signature only */
          min-height: {{ HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT }}pt; /* Use min-height */
          max-width: {{ HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT }}pt;
          margin: 5px auto;
          display: flex;
          align-items: center;
          justify-content: center;
          overflow: hidden;
      }
      .signature-area-html img { max-width: 100%; max-height: 100%; object-fit: contain; }
      .approver-details p { margin: 3px 0; font-size: 9pt; color: black; font-family: 'Helvetica'; }
      .attachment-preview { max-width: 100%; max-height: 500px; display: block; margin: 15px auto; padding: 5px; border: 1px solid #eee;}
      .page-break { page-break-before: always; }
  </style>
</head>
<body>
   <div class="header">
       <h1>REQUEST FORM</h1>
       <p>Request ID: {{ request_id_display }}</p>
       <p>Date: {{ request_date }}</p>
       <p>Company: {{ data.get('companyName', '') }}</p>
   </div>


   <table class="info-table">
       <tr><th>Field</th><th>Details</th></tr>
       <tr><td><strong>Requester Name</strong></td><td>{{ data.get('name', '') }}</td></tr>
       <tr><td><strong>Email</strong></td><td>{{ data.get('email', '') }}</td></tr>
       <tr><td><strong>Account Title</strong></td><td>{{ data.get('accountTitle', '') }}</td></tr>
       <tr><td><strong>Account Number</strong></td><td>{{ data.get('accountNumber', '') }}</td></tr>
       <tr><td><strong>IBAN Number</strong></td><td>{{ data.get('ibanNumber', '') }}</td></tr>
       <tr><td><strong>Bank Name</strong></td><td>{{ data.get('bankName', '') }}</td></tr>
       <tr><td><strong>Payment Type</strong></td><td>{{ data.get('paymentType', '') }}</td></tr>
       <tr><td><strong>Description</strong></td><td>{{ data.get('description', '') }}</td></tr>
       <tr><td><strong>Quantity</strong></td><td>{{ data.get('quantity', '') }}</td></tr>
       <tr><td><strong>Amount</strong></td><td>{{ data.get('amount', '') }} {{ data.get('currency', '') }}</td></tr>
       <tr><td><strong>Supporting Document File</strong></td><td>{{ data.get('document', 'No file uploaded') }}</td></tr>
   </table>


   <div class="overall-approval-area">
       <div class="approval-title">Approval & Request Information</div>
       <div class="approval-separator"></div>
   </div>


   <div class="approval-container">
       <div class="approval-box requester-approval">
           <h3>Requested By:</h3>
           <div class="signature-area-html">
               {{ requester_signature_html | safe }}
           </div>
           <div class="approval-details">
               <p>Name: {{ requester_name_from_form }}</p>
               <p>Date: {{ request_date }}</p>
           </div>
       </div>
   </div>
   {{ attachment_html_content | safe }} {# This is for embedding non-PDF preview #}

</body></html>
   """
    approval_section_heading = "Approval Section"
    approver_display_name = "PENDING APPROVAL"
    if approval_type == "CEO":
        approval_section_heading = "Approved By (CEO):"
        approver_display_name = CEO_APPROVER_NAME
    elif approval_type == "Standard":
        approval_section_heading = "Approved By (Standard):"
        approver_display_name = STANDARD_APPROVER_NAME

    env = Environment(loader=FileSystemLoader('templates'), cache_size=0, auto_reload=True)
    template = env.from_string(html_template_str)
    html_output = template.render(
        request_id_display=request_id_display,
        request_date=request_date,
        data=data,
        requester_signature_html=requester_signature_html,
        requester_name_from_form=requester_name_from_form,
        approval_section_heading=approval_section_heading,
        approver_display_name=approver_display_name,
        attachment_html_content=attachment_html_content,
        HTML_REQUESTER_BOX_WIDTH_PERCENT=HTML_REQUESTER_BOX_WIDTH_PERCENT,
        HTML_APPROVER_BOX_WIDTH_PERCENT=HTML_APPROVER_BOX_WIDTH_PERCENT,
        HTML_BOX_MIN_HEIGHT_PT=HTML_BOX_MIN_HEIGHT_PT,
        HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT=HTML_REQUESTER_SIG_IMG_MAX_HEIGHT_PT,
        HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT=HTML_REQUESTER_SIG_IMG_MAX_WIDTH_PT
    )

    options = {
        'encoding': 'UTF-8', 'quiet': '', 'page-size': 'A4',
        'margin-top': '15mm', 'margin-right': '15mm',
        'margin-bottom': '15mm', 'margin-left': '15mm',
        'footer-center': 'Page [page] of [topage]', 'footer-font-size': '8'
    }
    generated_pdf_bytes = None;
    pdf_generation_error = None
    try:
        if PDFKIT_CONFIG:
            logging.info("Attempting PDF generation using pdfkit.")
            generated_pdf_bytes = pdfkit.from_string(html_output, False, configuration=PDFKIT_CONFIG, options=options)
            logging.info("PDF generated successfully using pdfkit.")
        elif HTML:
            logging.info("pdfkit config missing/failed. Attempting PDF generation with weasyprint.")
            generated_pdf_bytes = HTML(string=html_output).write_pdf()
            logging.info("PDF generated successfully using weasyprint.")
        else:
            pdf_generation_error = "Neither pdfkit nor weasyprint are configured/available."
            logging.error(pdf_generation_error)
            return None, pdf_generation_error


    except Exception as e:
        pdf_generation_error = f"PDF generation failed: {e}";
        logging.error(pdf_generation_error, exc_info=True)
        if not PDFKIT_CONFIG and HTML and generated_pdf_bytes is None:
            pdf_generation_error = f"Weasyprint (fallback) also failed: {e}"
        elif PDFKIT_CONFIG and HTML and generated_pdf_bytes is None:
            logging.info("pdfkit failed. Attempting PDF generation with weasyprint fallback.")
            try:
                generated_pdf_bytes = HTML(string=html_output).write_pdf()
                logging.info("PDF generated successfully using weasyprint fallback.")
                pdf_generation_error = None
            except Exception as weasy_e:
                pdf_generation_error = f"pdfkit failed, and Weasyprint fallback also failed: {weasy_e}"
                logging.error(pdf_generation_error, exc_info=True)
                generated_pdf_bytes = None

    if generated_pdf_bytes is None:
        final_error_msg = pdf_generation_error if pdf_generation_error else 'Unknown PDF generation error'
        logging.error(f"Final PDF generation resulted in None bytes. Error: {final_error_msg}")
        return None, final_error_msg

    # --- FIX ATTACHMENT MERGING LOGIC FOR REQUEST PDF ---
    # This block ensures the request PDF itself *includes* the attachment if it's a PDF.
    # It returns the merged PDF (or original if no PDF attachment)
    # The logic here is now correct to integrate the attachment into the request PDF generated at this step.
    if attachment_path and os.path.exists(attachment_path) and attachment_path.lower().endswith('.pdf'):
        try:
            logging.info("Attempting to merge generated Request PDF with PDF attachment.")
            merger = PdfMerger()
            merger.append(BytesIO(generated_pdf_bytes))  # The form HTML rendered to PDF
            merger.append(attachment_path)  # The actual PDF file from disk

            merged_pdf_io = BytesIO()
            merger.write(merged_pdf_io)
            merger.close()
            merged_pdf_io.seek(0)
            final_pdf_bytes_with_attachment = merged_pdf_io.read()
            logging.info("Generated Request PDF and attachment PDF merged successfully.")
            return final_pdf_bytes_with_attachment, None  # Return the merged PDF
        except Exception as e:
            merge_error = f"Error merging generated Request PDF with attachment {attachment_path}: {e}"
            logging.error(merge_error, exc_info=True)
            # If merging fails, return the original generated PDF but report the error
            return generated_pdf_bytes, f"Partial success: Form PDF generated, but PDF attachment merge failed: {merge_error}"
    # If no attachment or attachment is not a PDF, return the generated PDF as is
    return generated_pdf_bytes, None


@app.route('/')
def index():
    creds = get_credentials()
    if not creds:
        logging.info("User not authenticated, redirecting to authorize.")
        return redirect(url_for('authorize'))
    return render_template('request_form.html', company_logos=COMPANY_LOGOS)


@app.route('/authorize')
def authorize():
    flow = get_google_auth_flow()
    authorization_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true',
                                                      prompt='consent')
    session['state'] = state
    logging.info("Redirecting to Google authorization URL.")
    return redirect(authorization_url)


@app.route('/oauth2callback')
def oauth2callback():
    state = session.pop('state', None)
    if not state or not request.args.get('state') or state != request.args.get('state'):
        logging.error("OAuth2 callback: State mismatch or missing.")
        return 'Invalid state parameter or state missing in request.', 400

    flow = get_google_auth_flow()
    try:
        authorization_response = request.url
        if not authorization_response.startswith("https://"):
            authorization_response = authorization_response.replace("http://", "https://", 1)

        logging.info(
            f"Fetching token using authorization response: {authorization_response[:100]}...")
        flow.fetch_token(authorization_response=authorization_response)
        credentials = flow.credentials

        if not credentials or not credentials.token:
            logging.error("OAuth2 callback: Failed to obtain token from Google.")
            return 'Failed to obtain token from Google. The authorization response might have been invalid or token fetch failed.', 400

        creds_data = {
            'token': credentials.token, 'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri, 'client_id': credentials.client_id,
            'client_secret': credentials.client_secret, 'scopes': credentials.scopes
        }
        session['credentials'] = json.dumps(creds_data)
        logging.info("Successfully obtained and stored credentials in session.")

        try:
            sheets_service = build('sheets', 'v4', credentials=credentials)
            ensure_sheet_headers(sheets_service, GOOGLE_SHEETS_SPREADSHEET_ID)
            logging.info("Ensured sheet headers after successful authentication.")
        except Exception as e_sheet:
            logging.error(f"Error ensuring sheet headers after auth: {e_sheet}", exc_info=True)

        return redirect(url_for('index'))
    except Exception as e:
        logging.error(f"Error during OAuth2 callback processing: {e}", exc_info=True)
        session.pop('credentials', None)
        return 'An error occurred during authentication. Please try authorizing again. Details: ' + str(e), 500


@app.route('/logout')
def logout():
    session.pop('credentials', None)
    session.pop('ceo_authenticated', None)
    session.pop('standard_authenticated', None)
    session.pop('dashboard_authenticated', None)
    session.pop('current_username', None)
    logging.info("User logged out.")
    return redirect(url_for('index'))


@app.route('/submit', methods=['POST'])
def submit():
    creds = get_credentials()
    if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401

    temp_dir = None
    try:
        action = request.form.get('action')
        if action not in ['preview', 'standard_approval', 'ceo_approval']:
            return jsonify({'success': False, 'message': 'Invalid action specified in form submission.'}), 400

        form_data = {
            'name': request.form.get('name', '').strip(),
            'email': request.form.get('email', '').strip(),
            'companyName': request.form.get('companyName', '').strip(),
            'accountTitle': request.form.get('accountTitle', '').strip(),
            'accountNumber': request.form.get('accountNumber', '').strip(),
            'ibanNumber': request.form.get('ibanNumber', '').strip(),
            'bankName': request.form.get('bankName', '').strip(),
            'paymentType': request.form.get('paymentType', '').strip(),
            'description': request.form.get('description', '').strip(),
            'quantity': request.form.get('quantity', '').strip(),
            'amount': request.form.get('amount', '').strip(),
            'currency': request.form.get('currency', '').strip(),
        }

        if action == 'preview':
            form_data['requestId'] = 'PREVIEW-ID'
        else:
            sheets_service = build('sheets', 'v4', credentials=creds)
            request_id_val, id_error = get_next_request_id(sheets_service, GOOGLE_SHEETS_SPREADSHEET_ID)
            if request_id_val is None:
                logging.error(f"Failed to generate Request ID: {id_error}")
                return jsonify({'success': False, 'message': f'Failed to generate Request ID: {id_error}'}), 500
            form_data['requestId'] = request_id_val

        attachment_path = None;
        attachment_link = None
        if 'document' in request.files:
            file = request.files['document']
            if file and file.filename:
                if allowed_file(file.filename):
                    if temp_dir is None:
                        temp_dir = tempfile.mkdtemp(prefix='request_submit_')
                        logging.info(f"Created temporary directory: {temp_dir}")

                    filename = secure_filename(file.filename)
                    attachment_path = os.path.join(temp_dir, filename)
                    file.save(attachment_path)
                    logging.info(f"Saved uploaded file to temporary path: {attachment_path}")
                    form_data['document'] = filename

                    if action in ['standard_approval', 'ceo_approval'] and attachment_path:
                        try:
                            logging.info(f"Uploading attachment: {filename} from path {attachment_path} to Drive...")
                            attachment_link = upload_file_from_path(file_path=attachment_path,
                                                                    file_name=f"Attachment_{form_data['requestId']}_{filename}",
                                                                    mime_type=file.content_type)
                            if not attachment_link: logging.error("Failed to upload attachment to Google Drive.")
                        except Exception as e_upload:
                            logging.error(f"Exception during attachment upload: {e_upload}", exc_info=True)
                else:
                    return jsonify({'success': False,
                                    'message': f'Unsupported file type: {file.filename}. Allowed: {ALLOWED_EXTENSIONS}'}), 400

        approval_type_str = "Standard" if action == "standard_approval" else (
            "CEO" if action == "ceo_approval" else "Preview")
        logging.info(
            f"Generating PDF for action '{action}' with approval type '{approval_type_str}' for Request ID: {form_data['requestId']}...")

        pdf_content, pdf_gen_error = generate_pdf(form_data, approval_type_str, attachment_path)

        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Cleaned up temporary directory tree: {temp_dir}")
            except OSError as e_rm_tree:
                logging.warning(f"Could not remove temporary directory tree {temp_dir}: {e_rm_tree}.", exc_info=True)
            except Exception as e_final_clean:
                logging.error(f"Unexpected error during temp directory tree cleanup {temp_dir}: {e_final_clean}",
                              exc_info=True)

        if pdf_content is None:
            err_msg = f'Failed to generate PDF. {pdf_gen_error if pdf_gen_error else "Unknown PDF generation error."}'.strip()
            logging.error(err_msg)
            return jsonify({'success': False, 'message': err_msg}), 500

        if action == 'preview':
            logging.info("Serving PDF preview.");
            return Response(pdf_content, mimetype='application/pdf',
                            headers={"Content-Disposition": "inline; filename=preview_request.pdf"})


        elif action in ['standard_approval', 'ceo_approval']:
            pdf_filename = f"Request_{form_data['requestId']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            pdf_link = None
            try:
                logging.info(f"Uploading generated PDF: {pdf_filename} (bytes) to Drive...")
                pdf_link = upload_file_from_bytes(file_content=pdf_content, file_name=pdf_filename,
                                                  mime_type='application/pdf')
                if not pdf_link:
                    return jsonify(
                        {'success': False, 'message': 'Failed to upload generated Request PDF to Google Drive.'}), 500
            except Exception as e_pdf_upload:
                logging.error(f"Exception during PDF upload for submission: {e_pdf_upload}", exc_info=True)
                return jsonify(
                    {'success': False, 'message': f'An error occurred during PDF upload: {str(e_pdf_upload)}'}), 500

            status = "Pending Standard Approval" if action == "standard_approval" else "Pending CEO Approval"

            logging.info(
                f"Adding data to Google Sheet with Request ID {form_data['requestId']} and status '{status}'...")
            sheet_added, sheet_response = add_to_sheet(form_data, pdf_link, attachment_link, status, approval_type_str)

            if sheet_added:
                msg = f'Request submitted for {approval_type_str} approval. Request ID: {form_data["requestId"]}'
                logging.info(msg)
                return jsonify({'success': True, 'message': msg, 'request_id': form_data['requestId']})
            else:
                err_msg_sheet = f'Failed to record request in Google Sheet: {sheet_response}'
                logging.error(err_msg_sheet)
                return jsonify({'success': False, 'message': err_msg_sheet}), 500


    except Exception as e:
        logging.error(f"An unexpected error occurred during submit: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'An internal server error occurred: {str(e)}'}), 500
    finally:
        pass


@app.route('/dashboard_login', methods=['GET', 'POST'])
def dashboard_login():
    if check_user_auth('dashboard'):
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if authenticate_user(username, password, 'dashboard'):
            logging.info(f"Dashboard login successful for user: {username}")
            return redirect(url_for('dashboard'))
        else:
            logging.warning(f"Dashboard login failed for user: {username}")
            return render_template('portal_login.html', portal_name="Dashboard", error="Invalid username or password")
    return render_template('portal_login.html', portal_name="Dashboard")


@app.route('/standard_login', methods=['GET', 'POST'])
def standard_login():
    if check_user_auth('standard'):
        return redirect(url_for('standard_approval'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if authenticate_user(username, password, 'standard'):
            logging.info(f"Standard login successful for user: {username}")
            return redirect(url_for('standard_approval'))
        else:
            logging.warning(f"Standard login failed for user: {username}")
            return render_template('portal_login.html', portal_name="Standard Approval",
                                   error="Invalid username or password")
    return render_template('portal_login.html', portal_name="Standard Approval")


@app.route('/ceo_login', methods=['GET', 'POST'])
def ceo_login():
    creds = get_credentials()
    if not creds: return redirect(url_for('authorize'))

    if session.get('ceo_authenticated'):
        return redirect(url_for('ceo_approval'))

    if request.method == 'POST':
        username = request.form.get('username');
        password = request.form.get('password')
        if authenticate_user(username, password, 'ceo'):
            logging.info(f"CEO login successful for user: {username}")
            return redirect(url_for('ceo_approval'))
        else:
            logging.warning(f"CEO login failed for user: {username}")
            return render_template('ceo_login.html', error="Invalid username or password")
    return render_template('ceo_login.html')


@app.route('/dashboard')
@require_auth('dashboard')
def dashboard():
    requests_list, error = get_requests_from_sheet(status_filter=None)

    if error:
        logging.error(f"Error fetching dashboard data: {error}")
        return render_template('error.html', message=f"Error fetching dashboard data: {error}")
    if requests_list is None:
        return render_template('error.html',
                               message="Could not retrieve requests. Please try logging in again or check sheet access.")

    try:
        requests_list.sort(
            key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get(
                'Timestamp') else datetime.min,
            reverse=True
        )
    except ValueError as ve:
        logging.warning(f"Could not sort requests by timestamp due to ValueError (likely malformed date): {ve}",
                        exc_info=True)
    except Exception as e_sort:
        logging.warning(f"An unexpected error occurred while sorting requests: {e_sort}", exc_info=True)

    return render_template('dashboard.html', requests=requests_list)


@app.route('/standard_approval')
@require_auth('standard')
def standard_approval():
    requests_list, error = get_requests_from_sheet(status_filter="Pending Standard Approval")
    if error: return render_template('error.html', message=f"Error fetching standard approval data: {error}")
    if requests_list is None: return render_template('error.html',
                                                     message="Could not retrieve standard approval requests.")
    try:
        requests_list.sort(key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get(
            'Timestamp') else datetime.min, reverse=True)
    except Exception:
        pass
    return render_template('standard_approval.html', requests=requests_list)


@app.route('/ceo_approval')
@require_auth('ceo')
def ceo_approval():
    requests_list, error = get_requests_from_sheet(status_filter="Pending CEO Approval")
    if error: return render_template('error.html', message=f"Error fetching CEO approval data: {error}")
    if requests_list is None: return render_template('error.html', message="Could not retrieve CEO approval requests.")
    try:
        requests_list.sort(key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get(
            'Timestamp') else datetime.min, reverse=True)
    except Exception:
        pass
    return render_template('ceo_approval.html', requests=requests_list)


def download_drive_file_bytes(file_link_or_id, creds):
    """Downloads a file from Google Drive and returns its byte content."""
    drive_service = build('drive', 'v3', credentials=creds)
    file_id = None
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', file_link_or_id)
    if match:
        file_id = match.group(1)
    elif re.match(r'^[a-zA-Z0-9_-]{25,}$', file_link_or_id):
        file_id = file_link_or_id
    elif 'drive.google.com' in file_link_or_id and 'id=' in file_link_or_id:
        try:
            file_id = file_link_or_id.split('id=')[-1].split('&')[0]
        except Exception:
            pass

    if not file_id:
        raise ValueError(f"Could not extract a valid Google Drive file ID from: {file_link_or_id}")

    logging.info(f"Attempting to download Drive file ID for merge/stamp: {file_id}")
    try:
        request_dl = drive_service.files().get_media(fileId=file_id)
        download_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(download_stream, request_dl)
        done = False
        while not done:
            status_dl, done = downloader.next_chunk()
            if status_dl: logging.debug(
                f"Download progress for {file_id}: {int(status_dl.progress() * 100)}% ({status_dl.resumable_progress} bytes).")
        download_stream.seek(0)
        file_bytes = download_stream.read()
        if not file_bytes:
            raise Exception(f"Downloaded empty file for ID {file_id}")
        return file_bytes
    except HttpError as e:
        logging.error(f"Google API HttpError downloading file {file_id}: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        raise Exception(f"Failed to download Drive file {file_id}: {e._get_reason()}") from e
    except Exception as e_dl:
        logging.error(f"Generic error downloading file {file_id}: {e_dl}", exc_info=True)
        raise Exception(f"Failed to download Drive file {file_id}: {str(e_dl)}") from e_dl


@app.route('/approve/<request_id>', methods=['POST'])
def approve_request(request_id):
    creds = get_credentials()
    if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401

    req, error = get_request_by_id(request_id)
    if error or req is None:
        return jsonify({'success': False, 'message': error or 'Request data not found'}), 404

    approval_type = req.get("Approval Type", "Unknown")
    current_status = req.get("Status")
    original_pdf_link = req.get("Request PDF Link")

    if approval_type == "CEO" and not session.get('ceo_authenticated'):
        return jsonify({'success': False, 'message': 'CEO authentication required.'}), 403
    elif approval_type == "Standard" and not session.get('standard_authenticated'):
        return jsonify({'success': False, 'message': 'Standard authentication required.'}), 403

    if current_status not in ["Pending Standard Approval", "Pending CEO Approval"]:
        return jsonify({'success': False, 'message': f'Request not pending approval (Status: {current_status})'}), 400

    if not original_pdf_link or "drive.google.com" not in original_pdf_link:
        return jsonify({'success': False, 'message': 'Original PDF link invalid or missing.'}), 400

    stamped_pdf_bytes = None
    overall_stamping_error = None
    approval_date_for_stamping = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    final_pdf_link = original_pdf_link
    original_pdf_file_id = None

    if not PDF_STAMPING_AVAILABLE:
        overall_stamping_error = "PDF stamping libraries unavailable."
    else:
        signature_image_url_to_stamp, approver_name_for_stamping, approval_section_heading_for_stamp = None, "", ""
        if approval_type == "CEO":
            signature_image_url_to_stamp = CEO_SIGNATURE_IMAGE_URL
            approver_name_for_stamping = CEO_APPROVER_NAME
            approval_section_heading_for_stamp = "Approved By (CEO):"
        elif approval_type == "Standard":
            signature_image_url_to_stamp = STANDARD_SIGNATURE_IMAGE_URL
            approver_name_for_stamping = STANDARD_APPROVER_NAME
            approval_section_heading_for_stamp = "Approved By (Standard):"
        else:
            overall_stamping_error = f"Unknown approval type '{approval_type}' for stamping."

        if not signature_image_url_to_stamp and not overall_stamping_error:
            overall_stamping_error = f"No signature image URL for approval type: {approval_type}"

        if not overall_stamping_error:
            signature_base64, signature_mime_type, sig_err = get_signature_data_from_url(signature_image_url_to_stamp)
            if sig_err:
                overall_stamping_error = f"Signature fetch error: {sig_err}"
            else:
                try:
                    original_pdf_bytes = download_drive_file_bytes(original_pdf_link, creds)

                    match_id = re.search(r'/d/([a-zA-Z0-9_-]+)', original_pdf_link)
                    if match_id:
                        original_pdf_file_id = match_id.group(1)
                    elif 'id=' in original_pdf_link:
                        original_pdf_file_id = original_pdf_link.split('id=')[-1].split('&')[0]
                    else:
                        original_pdf_file_id = original_pdf_link

                    if not original_pdf_file_id:
                        raise ValueError("Could not determine original PDF file ID for update.")

                    logging.info(
                        f"Stamping PDF for {request_id}. Approver: {approver_name_for_stamping}, Page: {STAMP_PAGE_INDEX}")
                    stamped_pdf_bytes, stamp_err = stamp_pdf_with_signature(
                        original_pdf_bytes, signature_base64, signature_mime_type,
                        approver_name_for_stamping, approval_date_for_stamping,
                        approval_section_heading_for_stamp, page=STAMP_PAGE_INDEX
                    )
                    if stamp_err: overall_stamping_error = f"Stamping failed: {stamp_err}"
                except Exception as e_dl_stamp:
                    overall_stamping_error = f"Download/Pre-stamp error: {str(e_dl_stamp)}"

    upload_stamped_error = None
    if stamped_pdf_bytes and not overall_stamping_error:
        try:
            if not original_pdf_file_id:
                raise Exception("Original PDF File ID not available for stamped PDF update.")
            stamped_file_name = f"Request_Signed_{request_id}_{datetime.now().strftime('%Y%m%d')}.pdf"
            logging.info(f"Uploading stamped PDF, replacing ID {original_pdf_file_id}...")
            uploaded_link = upload_file_from_bytes(
                file_content=stamped_pdf_bytes,
                file_name=stamped_file_name,
                mime_type='application/pdf',
                file_id_to_update=original_pdf_file_id
            )
            if uploaded_link:
                final_pdf_link = uploaded_link
                logging.info(f"Stamped PDF uploaded. New/Updated link: {final_pdf_link}")
            else:
                upload_stamped_error = "Failed to upload stamped PDF (upload_file_from_bytes returned None)."
        except Exception as e_upload_stamped:
            upload_stamped_error = f"Error uploading stamped PDF: {str(e_upload_stamped)}"
            logging.error(upload_stamped_error, exc_info=True)

    new_status = f"Approved by {approval_type}"
    updated, message = update_sheet_status(
        request_id, status=new_status,
        approval_date=approval_date_for_stamping,
        rejection_reason="",
        pdf_link=final_pdf_link
    )

    if updated:
        response_message = f"Request approved by {approval_type}."
        if stamped_pdf_bytes and not overall_stamping_error and not upload_stamped_error:
            response_message += " PDF stamped and updated."
        elif overall_stamping_error:
            response_message += f" Approval recorded, but PDF stamping failed: {overall_stamping_error}."
        elif upload_stamped_error:
            response_message += f" Approval recorded, PDF stamped, but upload of stamped PDF failed: {upload_stamped_error}."
        logging.info(f"Approval processed for {request_id}. Response: {response_message}")
        return jsonify({'success': True, 'message': response_message})
    else:
        logging.error(f"Failed to update sheet status for approved request {request_id}: {message}")
        return jsonify({'success': False,
                        'message': f'Failed to update request status in sheet after approval: {message}. Manual check required.'}), 500


@app.route('/reject/<request_id>', methods=['POST'])
def reject_request(request_id):
    creds = get_credentials()
    if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401

    req, error = get_request_by_id(request_id)
    if error or req is None:
        return jsonify({'success': False, 'message': error or 'Request data not found'}), 404

    rejection_reason = request.form.get('reason', 'No reason provided').strip()
    approval_type = req.get("Approval Type", "Unknown")
    current_status = req.get("Status")

    if approval_type == "CEO" and not session.get('ceo_authenticated'):
        return jsonify({'success': False, 'message': 'CEO authentication required to reject this request.'}), 403
    elif approval_type == "Standard" and not session.get('standard_authenticated'):
        return jsonify({'success': False, 'message': 'Standard authentication required to reject this request.'}), 403

    if current_status not in ["Pending Standard Approval", "Pending CEO Approval"]:
        return jsonify(
            {'success': False, 'message': f'Request is not currently pending approval (Status: {current_status})'}), 400

    new_status = f"Rejected by {approval_type}"
    rejection_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_rejection_reason = f"Rejected by {session.get('current_username', 'Unknown User')} on {rejection_timestamp}: {rejection_reason}"

    original_pdf_link = req.get("Request PDF Link")

    updated, message = update_sheet_status(
        request_id, status=new_status,
        approval_date="",
        rejection_reason=full_rejection_reason,
        pdf_link=original_pdf_link
    )

    if updated:
        logging.info(f"Request {request_id} rejected. Status updated successfully.")
        return jsonify({'success': True, 'message': 'Request rejected successfully'})
    else:
        logging.error(f"Failed to update sheet for rejection of request {request_id}: {message}")
        return jsonify({'success': False, 'message': f'Failed to update request status in sheet: {message}'}), 500


@app.route('/approve_voucher/<request_id>', methods=['POST'])
@require_auth('dashboard')
def approve_voucher(request_id):
    creds = get_credentials()
    if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401

    req, error = get_request_by_id(request_id)
    if error or req is None:
        return jsonify({'success': False, 'message': error or 'Request data not found'}), 404

    current_voucher_status = req.get("Voucher Approved By", "").strip()

    if current_voucher_status:
        return jsonify({'success': False, 'message': f'Voucher already finalized ({current_voucher_status}).'}), 400

    approver_name = session.get('current_username', 'Dashboard User')

    updated, message = update_sheet_status(
        request_id,
        voucher_approved_by=f"Approved by {approver_name} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        voucher_rejection_reason=""
    )

    if updated:
        logging.info(f"Voucher for Request ID {request_id} approved by {approver_name}.")
        return jsonify({'success': True, 'message': 'Voucher approved successfully!'})
    else:
        logging.error(f"Failed to update sheet for voucher approval for Request ID {request_id}: {message}")
        return jsonify({'success': False, 'message': f'Failed to approve voucher: {message}'}), 500


@app.route('/reject_voucher/<request_id>', methods=['POST'])
@require_auth('dashboard')
def reject_voucher(request_id):
    creds = get_credentials()
    if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401

    req, error = get_request_by_id(request_id)
    if error or req is None:
        return jsonify({'success': False, 'message': error or 'Request data not found'}), 404

    current_voucher_status = req.get("Voucher Approved By", "").strip()
    if current_voucher_status:
        return jsonify({'success': False, 'message': f'Voucher already finalized ({current_voucher_status}).'}), 400

    rejection_reason = request.form.get('reason', 'No reason provided').strip()
    rejector_name = session.get('current_username', 'Dashboard User')
    rejection_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_rejection_reason = f"Rejected by {rejector_name} on {rejection_timestamp}: {rejection_reason}"

    updated, message = update_sheet_status(
        request_id,
        voucher_approved_by=f"Rejected by {rejector_name}",
        voucher_rejection_reason=full_rejection_reason
    )

    if updated:
        logging.info(f"Voucher for Request ID {request_id} rejected by {rejector_name}.")
        return jsonify({'success': True, 'message': 'Voucher rejected successfully!'})
    else:
        logging.error(f"Failed to update sheet for voucher rejection for Request ID {request_id}: {message}")
        return jsonify({'success': False, 'message': f'Failed to reject voucher: {message}'}), 500


@app.route('/error')
def error_page():
    message = request.args.get('message', 'An unexpected error occurred.')
    return render_template('error.html', message=message)


@app.route('/edit_voucher_details/<request_id>')
def edit_voucher_details(request_id):
    creds = get_credentials()
    if not creds:
        return "Authentication required. Please login to the main application and try again.", 401

    sheets_service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID

    approver_data = get_approver_signatures_from_sheet(sheets_service, spreadsheet_id)

    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range="Sheet1!A:Z"
        ).execute()
    except HttpError as e:
        logging.error(f"API error fetching sheet data for editing voucher (ID: {request_id}): {e}")
        return f"Error fetching sheet data from Google: {e.resp.status} {e.resp.reason}. Please check sheet access and try again.", 500

    values = result.get('values', [])
    if not values or len(values) < 2:
        return f"Sheet '{spreadsheet_id}' is empty or missing headers.", 404

    headers = values[0]
    row_data_for_voucher = None
    for row in values[1:]:
        if row and len(row) > 0 and str(row[0]).strip() == str(request_id).strip():
            row_data_for_voucher = dict(zip(headers, row))
            break

    if not row_data_for_voucher:
        return f"Request ID '{request_id}' not found in the Google Sheet.", 404

    row_data_for_voucher.setdefault('Bank Name', row_data_for_voucher.get('Bank Name', ''))
    row_data_for_voucher.setdefault('IBAN', row_data_for_voucher.get('IBAN Number', ''))
    row_data_for_voucher.setdefault('Finance Review', approver_data.get('finance_review_name_default', ''))

    company_name_from_request = row_data_for_voucher.get('Company Name', 'Bpro')
    row_data_for_voucher['logo_url'] = COMPANY_LOGOS.get(company_name_from_request, COMPANY_LOGOS['Bpro'])

    try:
        amount = float(row_data_for_voucher.get('Amount', 0))
        quantity = float(row_data_for_voucher.get('Quantity', 1))
        if quantity == 0: quantity = 1
        rate = amount / quantity
        row_data_for_voucher.setdefault('Rate', str(round(rate, 2)))
    except (ValueError, TypeError):
        row_data_for_voucher.setdefault('Rate', row_data_for_voucher.get('Amount', '0'))

    currency_symbols = {"PKR": "Rs.", "USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "AUD": "A$", "CAD": "C$",
                        "CNY": "¥", "HKD": "HK$", "NZD": "NZ$", "SEK": "kr", "KRW": "₩", "SGD": "S$", "NOK": "kr",
                        "MXN": "Mex$", "INR": "₹", "BRL": "R$", "RUB": "₽", "ZAR": "R", "AED": "د.إ", "SAR": "ر.س",
                        "TRY": "₺", "THB": "฿", "MYR": "RM", "IDR": "Rp", "PHP": "₱", "VND": "₫", "PLN": "zł",
                        "CZK": "Kč", "HUF": "Ft", "RON": "lei", "ILS": "₪", "CLP": "CLP$", "COP": "COL$", "PEN": "S/.",
                        "ARS": "$"}
    row_data_for_voucher.setdefault('Currency_Symbol', currency_symbols.get(row_data_for_voucher['Currency'],
                                                                            row_data_for_voucher['Currency']))

    if 'Request PDF Link' not in row_data_for_voucher or not row_data_for_voucher['Request PDF Link']:
        return f"The 'Request PDF Link' (original form PDF) is missing for Request ID {request_id} in the sheet. This is required to generate a merged voucher.", 400

    prepared_by_names = approver_data.get('prepared_by_names', [])
    requester_name = row_data_for_voucher.get('Name', '')
    if requester_name and requester_name not in prepared_by_names:
        prepared_by_names.insert(0, requester_name)

    return render_template('voucher_edit_form.html',
                           request_data=row_data_for_voucher,
                           CEO_APPROVER_NAME=CEO_APPROVER_NAME,
                           STANDARD_APPROVER_NAME=STANDARD_APPROVER_NAME,
                           prepared_by_names=prepared_by_names)

@app.route('/generate_voucher', methods=['POST'])
def generate_voucher_route():
    creds = get_credentials()
    if not creds:
        return jsonify({"success": False, "message": "Authentication required"}), 401

    form_data = request.form
    request_id = form_data.get('request_id')
    if not request_id:
        return jsonify({"success": False, "message": "Missing request_id from form"}), 400

    original_request_pdf_link = form_data.get('original_request_pdf_link')
    if not original_request_pdf_link:
        return jsonify({"success": False,
                        "message": "Missing original_request_pdf_link from form. This is the link to the main request PDF."}), 400

    original_req_data, err = get_request_by_id(request_id)
    if err:
        logging.error(f"Failed to fetch original request data for voucher generation: {err}")
        return jsonify({"success": False, "message": f"Failed to fetch original request data: {err}"}), 500

    voucher_data_from_form = {key: form_data.get(key) for key in form_data}
    voucher_data_from_form['request_id'] = request_id
    company_name_for_voucher = original_req_data.get('Company Name', 'Bpro')
    voucher_data_from_form['voucher_logo_url'] = COMPANY_LOGOS.get(company_name_for_voucher, COMPANY_LOGOS['Bpro'])

    voucher_data_from_form['payment_from_bank'] = form_data.get('payment_from_bank', '')

    voucher_data_from_form['voucher_currency_for_total'] = form_data.get('voucher_currency_for_total',
                                                                         original_req_data.get('Currency_Symbol',
                                                                                               'Rs.'))
    voucher_data_from_form['currency_code'] = original_req_data.get('Currency', 'PKR')

    try:
        voucher_data_from_form['items_for_loop'] = []
        for i in range(1, 6):
            item_name = form_data.get(f'item_{i}_name', '').strip()
            if item_name:
                item_qty = float(form_data.get(f'item_{i}_quantity', 0))
                item_rate = float(form_data.get(f'item_{i}_rate', 0))
                item_amount = float(form_data.get(f'item_{i}_amount', 0))
                item_desc = form_data.get(f'item_{i}_description', '').strip()
                voucher_data_from_form['items_for_loop'].append({
                    'name': item_name,
                    'description': item_desc,
                    'quantity': item_qty,
                    'rate': item_rate,
                    'amount': item_amount
                })

        if not voucher_data_from_form['items_for_loop']:
            voucher_data_from_form['items_for_loop'].append({
                'name': '', 'description': original_req_data.get('Description', ''),
                'quantity': 0, 'rate': 0, 'amount': 0
            })

        voucher_data_from_form['voucher_total_amount'] = float(form_data.get('voucher_total_amount', 0))
    except ValueError:
        return jsonify(
            {"success": False, "message": "Invalid numeric value submitted for amount, quantity, or rate."}), 400

    if 'approval_date' not in voucher_data_from_form or not voucher_data_from_form['approval_date'].strip():
        voucher_data_from_form['approval_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # The prepared_by_name_selected comes from the form's dropdown
    voucher_data_from_form['prepared_by_name_selected'] = form_data.get('prepared_by_name_selected')


    # --- FIX START: Ensure all signature data is fetched BEFORE usage ---
    creds_for_sheet2 = get_credentials()
    sheets_service_for_sheet2 = build('sheets', 'v4', credentials=creds_for_sheet2)
    approver_data_for_voucher_sigs = get_approver_signatures_from_sheet(sheets_service_for_sheet2, GOOGLE_SHEETS_SPREADSHEET_ID)

    SIG_IMG_HTML_STYLE = 'max-width:100px; max-height:40px; object-fit:contain;'

    # Prepared By Signature: Look up based on selected name from form
    selected_prepared_by_name = voucher_data_from_form.get('prepared_by_name_selected', original_req_data.get('Name', 'N/A'))
    prepared_by_sig_url = approver_data_for_voucher_sigs.get("prepared_by_signature_urls_map", {}).get(
        selected_prepared_by_name, ""
    )
    prepared_by_sig_b64, prepared_by_mime, prepared_by_err = get_signature_data_from_url(prepared_by_sig_url)
    voucher_data_from_form[
        'prepared_by_signature_html'] = f'<img src="data:{prepared_by_mime};base64,{prepared_by_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if prepared_by_sig_b64 else ""

    # Finance Review Signature: Use the default from Sheet2
    finance_sig_url = approver_data_for_voucher_sigs.get('finance_review_signature_url', '')
    finance_sig_b64, finance_mime, finance_err = get_signature_data_from_url(finance_sig_url)
    voucher_data_from_form[
        'finance_signature_html'] = f'<img src="data:{finance_mime};base64,{finance_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if finance_sig_b64 else ""
    # Ensure finance_review_name is passed to template, prioritizing form data, then Sheet2 default
    voucher_data_from_form['finance_review_name'] = form_data.get('finance_review_name', approver_data_for_voucher_sigs.get('finance_review_name_default', 'N/A'))


    # Approved By Signature: Use the default from Sheet2
    approved_by_voucher_sig_url = approver_data_for_voucher_sigs.get('approved_by_signature_url', '')
    approved_by_sig_b64, approved_by_mime, approved_by_err = get_signature_data_from_url(approved_by_voucher_sig_url)
    voucher_data_from_form[
        'approved_by_signature_html'] = f'<img src="data:{approved_by_mime};base64,{approved_by_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if approved_by_sig_b64 else ""
    # Ensure approved_by_name is passed to template, prioritizing form data, then Sheet2 default
    voucher_data_from_form['approved_by_name'] = form_data.get('approved_by_name', approver_data_for_voucher_sigs.get('approved_by_name_default', 'N/A'))
    # --- FIX END ---


    env = Environment(loader=FileSystemLoader('templates'), cache_size=0, auto_reload=True)
    template = env.get_template('voucher_template.html')
    html_for_voucher_pdf = template.render(voucher_data_from_form=voucher_data_from_form)

    voucher_only_pdf_bytes = None
    merged_pdf_bytes = None
    final_merged_pdf_url = None
    temp_dir = tempfile.mkdtemp(prefix='voucher_gen_')

    try:
        pdf_options = {'encoding': 'UTF-8', 'quiet': '', 'page-size': 'A4', 'margin-top': '10mm',
                       'margin-bottom': '10mm', 'margin-left': '10mm', 'margin-right': '10mm'}
        if PDFKIT_CONFIG:
            logging.info("Attempting PDF generation using pdfkit.")
            generated_pdf_bytes = pdfkit.from_string(html_for_voucher_pdf, False, configuration=PDFKIT_CONFIG,
                                                     options=pdf_options)
            logging.info("PDF generated successfully using pdfkit.")
        elif HTML:
            logging.info("pdfkit config missing/failed. Attempting PDF generation with weasyprint.")
            generated_pdf_bytes = HTML(string=html_for_voucher_pdf).write_pdf()
            logging.info("PDF generated successfully using weasyprint.")
        else:
            pdf_generation_error = "Neither pdfkit nor weasyprint are configured/available."
            logging.error(pdf_generation_error)
            return jsonify(
                {"success": False, "message": pdf_generation_error}), 500

        if not generated_pdf_bytes:
            return jsonify({"success": False, "message": "Failed to generate voucher-only PDF content (empty)."}), 500

        # --- ATTACHMENT MERGING LOGIC FOR FINAL VOUCHER PDF ---
        merger = PdfMerger()
        merger.append(BytesIO(generated_pdf_bytes))  # 1st: Append the newly generated voucher PDF

        # 2nd: Append the original request PDF (which should already contain its attachment if PDF)
        original_request_pdf_link_from_sheet = original_req_data.get('Request PDF Link')
        if original_request_pdf_link_from_sheet and "drive.google.com" in original_request_pdf_link_from_sheet:
            try:
                logging.info(f"Downloading original request PDF from sheet: {original_request_pdf_link_from_sheet}")
                downloaded_request_pdf_bytes = download_drive_file_bytes(original_request_pdf_link_from_sheet, creds)
                merger.append(BytesIO(downloaded_request_pdf_bytes))
            except Exception as e:
                logging.warning(
                    f"Could not append original request PDF {original_request_pdf_link_from_sheet} to voucher: {e}",
                    exc_info=True)
        else:
            logging.warning(f"Original request PDF link missing or invalid: {original_request_pdf_link_from_sheet}")

        merged_output_stream = BytesIO()
        merger.write(merged_output_stream)
        merger.close()
        merged_pdf_bytes = merged_output_stream.getvalue()
        # --- END ATTACHMENT MERGING LOGIC FOR FINAL VOUCHER PDF ---


        merged_pdf_filename = f'Voucher_Merged_{request_id}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        final_merged_pdf_url = upload_file_from_bytes(
            file_content=merged_pdf_bytes,
            file_name=merged_pdf_filename,
            mime_type='application/pdf'
        )
        if not final_merged_pdf_url:
            return jsonify({"success": False, "message": "Failed to upload merged voucher PDF to Drive"}), 500

        sheets_service = build('sheets', 'v4', credentials=creds)
        spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID

        id_column_data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                    range="Sheet1!A:A").execute()
        ids_column_values = id_column_data.get('values', [])
        row_idx_to_update = -1
        for i, row_val_list in enumerate(ids_column_values):
            if row_val_list and str(row_val_list[0]).strip() == str(request_id).strip():
                row_idx_to_update = i
                break

        if row_idx_to_update == -1:
            return jsonify({"success": False,
                            "message": f"Request ID {request_id} not found in sheet for final update after voucher generation."}), 404

        updated, message = update_sheet_status(
            request_id,
            voucher_link=final_merged_pdf_url,
            voucher_generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )

        if not updated:
            logging.error(f"Failed to update sheet with voucher link for Request ID {request_id}: {message}")
            return jsonify(
                {"success": False, "message": f"Failed to update sheet after voucher generation: {message}"}), 500

        logging.info(
            f"Successfully generated, merged, and uploaded voucher for {request_id}. Final merged PDF link: {final_merged_pdf_url}")
        return jsonify({"success": True, "voucher_url": final_merged_pdf_url, "request_id": request_id})


    except HttpError as e:
        logging.error(f"Google API HttpError in generate_voucher route: {e.resp.status} - {e._get_reason()}",
                      exc_info=True)
        return jsonify({"success": False, "message": f"Google API Error: {e._get_reason()}"}), 500
    except ValueError as ve:
        logging.error(f"ValueError in generate_voucher route: {ve}", exc_info=True)
        return jsonify({"success": False, "message": str(ve)}), 400
    except Exception as e:
        import traceback
        logging.error(f"Unexpected error in generate_voucher route: {traceback.format_exc()}")
        return jsonify({"success": False, "message": f"An unexpected server error occurred: {str(e)}"}), 500
    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Cleaned up temp directory: {temp_dir}")
            except Exception as e_clean:
                logging.warning(f"Could not remove temp dir {temp_dir}: {e_clean}")


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        try:
            os.makedirs(UPLOAD_FOLDER);
            logging.info(f"Created upload folder: {UPLOAD_FOLDER}")
        except Exception as e:
            logging.error(f"Failed to create UPLOAD_FOLDER {UPLOAD_FOLDER}: {e}", exc_info=True)

    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
    logging.warning("OAUTHLIB_INSECURE_TRANSPORT is enabled. DO NOT USE THIS IN A PRODUCTION ENVIRONMENT.")

    logging.info("Starting Flask development server...")
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=True)
