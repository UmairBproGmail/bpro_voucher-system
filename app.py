import os
import json
from reportlab.lib.units import inch
import shutil
import io
import base64
import re
import requests
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

# For email sending
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

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
app.secret_key = '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' # IMPORTANT: Change this to a strong, random key in production!

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

REQUESTER_USERNAME = "requester_user" # NEW
REQUESTER_PASSWORD = "requester_password" # NEW


PORTAL_USERS = {
   DASHBOARD_USERNAME: DASHBOARD_PASSWORD,
   STANDARD_USERNAME: STANDARD_PASSWORD,
   CEO_USERNAME: CEO_PASSWORD,
   REQUESTER_USERNAME: REQUESTER_PASSWORD # NEW
}
# --- END NEW ---

# Email Configuration (NEW)
# IMPORTANT: For Gmail, use an App Password if you have 2-Factor Authentication enabled.
# Generate it in your Google Account security settings.
EMAIL_SENDER_ADDRESS = "umair.shahid@bpro.ai" # REPLACE WITH YOUR SENDER EMAIL
EMAIL_SENDER_PASSWORD = "sbpkblruzfoofxdo" # REPLACE WITH YOUR APP PASSWORD (for Gmail) or SMTP password
EMAIL_RECEIVER_ADDRESS = "umair.recordme@gmail.com"
EMAIL_SMTP_SERVER = "smtp.gmail.com" # Use "smtp.office365.com" for Outlook/Office 365
EMAIL_SMTP_PORT = 587


# --- START OF MODIFIED SECTION FOR GITHUB LOGOS ---

# GitHub Configuration for Logos (USER ACTION REQUIRED)
# Replace with your actual GitHub username, repository name, and branch name.
# The repository MUST be public for this to work without authentication.
GITHUB_USERNAME = "UmairBproGmail"
GITHUB_REPOSITORY_NAME = "bpro_voucher-system"
GITHUB_BRANCH_NAME = "main"

# Correct base URL for raw GitHub content
GITHUB_RAW_CONTENT_BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPOSITORY_NAME}/{GITHUB_BRANCH_NAME}/static/images/"

# Company Logos - now sourced from GitHub (these are URLs)
COMPANY_LOGOS = {
    "ML-1": GITHUB_RAW_CONTENT_BASE_URL + "machine-learning-1-logo.jpg",
    "Mpro": GITHUB_RAW_CONTENT_BASE_URL + "Market-Pro-Logo.png",
    "Enlatics": GITHUB_RAW_CONTENT_BASE_URL + "Enlatics-Logo.png",
    "DS": GITHUB_RAW_CONTENT_BASE_URL + "Developers-Studio-Logo.png",
    "CS": GITHUB_RAW_CONTENT_BASE_URL + "Cappersoft-Logo.png",
    "HRB": GITHUB_RAW_CONTENT_BASE_URL + "HRB-Logo.png",
    "Peace": GITHUB_RAW_CONTENT_BASE_URL + "peace-logo.jpg",
    "Zoompay": GITHUB_RAW_CONTENT_BASE_URL + "Zoom-Pay-Logo.png",
    "AML Watcher": GITHUB_RAW_CONTENT_BASE_URL + "AML-Watcher-Logo.png",
    "Bpro": GITHUB_RAW_CONTENT_BASE_URL + "bpro-ai-logo.jpg",
    "Facia": GITHUB_RAW_CONTENT_BASE_URL + "Facia-Logo.png",
    "the kyb": GITHUB_RAW_CONTENT_BASE_URL + "The-KYB-Logo.png",
    "Kyc/Aml": GITHUB_RAW_CONTENT_BASE_URL + "KYC-AML-Guide-Logo.png",
    "Techub": GITHUB_RAW_CONTENT_BASE_URL + "Techub-Logo.png",
}

# Base64 logo dictionary (this will be populated by the code below, containing data URLs)
COMPANY_LOGOS_BASE64 = {}

# Output folder for downloaded logos (optional, primarily for debugging or local caching)
LOGOS_DIR = "logos"
os.makedirs(LOGOS_DIR, exist_ok=True)


def download_and_convert_logos():
    """Download logos from GitHub and convert them to base64 data URLs"""
    logging.info("Starting download and base64 encoding of company logos from GitHub...")

    for name, url in COMPANY_LOGOS.items():
        try:
            # Critical check: Ensure user has replaced placeholders
            if "YOUR_GITHUB_USERNAME" in GITHUB_USERNAME or \
                    "YOUR_REPOSITORY_NAME" in GITHUB_REPOSITORY_NAME or \
                    "YOUR_GITHUB_USERNAME" in url or \
                    "YOUR_REPOSITORY_NAME" in url:
                logging.error(
                    f"GitHub configuration placeholder detected for logo '{name}'. URL: {url}. "
                    f"Please update GITHUB_USERNAME and GITHUB_REPOSITORY_NAME variables."
                )
                COMPANY_LOGOS_BASE64[name] = ""
                continue

            logging.info(f"Processing logo for: {name} from {url}")
            response = requests.get(url, timeout=20)
            response.raise_for_status()

            # Get MIME type from Content-Type header
            content_type = response.headers.get("Content-Type")
            if not content_type or not content_type.startswith("image/"):
                logging.warning(
                    f"Content-Type for '{name}' is '{content_type}', not recognized as image. "
                    f"Attempting to infer MIME type from URL extension."
                )
                # Try to guess from URL extension as fallback
                file_extension = os.path.splitext(url)[1].lower()
                mime_map = {
                    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                    '.gif': 'image/gif', '.svg': 'image/svg+xml', '.webp': 'image/webp'
                }
                if file_extension in mime_map:
                    content_type = mime_map[file_extension]
                    logging.info(f"Inferred MIME type '{content_type}' from extension '{file_extension}'.")
                else:
                    logging.error(
                        f"Could not determine valid image MIME type for '{name}'. "
                        f"Content-Type: '{content_type}', Extension: '{file_extension}'. Skipping."
                    )
                    COMPANY_LOGOS_BASE64[name] = ""
                    continue

            # Convert to base64
            image_data = response.content
            base64_data = base64.b64encode(image_data).decode('utf-8')
            data_url = f"data:{content_type};base64,{base64_data}"
            COMPANY_LOGOS_BASE64[name] = data_url

            # Optionally save to local cache
            filename = os.path.join(LOGOS_DIR, f"{name.replace('/', '-')}.{content_type.split('/')[-1]}")
            with open(filename, 'wb') as f:
                f.write(image_data)
            logging.info(f"Successfully processed logo for {name}")

        except requests.exceptions.RequestException as e:
            logging.error(f"Failed to download logo for {name} from {url}: {str(e)}")
            COMPANY_LOGOS_BASE64[name] = ""
        except Exception as e:
            logging.error(f"Unexpected error processing logo for {name}: {str(e)}")
            COMPANY_LOGOS_BASE64[name] = ""


# Execute the logo processing
download_and_convert_logos()

# Verify results
logging.info("Logo processing completed. Results:")
for name, data_url in COMPANY_LOGOS_BASE64.items():
    status = "SUCCESS" if data_url else "FAILED"
    logging.info(f"{name}: {status}")

# --- END OF MODIFIED SECTION FOR GITHUB LOGOS ---


UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'gif', 'doc', 'docx'}
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.metadata'
]


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
   # For Render, you might need to specify the path if wkhtmltopdf is installed via a buildpack
   # Or rely solely on weasyprint if that's easier to set up.
   WKHTMLTOPDF_PATH = '/usr/local/bin/wkhtmltopdf' # Common path for Linux/macOS
   # Example for Render's buildpack: WKHTMLTOPDF_PATH = '/app/.wkhtmltopdf/bin/wkhtmltopdf'
   # You should test this on your Render environment.
   logging.warning(f"Default WKHTMLTOPDF_PATH set to {WKHTMLTOPDF_PATH}. Verify this path for your Render deployment.")


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
   elif portal_type == 'requester' and username == REQUESTER_USERNAME and password == REQUESTER_PASSWORD: # NEW
       session['requester_authenticated'] = True
       session['current_username'] = REQUESTER_USERNAME
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
               elif portal_type == 'requester': # NEW
                   return redirect(url_for('requester_login'))
           return f(*args, **kwargs)

       return decorated_function

   return decorator




def allowed_file(filename):
   return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS




def get_google_auth_flow():
   return Flow.from_client_secrets_file('credentials.json', scopes=SCOPES, redirect_uri=url_for('oauth2callback', _external=True))








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
    drive_service.permissions().create(
        fileId=file_id, 
        body=permission, 
        fields='id',
        supportsAllDrives=True  # Add this for shared drives
    ).execute()
    logging.info(f"Set public permission for file ID: {file_id}")
except HttpError as error:
    if error.resp.status == 403:
        logging.warning(f"No permission to make file public: {error}")
    else:
        logging.warning(f"Could not set public permission for {file_id}: {error}")




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
    drive_service.permissions().create(
        fileId=file_id, 
        body=permission, 
        fields='id',
        supportsAllDrives=True  # Add this for shared drives
    ).execute()
    logging.info(f"Set public permission for file ID: {file_id}")
except HttpError as error:
    if error.resp.status == 403:
        logging.warning(f"No permission to make file public: {error}")
    else:
        logging.warning(f"Could not set public permission for {file_id}: {error}")



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
           "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason",
           "Voucher Prepared By" # NEW COLUMN
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
           "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason",
           "Voucher Prepared By" # NEW COLUMN
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
           "Voucher Rejection Reason": "",
           "Voucher Prepared By": ""  # This is already initialized as empty, which is what we need.
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
                       voucher_approved_by=None, voucher_rejection_reason=None,
                       voucher_prepared_by=None, voucher_link_status=None): # NEW PARAMETER: voucher_prepared_by, voucher_link_status
   creds = get_credentials()
   if not creds: return False, "Authentication failed"
   try:
       sheets_service = build('sheets', 'v4', credentials=creds)
       spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID


       header_result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                  range="Sheet1!1:1").execute()
       headers = header_result.get('values', [[]])[0]
       if not headers:
           ensure_sheet_headers(sheets_service, spreadsheet_id) # Attempt to fix headers if missing
           header_result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                      range="Sheet1!1:1").execute()
           headers = header_result.get('values', [[]])[0]
           if not headers: return False, "Sheet headers not found even after attempting fix."




       id_column_data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                                   range="Sheet1!A:A").execute()
       ids = id_column_data.get('values', [])
       row_index_to_update = -1
       if ids and len(ids) > 1: # Check if ids list is not empty and has more than just headers
           for i, row_val in enumerate(ids): # Start from 0 (which could be header row if not careful)
               if row_val and row_val[0] == request_id: # Check if row_val is not empty and first element matches
                   row_index_to_update = i # This is the 0-based index in the 'ids' list
                   break


       if row_index_to_update == -1:
           return False, "Request ID not found in sheet."


       sheet_row_num = row_index_to_update + 1 # Sheet rows are 1-based


       update_data = []


       def get_col_idx(header_name):
           try:
               return headers.index(header_name)
           except ValueError:
               logging.warning(f"Header '{header_name}' not found in sheet. Update will be skipped for this column.")
               return None


       def col_letter_from_0_idx(n_idx):
           string = ""
           n = n_idx + 1 # Convert 0-based index to 1-based for calculation
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
       if pdf_link is not None: # For updating Request PDF Link (e.g., after stamping)
           col_idx = get_col_idx("Request PDF Link")
           if col_idx is not None:
               update_data.append({
                   'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                   'values': [[pdf_link]]
               })


       # Voucher specific updates
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
       if voucher_prepared_by is not None: # NEW UPDATE BLOCK
           col_idx = get_col_idx("Voucher Prepared By")
           if col_idx is not None:
               update_data.append({
                   'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                   'values': [[voucher_prepared_by]]
               })
       if voucher_link_status is not None: # NEW UPDATE BLOCK (for "Voucher Sent for Payment" status)
           col_idx = get_col_idx("Voucher Approved By") # Reuse this column to display the new status text
           if col_idx is not None:
               update_data.append({
                   'range': f"Sheet1!{col_letter_from_0_idx(col_idx)}{sheet_row_num}",
                   'values': [[voucher_link_status]]
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


       if not values or len(values) < 2: # Need at least headers and one data row
           logging.warning("Sheet2 for signatures is empty or missing headers/data. Please populate it correctly.")
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
           if key not in col_indices: # Log if a primary category of column is missing
               logging.warning(f"Column category '{key}' not found in Sheet2 headers ({headers}). Signature data for this role might be incomplete.")




       for row_data in values[1:]:  # Skip header row
           if not row_data: # Skip empty rows
               continue


           # Helper to safely get data from row
           def get_cell_value(key_name):
               idx = col_indices.get(key_name)
               if idx is not None and idx < len(row_data):
                   return row_data[idx].strip()
               return ""


           # Collect all Prepared By names and their signature URLs
           prepared_name = get_cell_value('prepared_by_name')
           if prepared_name and prepared_name not in signatures_data['prepared_by_names']:
               signatures_data['prepared_by_names'].append(prepared_name)
               prepared_sig_url = get_cell_value('prepared_by_sig')
               if prepared_sig_url:
                   signatures_data['prepared_by_signature_urls_map'][prepared_name] = prepared_sig_url


           # Get Finance Reviewer data (use first valid entry)
           if signatures_data['finance_review_name_default'] == 'N/A': # Only if not already set
               finance_name_val = get_cell_value('finance_name')
               if finance_name_val:
                   signatures_data['finance_review_name_default'] = finance_name_val
                   finance_sig_url_val = get_cell_value('finance_sig')
                   if finance_sig_url_val:
                       signatures_data['finance_review_signature_url'] = finance_sig_url_val


           # Get Approved By data (use first valid entry)
           if signatures_data['approved_by_name_default'] == 'N/A': # Only if not already set
               approved_name_val = get_cell_value('approved_by_name')
               if approved_name_val:
                   signatures_data['approved_by_name_default'] = approved_name_val
                   approved_sig_url_val = get_cell_value('approved_by_sig')
                   if approved_sig_url_val:
                       signatures_data['approved_by_signature_url'] = approved_sig_url_val


       logging.info(f"Fetched signatures data from Sheet2: {signatures_data}")
       return signatures_data


   except HttpError as e:
       logging.error(f"Google API HttpError fetching signatures from Sheet2: {e.resp.status} - {e._get_reason()}",
                     exc_info=True)
       return signatures_data
   except Exception as e:
       logging.error(f"Error fetching signatures from Sheet2: {e}", exc_info=True)
       return signatures_data




# --- END NEW ---




def get_requests_from_sheet(status_filter=None):
   creds = get_credentials()
   if not creds: return None, "Authentication failed"
   try:
       sheets_service = build('sheets', 'v4', credentials=creds)
       spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID


       result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Sheet1!A:AA").execute() # Changed range to AA
       values = result.get('values', [])


       if not values: return [], None # No data at all


       headers = values[0]
       header_map = {h.strip(): i for i, h in enumerate(headers)}


       # Define expected headers carefully based on ensure_sheet_headers
       expected_headers_for_read = [
           "Request ID", "Timestamp", "Name", "Email", "Company Name",
           "Account Title", "Account Number", "IBAN Number", "Bank Name",
           "Payment Type", "Description", "Quantity", "Amount", "Currency",
           "Supporting Document Link", "Request PDF Link", "Status",
           "Approval Type", "Approval Date", "Rejection Reason",
           "Voucher PDF Link", "Voucher Generated At", "Voucher Approved By", "Voucher Rejection Reason",
           "Voucher Prepared By" # NEW COLUMN
       ]


       requests_list = []
       for i, row_data in enumerate(values[1:]): # Start from 1 to skip header row
           if not row_data or not any(cell.strip() for cell in row_data): # Skip completely empty rows
               continue


           request_data = {}
           # Populate request_data ensuring all expected keys exist, even if empty
           for header_name in expected_headers_for_read:
               index = header_map.get(header_name) # Get index from actual sheet headers
               if index is not None and index < len(row_data):
                   request_data[header_name] = row_data[index].strip()
               else:
                   request_data[header_name] = "" # Default to empty string if column missing or data missing


           if not request_data.get("Request ID"): # Critical check
               logging.warning(f"Skipping row {i + 2} due to missing Request ID: {row_data}")
               continue


           # Apply status filter if provided
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
   requests_list, error = get_requests_from_sheet(status_filter=None) # Get all requests
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
       elif HTML: # Fallback to weasyprint if HTML is available
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
       # Try weasyprint as a last resort if pdfkit failed and weasyprint is available
       if PDFKIT_CONFIG and HTML and generated_pdf_bytes is None: # Check if pdfkit was tried and failed, and HTML (weasyprint) is an option
           logging.info("pdfkit failed. Attempting PDF generation with weasyprint fallback.")
           try:
               generated_pdf_bytes = HTML(string=html_output).write_pdf()
               logging.info("PDF generated successfully using weasyprint fallback.")
               pdf_generation_error = None # Clear previous error if fallback succeeds
           except Exception as weasy_e:
               pdf_generation_error = f"pdfkit failed, and Weasyprint fallback also failed: {weasy_e}"
               logging.error(pdf_generation_error, exc_info=True)
               generated_pdf_bytes = None # Ensure it's None if both fail
       elif not PDFKIT_CONFIG and HTML and generated_pdf_bytes is None: # Weasyprint was primary and failed
            pdf_generation_error = f"Weasyprint (primary) failed: {e}"




   if generated_pdf_bytes is None:
       final_error_msg = pdf_generation_error if pdf_generation_error else 'Unknown PDF generation error'
       logging.error(f"Final PDF generation resulted in None bytes. Error: {final_error_msg}")
       return None, final_error_msg


   if attachment_path and os.path.exists(attachment_path) and attachment_path.lower().endswith('.pdf'):
       try:
           logging.info("Attempting to merge generated Request PDF with PDF attachment.")
           merger = PdfMerger()
           merger.append(BytesIO(generated_pdf_bytes))
           merger.append(attachment_path)


           merged_pdf_io = BytesIO()
           merger.write(merged_pdf_io)
           merger.close()
           merged_pdf_io.seek(0)
           final_pdf_bytes_with_attachment = merged_pdf_io.read()
           logging.info("Generated Request PDF and attachment PDF merged successfully.")
           return final_pdf_bytes_with_attachment, None
       except Exception as e:
           merge_error = f"Error merging generated Request PDF with attachment {attachment_path}: {e}"
           logging.error(merge_error, exc_info=True)
           return generated_pdf_bytes, f"Partial success: Form PDF generated, but PDF attachment merge failed: {merge_error}"
   return generated_pdf_bytes, None


# NEW: Send email function
def send_email_with_pdf(subject, body, to_email, pdf_bytes, pdf_filename, sender_email, sender_password, smtp_server, smtp_port):
    logging.info(f"Attempting to send email with subject: '{subject}' to {to_email}")
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'html')) # Changed to 'html' to allow basic formatting if needed

        if pdf_bytes and pdf_filename:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{pdf_filename}"')
            msg.attach(part)

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, sender_password)
            server.send_message(msg)
        logging.info(f"Email with PDF '{pdf_filename}' sent successfully to {to_email}")
        return True, None
    except Exception as e:
        logging.error(f"Failed to send email with PDF '{pdf_filename}' to {to_email}: {e}", exc_info=True)
        return False, str(e)


@app.route('/')
@require_auth('requester') # NEW: Protect request_form.html with requester login
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
       if not authorization_response.startswith("https://"): # Ensure HTTPS for Render
           authorization_response = authorization_response.replace("http://", "https://", 1)


       logging.info(
           f"Fetching token using authorization response: {authorization_response[:100]}...") # Log part of URL for debug
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
           # Don't fail the whole auth for this, but log it.


       return redirect(url_for('index'))
   except Exception as e:
       logging.error(f"Error during OAuth2 callback processing: {e}", exc_info=True)
       session.pop('credentials', None) # Clear potentially bad creds
       return 'An error occurred during authentication. Please try authorizing again. Details: ' + str(e), 500




@app.route('/logout')
def logout():
   session.pop('credentials', None)
   session.pop('ceo_authenticated', None)
   session.pop('standard_authenticated', None)
   session.pop('dashboard_authenticated', None)
   session.pop('requester_authenticated', None) # NEW
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
       else: # standard_approval or ceo_approval
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
           if file and file.filename: # Check if a file was actually selected
               if allowed_file(file.filename):
                   if temp_dir is None: # Create temp_dir only if needed
                       temp_dir = tempfile.mkdtemp(prefix='request_submit_')
                       logging.info(f"Created temporary directory: {temp_dir}")


                   filename = secure_filename(file.filename)
                   attachment_path = os.path.join(temp_dir, filename)
                   file.save(attachment_path)
                   logging.info(f"Saved uploaded file to temporary path: {attachment_path}")
                   form_data['document'] = filename # Store filename in form_data for PDF generation


                   # Upload attachment to Drive only if it's not a preview action
                   if action in ['standard_approval', 'ceo_approval'] and attachment_path:
                       try:
                           logging.info(f"Uploading attachment: {filename} from path {attachment_path} to Drive...")
                           attachment_link = upload_file_from_path(file_path=attachment_path,
                                                                   file_name=f"Attachment_{form_data['requestId']}_{filename}",
                                                                   mime_type=file.content_type) # Use file.content_type
                           if not attachment_link: logging.error("Failed to upload attachment to Google Drive.")
                       except Exception as e_upload:
                           logging.error(f"Exception during attachment upload: {e_upload}", exc_info=True)
                           # Optionally, decide if this is a critical failure
               else:
                   return jsonify({'success': False,
                                   'message': f'Unsupported file type: {file.filename}. Allowed: {ALLOWED_EXTENSIONS}'}), 400


       approval_type_str = "Standard" if action == "standard_approval" else ("CEO" if action == "ceo_approval" else "Preview")
       logging.info(f"Generating PDF for action '{action}' with approval type '{approval_type_str}' for Request ID: {form_data['requestId']}...")


       pdf_content, pdf_gen_error = generate_pdf(form_data, approval_type_str, attachment_path)


       # Cleanup temp_dir if it was created
       if temp_dir and os.path.exists(temp_dir):
           try:
               shutil.rmtree(temp_dir)
               logging.info(f"Cleaned up temporary directory tree: {temp_dir}")
           except OSError as e_rm_tree: # Catch specific OSError
               logging.warning(f"Could not remove temporary directory tree {temp_dir}: {e_rm_tree}.", exc_info=True)
           except Exception as e_final_clean: # Catch any other unexpected error during cleanup
               logging.error(f"Unexpected error during temp directory tree cleanup {temp_dir}: {e_final_clean}", exc_info=True)




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
                   return jsonify({'success': False, 'message': 'Failed to upload generated Request PDF to Google Drive.'}), 500
           except Exception as e_pdf_upload:
               logging.error(f"Exception during PDF upload for submission: {e_pdf_upload}", exc_info=True)
               return jsonify({'success': False, 'message': f'An error occurred during PDF upload: {str(e_pdf_upload)}'}), 500


           status = "Pending Standard Approval" if action == "standard_approval" else "Pending CEO Approval"


           logging.info(f"Adding data to Google Sheet with Request ID {form_data['requestId']} and status '{status}'...")
           sheet_added, sheet_response = add_to_sheet(form_data, pdf_link, attachment_link, status, approval_type_str)


           if sheet_added:
               msg = f'Request submitted for {approval_type_str} approval. Request ID: {form_data["requestId"]}'
               logging.info(msg)
               return jsonify({'success': True, 'message': msg, 'request_id': form_data['requestId']})
           else:
               err_msg_sheet = f'Failed to record request in Google Sheet: {sheet_response}'
               logging.error(err_msg_sheet)
               # Potentially, here you might want to delete the uploaded PDF and attachment from Drive if sheet fails
               return jsonify({'success': False, 'message': err_msg_sheet}), 500


   except Exception as e:
       logging.error(f"An unexpected error occurred during submit: {e}", exc_info=True)
       # Ensure temp_dir is cleaned up even if an error occurs mid-process
       if temp_dir and os.path.exists(temp_dir):
           try:
               shutil.rmtree(temp_dir)
               logging.info(f"Cleaned up temporary directory tree in error handler: {temp_dir}")
           except Exception as e_clean_err:
               logging.error(f"Error cleaning temp dir in exception handler: {e_clean_err}", exc_info=True)
       return jsonify({'success': False, 'message': f'An internal server error occurred: {str(e)}'}), 500
   # finally block removed as cleanup is handled within try and except


@app.route('/requester_login', methods=['GET', 'POST']) # NEW
def requester_login():
   if check_user_auth('requester'):
       return redirect(url_for('index'))
   if request.method == 'POST':
       username = request.form.get('username')
       password = request.form.get('password')
       if authenticate_user(username, password, 'requester'):
           session['requester_authenticated'] = True
           session['current_username'] = REQUESTER_USERNAME
           logging.info(f"Requester login successful for user: {username}")
           return redirect(url_for('index'))
       else:
           logging.warning(f"Requester login failed for user: {username}")
           return render_template('portal_login.html', portal_name="Requester Portal", error="Invalid username or password")
   return render_template('portal_login.html', portal_name="Requester Portal")


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
   if not creds: return redirect(url_for('authorize')) # Ensure Google auth first


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
   requests_list, error = get_requests_from_sheet(status_filter=None) # Get all requests


   if error:
       logging.error(f"Error fetching dashboard data: {error}")
       return render_template('error.html', message=f"Error fetching dashboard data: {error}")
   if requests_list is None: # Should ideally not happen if error is None, but good check
       return render_template('error.html',
                              message="Could not retrieve requests. Please try logging in again or check sheet access.")


   try:
       # Sort by Timestamp, newest first. Handle missing or malformed timestamps.
       requests_list.sort(
           key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get('Timestamp') else datetime.min,
           reverse=True
       )
   except ValueError as ve:
       logging.warning(f"Could not sort requests by timestamp due to ValueError (likely malformed date): {ve}", exc_info=True)
       # Continue with unsorted or partially sorted list
   except Exception as e_sort:
       logging.warning(f"An unexpected error occurred while sorting requests: {e_sort}", exc_info=True)
       # Continue with unsorted or partially sorted list


   return render_template('dashboard.html', requests=requests_list)




@app.route('/standard_approval')
@require_auth('standard')
def standard_approval():
   requests_list, error = get_requests_from_sheet(status_filter="Pending Standard Approval")
   if error: return render_template('error.html', message=f"Error fetching standard approval data: {error}")
   if requests_list is None: return render_template('error.html', message="Could not retrieve standard approval requests.")
   try:
       requests_list.sort(key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get('Timestamp') else datetime.min, reverse=True)
   except Exception: # Broad except for sorting, just proceed if fails
       pass
   return render_template('standard_approval.html', requests=requests_list)




@app.route('/ceo_approval')
@require_auth('ceo')
def ceo_approval():
   requests_list, error = get_requests_from_sheet(status_filter="Pending CEO Approval")
   if error: return render_template('error.html', message=f"Error fetching CEO approval data: {error}")
   if requests_list is None: return render_template('error.html', message="Could not retrieve CEO approval requests.")
   try:
       requests_list.sort(key=lambda x: datetime.strptime(x.get('Timestamp'), "%Y-%m-%d %H:%M:%S") if x.get('Timestamp') else datetime.min, reverse=True)
   except Exception:
       pass
   return render_template('ceo_approval.html', requests=requests_list)




def download_drive_file_bytes(file_link_or_id, creds):
   drive_service = build('drive', 'v3', credentials=creds)
   file_id = None
   # Regex to find file ID from common Drive link formats
   match_d_view = re.search(r'/d/([a-zA-Z0-9_-]+)', file_link_or_id)
   match_open_id = re.search(r'id=([a-zA-Z0-9_-]+)', file_link_or_id)


   if match_d_view:
       file_id = match_d_view.group(1)
   elif match_open_id:
       file_id = match_open_id.group(1)
   elif re.match(r'^[a-zA-Z0-9_-]{25,}$', file_link_or_id): # Check if it's likely an ID itself
       file_id = file_link_or_id


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
           if status_dl: logging.debug(f"Download progress for {file_id}: {int(status_dl.progress() * 100)}%")
       download_stream.seek(0)
       file_bytes = download_stream.read()
       if not file_bytes:
           raise Exception(f"Downloaded empty file for ID {file_id}")
       return file_bytes
   except HttpError as e:
       logging.error(f"Google API HttpError downloading file {file_id}: {e.resp.status} - {e._get_reason()}", exc_info=True)
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
   original_pdf_link = req.get("Request PDF Link") # This is the link to the PDF stored on Drive


   # Check portal authentication
   if approval_type == "CEO" and not session.get('ceo_authenticated'):
       return jsonify({'success': False, 'message': 'CEO authentication required.'}), 403
   elif approval_type == "Standard" and not session.get('standard_authenticated'):
       return jsonify({'success': False, 'message': 'Standard authentication required.'}), 403


   if current_status not in ["Pending Standard Approval", "Pending CEO Approval"]:
       return jsonify({'success': False, 'message': f'Request not pending approval (Status: {current_status})'}), 400


   if not original_pdf_link or ("drive.google.com" not in original_pdf_link and not re.match(r'^[a-zA-Z0-9_-]{25,}$', original_pdf_link)):
       return jsonify({'success': False, 'message': 'Original PDF link/ID invalid or missing.'}), 400


   stamped_pdf_bytes = None
   overall_stamping_error = None
   approval_date_for_stamping = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
   final_pdf_link_for_sheet = original_pdf_link # Default to original if stamping fails or not applicable
   original_pdf_file_id = None # Will be extracted from original_pdf_link


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


       if not signature_image_url_to_stamp and not overall_stamping_error: # If URL is missing for a known type
           overall_stamping_error = f"No signature image URL for approval type: {approval_type}"


       if not overall_stamping_error: # Proceed if no errors so far
           signature_base64, signature_mime_type, sig_err = get_signature_data_from_url(signature_image_url_to_stamp)
           if sig_err:
               overall_stamping_error = f"Signature fetch error: {sig_err}"
           else:
               try:
                   # Download the original PDF from Drive
                   original_pdf_bytes = download_drive_file_bytes(original_pdf_link, creds)


                   # Extract File ID for potential update
                   match_id_d = re.search(r'/d/([a-zA-Z0-9_-]+)', original_pdf_link)
                   match_id_id = re.search(r'id=([a-zA-Z0-9_-]+)', original_pdf_link)
                   if match_id_d: original_pdf_file_id = match_id_d.group(1)
                   elif match_id_id: original_pdf_file_id = match_id_id.group(1)
                   elif re.match(r'^[a-zA-Z0-9_-]{25,}$', original_pdf_link): original_pdf_file_id = original_pdf_link


                   if not original_pdf_file_id:
                       raise ValueError("Could not determine original PDF file ID for update.")


                   logging.info(f"Stamping PDF for {request_id}. Approver: {approver_name_for_stamping}, Page: {STAMP_PAGE_INDEX}")
                   stamped_pdf_bytes, stamp_err = stamp_pdf_with_signature(
                       original_pdf_bytes, signature_base64, signature_mime_type,
                       approver_name_for_stamping, approval_date_for_stamping,
                       approval_section_heading_for_stamp, page=STAMP_PAGE_INDEX
                   )
                   if stamp_err: overall_stamping_error = f"Stamping failed: {stamp_err}"
               except Exception as e_dl_stamp:
                   overall_stamping_error = f"Download/Pre-stamp error: {str(e_dl_stamp)}"


   # Upload the stamped PDF (if successful)
   upload_stamped_error = None
   if stamped_pdf_bytes and not overall_stamping_error:
       try:
           if not original_pdf_file_id: # Should have been set if stamping occurred
               raise Exception("Original PDF File ID not available for stamped PDF update.")


           # The name of the file in Drive will be updated if file_id_to_update is provided.
           # The original name might be like "Request_00001_timestamp.pdf"
           # Keeping the original name or updating it can be a choice.
           # For simplicity, we let upload_file_from_bytes handle the name if updating.
           # If we wanted a new name, we would provide file_name and not file_id_to_update.
           stamped_file_name_in_drive = req.get("Request PDF Link").split('/')[-1] # Try to get original name
           if '?' in stamped_file_name_in_drive: stamped_file_name_in_drive = stamped_file_name_in_drive.split('?')[0]




           logging.info(f"Uploading stamped PDF, replacing/updating ID {original_pdf_file_id}...")
           uploaded_link = upload_file_from_bytes(
               file_content=stamped_pdf_bytes,
               file_name=stamped_file_name_in_drive, # Pass a name, will be used if new, or ignored if updating
               mime_type='application/pdf',
               file_id_to_update=original_pdf_file_id # This ensures the original file is updated
           )
           if uploaded_link:
               final_pdf_link_for_sheet = uploaded_link # Use the new link from the update
               logging.info(f"Stamped PDF uploaded. New/Updated link: {final_pdf_link_for_sheet}")
           else:
               upload_stamped_error = "Failed to upload stamped PDF (upload_file_from_bytes returned None)."
       except Exception as e_upload_stamped:
           upload_stamped_error = f"Error uploading stamped PDF: {str(e_upload_stamped)}"
           logging.error(upload_stamped_error, exc_info=True)


   # Update sheet status
   new_status = f"Approved by {approval_type}"
   updated, message = update_sheet_status(
       request_id, status=new_status,
       approval_date=approval_date_for_stamping,
       rejection_reason="", # Clear any previous rejection reason
       pdf_link=final_pdf_link_for_sheet # Update with the potentially new link
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
       # This is a critical state: PDF might be stamped and uploaded, but sheet not updated.
       return jsonify({'success': False, 'message': f'Failed to update request status in sheet after approval: {message}. Manual check required.'}), 500




@app.route('/reject/<request_id>', methods=['POST'])
def reject_request(request_id):
   creds = get_credentials()
   if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401


   req, error = get_request_by_id(request_id)
   if error or req is None:
       return jsonify({'success': False, 'message': error or 'Request data not found'}), 404


   rejection_reason = request.form.get('reason', 'No reason provided').strip()
   approval_type = req.get("Approval Type", "Unknown") # From the original request submission
   current_status = req.get("Status")


   # Check portal authentication based on who is rejecting
   # This logic assumes the current portal user is the one whose queue it was in.
   # If Standard user rejects a "Pending Standard Approval" item, approval_type would be "Standard".
   # If CEO user rejects a "Pending CEO Approval" item, approval_type would be "CEO".
   acting_portal = None
   if session.get('ceo_authenticated'): acting_portal = "CEO"
   elif session.get('standard_authenticated'): acting_portal = "Standard"


   if not acting_portal:
       return jsonify({'success': False, 'message': 'User portal authentication unclear.'}), 403


   # Ensure the request is in a state that this portal user can reject
   if approval_type == "CEO" and not session.get('ceo_authenticated'):
       return jsonify({'success': False, 'message': 'CEO authentication required to reject this request.'}), 403
   elif approval_type == "Standard" and not session.get('standard_authenticated'):
       return jsonify({'success': False, 'message': 'Standard authentication required to reject this request.'}), 403


   if current_status not in ["Pending Standard Approval", "Pending CEO Approval"]:
       return jsonify({'success': False, 'message': f'Request is not currently pending approval (Status: {current_status})'}), 400


   # Determine rejector based on current portal session, not necessarily original req.approval_type
   rejector_role = "Unknown"
   if session.get('ceo_authenticated'):
       rejector_role = "CEO"
   elif session.get('standard_authenticated'):
       rejector_role = "Standard"
   # Add other roles if necessary


   new_status = f"Rejected by {rejector_role}" # Status reflects who performed the rejection
   rejection_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
   # Include who rejected it in the reason for clarity
   full_rejection_reason = f"Rejected by {session.get('current_username', rejector_role)} on {rejection_timestamp}: {rejection_reason}"


   original_pdf_link = req.get("Request PDF Link") # PDF link remains unchanged on rejection


   updated, message = update_sheet_status(
       request_id, status=new_status,
       approval_date="", # No approval date for rejection
       rejection_reason=full_rejection_reason,
       pdf_link=original_pdf_link # Keep the original PDF link
   )


   if updated:
       logging.info(f"Request {request_id} rejected. Status updated successfully.")
       return jsonify({'success': True, 'message': 'Request rejected successfully'})
   else:
       logging.error(f"Failed to update sheet for rejection of request {request_id}: {message}")
       return jsonify({'success': False, 'message': f'Failed to update request status in sheet: {message}'}), 500




@app.route('/approve_voucher/<request_id>', methods=['POST'])
@require_auth('dashboard') # Only dashboard users can approve/reject vouchers
def approve_voucher(request_id):
   creds = get_credentials()
   if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401


   req, error = get_request_by_id(request_id)
   if error or req is None:
       return jsonify({'success': False, 'message': error or 'Request data not found'}), 404


   current_voucher_status = req.get("Voucher Approved By", "").strip()
   # Allow approval only if not already approved or rejected or sent for payment
   if current_voucher_status and ("Approved by" in current_voucher_status or "Rejected by" in current_voucher_status or "Sent for Payment" in current_voucher_status):
       return jsonify({'success': False, 'message': f'Voucher already finalized ({current_voucher_status}).'}), 400


   approver_name = session.get('current_username', 'Dashboard User') # User from dashboard session
   voucher_pdf_link = req.get("Voucher PDF Link")

   if not voucher_pdf_link:
       return jsonify({'success': False, 'message': 'Voucher PDF link is missing, cannot send email.'}), 400

   # Download the voucher PDF
   voucher_pdf_bytes = None
   try:
       voucher_pdf_bytes = download_drive_file_bytes(voucher_pdf_link, creds)
   except Exception as e:
       logging.error(f"Failed to download voucher PDF for email: {e}", exc_info=True)
       return jsonify({'success': False, 'message': f'Failed to download voucher PDF for email: {e}'}), 500

   voucher_filename = f"Voucher_RequestID_{request_id}.pdf"
   email_subject = f"Payment Voucher for Request ID: {request_id} - Approved and Sent for Payment"
   email_body = f"""
   <p>Dear Payment Team,</p>
   <p>The payment voucher for Request ID <b>{request_id}</b> has been approved and is attached for your processing.</p>
   <p><b>Requester Name:</b> {req.get('Name', 'N/A')}</p>
   <p><b>Company:</b> {req.get('Company Name', 'N/A')}</p>
   <p><b>Description:</b> {req.get('Description', 'N/A')}</p>
   <p><b>Amount:</b> {req.get('Amount', 'N/A')} {req.get('Currency', 'N/A')}</p>
   <p>Regards,<br>Bpro Voucher System</p>
   """

   email_sent, email_error = send_email_with_pdf(
       subject=email_subject,
       body=email_body,
       to_email=EMAIL_RECEIVER_ADDRESS,
       pdf_bytes=voucher_pdf_bytes,
       pdf_filename=voucher_filename,
       sender_email=EMAIL_SENDER_ADDRESS,
       sender_password=EMAIL_SENDER_PASSWORD,
       smtp_server=EMAIL_SMTP_SERVER,
       smtp_port=EMAIL_SMTP_PORT
   )

   if not email_sent:
       logging.error(f"Email failed for voucher approval {request_id}: {email_error}")
       return jsonify({'success': False, 'message': f'Voucher approved in sheet, but email sending failed: {email_error}'}), 500


   updated, message = update_sheet_status(
       request_id,
       voucher_approved_by=f"Approved by {approver_name}", # Keep this for internal record of who approved
       voucher_rejection_reason="",
       voucher_link_status="Voucher Sent for Payment" # NEW: Update status for dashboard display
   )


   if updated:
       logging.info(f"Voucher for Request ID {request_id} approved by {approver_name} and email sent.")
       return jsonify({'success': True, 'message': 'Voucher approved and sent for payment successfully!', 'new_status': 'Voucher Sent for Payment'})
   else:
       logging.error(f"Failed to update sheet for voucher approval for Request ID {request_id}: {message}")
       return jsonify({'success': False, 'message': f'Failed to approve voucher: {message}'}), 500




@app.route('/reject_voucher/<request_id>', methods=['POST'])
@require_auth('dashboard') # Only dashboard users can approve/reject vouchers
def reject_voucher(request_id):
   creds = get_credentials()
   if not creds: return jsonify({'success': False, 'message': 'Authentication required'}), 401


   req, error = get_request_by_id(request_id)
   if error or req is None:
       return jsonify({'success': False, 'message': error or 'Request data not found'}), 404


   current_voucher_status = req.get("Voucher Approved By", "").strip()
   # Allow rejection only if not already approved or rejected or sent for payment
   if current_voucher_status and ("Approved by" in current_voucher_status or "Rejected by" in current_voucher_status or "Sent for Payment" in current_voucher_status):
       return jsonify({'success': False, 'message': f'Voucher already finalized ({current_voucher_status}).'}), 400


   rejection_reason = request.form.get('reason', 'No reason provided').strip()
   rejector_name = session.get('current_username', 'Dashboard User') # User from dashboard session
   rejection_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
   full_rejection_reason = f"Rejected by {rejector_name} on {rejection_timestamp}: {rejection_reason}"


   updated, message = update_sheet_status(
       request_id,
       voucher_approved_by=f"Rejected by {rejector_name}", # Mark as rejected by this user
       voucher_rejection_reason=full_rejection_reason
   )


   if updated:
       logging.info(f"Voucher for Request ID {request_id} rejected by {rejector_name}.")
       return jsonify({'success': True, 'message': 'Voucher rejected successfully!'})
   else:
       logging.error(f"Failed to update sheet for voucher rejection for Request ID {request_id}: {message}")
       return jsonify({'success': False, 'message': f'Failed to reject voucher: {message}'}), 500




@app.route('/edit_voucher_details/<request_id>')
@require_auth('dashboard') # Ensure only authenticated dashboard users can access this
def edit_voucher_details(request_id):
   creds = get_credentials()
   if not creds:
       return "Authentication required.", 401


   sheets_service = build('sheets', 'v4', credentials=creds)
   spreadsheet_id = GOOGLE_SHEETS_SPREADSHEET_ID


   # Fetch original request data from Sheet1
   original_req_data, err = get_request_by_id(request_id)
   if err or not original_req_data:
       logging.error(f"Error fetching original request data for voucher edit (ID: {request_id}): {err}")
       return f"Error fetching request data for ID '{request_id}': {err or 'Not found'}. Please ensure the request exists and is approved.", 404


   # Fetch approver signature data from Sheet2
   approver_data_from_sheet2 = get_approver_signatures_from_sheet(sheets_service, spreadsheet_id)


   # Prepare data for the voucher_edit_form.html template
   # Start with original request data and then override/add voucher specific fields
   voucher_form_data = original_req_data.copy() # Make a copy to avoid modifying original dict


   # Default/Derived values for the form
   voucher_form_data.setdefault('Bank Name', voucher_form_data.get('Bank Name', ''))
   voucher_form_data.setdefault('IBAN', voucher_form_data.get('IBAN Number', '')) # Use IBAN Number for IBAN field
   voucher_form_data.setdefault('Finance Review', approver_data_from_sheet2.get('finance_review_name_default', 'N/A')) # Default Finance Reviewer Name


   company_name_from_request = voucher_form_data.get('Company Name', 'Bpro') # Default to 'Bpro' if not found

   # Fetch base64 logo directly for embedding in the voucher HTML
   voucher_form_data['logo_data_url'] = COMPANY_LOGOS_BASE64.get(company_name_from_request, COMPANY_LOGOS_BASE64.get('Bpro', '')) # Fallback to Bpro logo then empty

   try:
       amount = float(voucher_form_data.get('Amount', 0))
       quantity = float(voucher_form_data.get('Quantity', 1))
       if quantity == 0: quantity = 1 # Avoid division by zero
       rate = amount / quantity
       voucher_form_data.setdefault('Rate', str(round(rate, 2)))
   except (ValueError, TypeError):
       voucher_form_data.setdefault('Rate', voucher_form_data.get('Amount', '0')) # Fallback for rate


   currency_symbols = {"PKR": "Rs.", "USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "AUD": "A$", "CAD": "C$",
                       "CNY": "¥", "HKD": "HK$", "NZD": "NZ$", "SEK": "kr", "KRW": "₩", "SGD": "S$", "NOK": "kr",
                       "MXN": "Mex$", "INR": "₹", "BRL": "R$", "RUB": "₽", "ZAR": "R", "AED": "د.إ", "SAR": "ر.س",
                       "TRY": "₺", "THB": "฿", "MYR": "RM", "IDR": "Rp", "PHP": "₱", "VND": "₫", "PLN": "zł",
                       "CZK": "Kč", "HUF": "Ft", "RON": "lei", "ILS": "₪", "CLP": "CLP$", "COP": "COL$", "PEN": "S/.",
                       "ARS": "$"}
   voucher_form_data.setdefault('Currency_Symbol', currency_symbols.get(voucher_form_data.get('Currency',''), voucher_form_data.get('Currency','')))


   # Get the 'Voucher Prepared By' name from the sheet if it exists, otherwise default to original requester name
   voucher_prepared_by_from_sheet = original_req_data['Voucher Prepared By'] if original_req_data.get('Voucher Prepared By') else original_req_data.get('Name', '')
   voucher_form_data.setdefault('Voucher Prepared By', voucher_prepared_by_from_sheet) # NEW

   # Prepare list for "Prepared By" dropdown, ensure original requester and sheet's prepared by are options
   prepared_by_names_list = list(approver_data_from_sheet2.get('prepared_by_names', [])) # Make a mutable copy
   original_requester_name = voucher_form_data.get('Name', '')
   if original_requester_name and original_requester_name not in prepared_by_names_list:
       prepared_by_names_list.insert(0, original_requester_name) # Add to start if not present
   if voucher_prepared_by_from_sheet and voucher_prepared_by_from_sheet not in prepared_by_names_list: # Ensure sheet's prepared by is also an option
       prepared_by_names_list.insert(0, voucher_prepared_by_from_sheet)


   if 'Request PDF Link' not in voucher_form_data or not voucher_form_data['Request PDF Link']:
       return f"The 'Request PDF Link' is missing for Request ID {request_id}. This is required for voucher generation.", 400


   return render_template('voucher_edit_form.html',
                          request_data=voucher_form_data, # This now contains merged data
                          CEO_APPROVER_NAME=CEO_APPROVER_NAME,
                          STANDARD_APPROVER_NAME=STANDARD_APPROVER_NAME,
                          prepared_by_names=prepared_by_names_list,
                          company_logo_data_url=voucher_form_data['logo_data_url']) # Pass this explicitly for hidden input


@app.route('/generate_voucher', methods=['POST'])
@require_auth('dashboard') # Ensure only authenticated dashboard users can access this
def generate_voucher_route():
   creds = get_credentials() # Already checked by require_auth
   if not creds:
       return jsonify({"success": False, "message": "Authentication required"}), 401


   form_data_from_html_form = request.form # Data submitted from voucher_edit_form.html
   request_id = form_data_from_html_form.get('request_id')
   if not request_id:
       return jsonify({"success": False, "message": "Missing request_id from form"}), 400


   # Fetch original request data to ensure consistency and for fallbacks
   original_req_data_from_sheet, err = get_request_by_id(request_id)
   if err or not original_req_data_from_sheet:
       logging.error(f"Failed to fetch original request data for voucher generation (ID: {request_id}): {err}")
       return jsonify({"success": False, "message": f"Failed to fetch original request data: {err or 'Not found'}"}), 500


   # Prepare the data payload for voucher_template.html
   # Prioritize form data, but use original sheet data for non-editable fields or fallbacks
   voucher_template_data = {}
   voucher_template_data['request_id'] = request_id
   voucher_template_data['voucher_payment_type'] = original_req_data_from_sheet.get('Payment Type', '') # From sheet
   voucher_template_data['payment_from_bank'] = form_data_from_html_form.get('payment_from_bank', '') # From form
   voucher_template_data['voucher_account_title'] = form_data_from_html_form.get('voucher_account_title', original_req_data_from_sheet.get('Account Title', ''))
   voucher_template_data['voucher_bank_name'] = form_data_from_html_form.get('voucher_bank_name', original_req_data_from_sheet.get('Bank Name', ''))
   voucher_template_data['voucher_iban'] = form_data_from_html_form.get('voucher_iban', original_req_data_from_sheet.get('IBAN Number', ''))


   # Company Logo: Use the logo data URL from the hidden input from voucher_edit_form
   voucher_template_data['voucher_logo_data_url'] = form_data_from_html_form.get('voucher_logo_data_url', '')


   # Approval Date: Use the one from the form (hidden input, sourced from sheet initially)
   voucher_template_data['approval_date'] = form_data_from_html_form.get('approval_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))




   # Item Details
   try:
       voucher_template_data['items_for_loop'] = []
       for i in range(1, 6): # Max 5 items as per voucher_edit_form.html
           item_name = form_data_from_html_form.get(f'item_{i}_name', '').strip()
           if item_name: # Only add if item name is present
               item_desc = form_data_from_html_form.get(f'item_{i}_description', item_name).strip() # Default desc to name
               item_qty_str = form_data_from_html_form.get(f'item_{i}_quantity', '0')
               item_rate_str = form_data_from_html_form.get(f'item_{i}_rate', '0')
               item_amount_str = form_data_from_html_form.get(f'item_{i}_amount', '0')


               voucher_template_data['items_for_loop'].append({
                   'name': item_name,
                   'description': item_desc,
                   'quantity': float(item_qty_str) if item_qty_str else 0,
                   'rate': float(item_rate_str) if item_rate_str else 0,
                   'amount': float(item_amount_str) if item_amount_str else 0
               })
       # If no items were parsed from form but original request had description/amount, create one item line
       if not voucher_template_data['items_for_loop'] and original_req_data_from_sheet.get('Description'):
           voucher_template_data['items_for_loop'].append({
               'name': original_req_data_from_sheet.get('Description', 'N/A'), # Use original description as item name
               'description': original_req_data_from_sheet.get('Description', 'N/A'),
               'quantity': float(original_req_data_from_sheet.get('Quantity', 1)),
               'rate': float(original_req_data_from_sheet.get('Amount', 0)) / (float(original_req_data_from_sheet.get('Quantity', 1)) or 1),
               'amount': float(original_req_data_from_sheet.get('Amount', 0))
           })
       elif not voucher_template_data['items_for_loop']: # Fallback if absolutely no items
            voucher_template_data['items_for_loop'].append({'name': 'N/A', 'description': 'N/A', 'quantity': 0, 'rate': 0, 'amount': 0})




       voucher_total_amount_str = form_data_from_html_form.get('voucher_total_amount', '0')
       voucher_template_data['voucher_total_amount'] = float(voucher_total_amount_str) if voucher_total_amount_str else 0
   except ValueError as ve:
       logging.error(f"ValueError processing item amounts for voucher: {ve}")
       return jsonify({"success": False, "message": "Invalid numeric value for item quantity, rate, or amount."}), 400


   # Currency for total display (symbol)
   voucher_template_data['voucher_currency_for_total'] = form_data_from_html_form.get('voucher_currency_for_total', 'Rs.') # From form


   # Signatories
   sheets_service = build('sheets', 'v4', credentials=creds) # For Sheet2
   approver_data_sheet2 = get_approver_signatures_from_sheet(sheets_service, GOOGLE_SHEETS_SPREADSHEET_ID)
   SIG_IMG_HTML_STYLE = 'max-width:100px; max-height:40px; object-fit:contain;'


   # Prepared By
   selected_prepared_by_name = form_data_from_html_form.get('prepared_by_name_selected', original_req_data_from_sheet.get('Name', 'N/A'))
   voucher_template_data['prepared_by_name_selected'] = selected_prepared_by_name
   prepared_by_sig_url = approver_data_sheet2.get("prepared_by_signature_urls_map", {}).get(selected_prepared_by_name, "")
   prepared_by_sig_b64, prepared_by_mime, _ = get_signature_data_from_url(prepared_by_sig_url)
   voucher_template_data['prepared_by_signature_html'] = f'<img src="data:{prepared_by_mime};base64,{prepared_by_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if prepared_by_sig_b64 else ""


   # Finance Review
   finance_review_name_on_form = form_data_from_html_form.get('finance_review_name', '').strip()
   voucher_template_data['finance_review_name'] = finance_review_name_on_form if finance_review_name_on_form else approver_data_sheet2.get('finance_review_name_default', 'N/A')
   finance_sig_url_to_use = approver_data_sheet2.get("prepared_by_signature_urls_map", {}).get(voucher_template_data['finance_review_name'], approver_data_sheet2.get('finance_review_signature_url', '')) # Try mapping first, then default
   finance_sig_b64, finance_mime, _ = get_signature_data_from_url(finance_sig_url_to_use)
   voucher_template_data['finance_signature_html'] = f'<img src="data:{finance_mime};base64,{finance_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if finance_sig_b64 else ""


   # Approved By
   approved_by_name_on_form = form_data_from_html_form.get('approved_by_name', '').strip() # This is readonly, from original approval type
   voucher_template_data['approved_by_name'] = approved_by_name_on_form if approved_by_name_on_form else approver_data_sheet2.get('approved_by_name_default', 'N/A')
   approved_by_sig_url_to_use = approver_data_sheet2.get("prepared_by_signature_urls_map", {}).get(voucher_template_data['approved_by_name'], approver_data_sheet2.get('approved_by_signature_url', ''))
   approved_by_sig_b64, approved_by_mime, _ = get_signature_data_from_url(approved_by_sig_url_to_use)
   voucher_template_data['approved_by_signature_html'] = f'<img src="data:{approved_by_mime};base64,{approved_by_sig_b64}" style="{SIG_IMG_HTML_STYLE}">' if approved_by_sig_b64 else ""




   env = Environment(loader=FileSystemLoader('templates'), cache_size=0, auto_reload=True)
   template = env.get_template('voucher_template.html')
   html_for_voucher_pdf = template.render(voucher_data_from_form=voucher_template_data)


   voucher_only_pdf_bytes = None
   merged_pdf_bytes = None
   final_merged_pdf_url = None
   # temp_dir = tempfile.mkdtemp(prefix='voucher_gen_') # Not needed if not saving voucher-only PDF locally


   try:
       pdf_options = {'encoding': 'UTF-8', 'quiet': '', 'page-size': 'A4', 'margin-top': '10mm',
                      'margin-bottom': '10mm', 'margin-left': '10mm', 'margin-right': '10mm'}
       if PDFKIT_CONFIG:
           logging.info("Attempting Voucher PDF generation using pdfkit.")
           voucher_only_pdf_bytes = pdfkit.from_string(html_for_voucher_pdf, False, configuration=PDFKIT_CONFIG, options=pdf_options)
       elif HTML:
           logging.info("pdfkit config missing. Attempting Voucher PDF generation with weasyprint.")
           voucher_only_pdf_bytes = HTML(string=html_for_voucher_pdf).write_pdf()
       else:
           return jsonify({"success": False, "message": "No PDF generation tool (pdfkit/weasyprint) available."}), 500


       if not voucher_only_pdf_bytes:
           return jsonify({"success": False, "message": "Failed to generate voucher-only PDF content (empty)."}), 500


       # Merge with Original Request PDF
       merger = PdfMerger()
       merger.append(BytesIO(voucher_only_pdf_bytes))


       original_request_pdf_drive_link = original_req_data_from_sheet.get('Request PDF Link')
       if original_request_pdf_drive_link and ("drive.google.com" in original_request_pdf_drive_link or re.match(r'^[a-zA-Z0-9_-]{25,}$', original_request_pdf_drive_link)):
           try:
               logging.info(f"Downloading original request PDF from Drive: {original_request_pdf_drive_link}")
               downloaded_request_pdf_bytes = download_drive_file_bytes(original_request_pdf_drive_link, creds)
               merger.append(BytesIO(downloaded_request_pdf_bytes))
           except Exception as e:
               logging.warning(f"Could not download/append original request PDF ({original_request_pdf_drive_link}) to voucher: {e}", exc_info=True)
               # Decide if this is critical. For now, we proceed with voucher-only if original fails.
       else:
           logging.warning(f"Original request PDF link missing, invalid, or not a Drive link: {original_request_pdf_drive_link}. Voucher will not include it.")


       merged_output_stream = BytesIO()
       merger.write(merged_output_stream)
       merger.close()
       merged_pdf_bytes = merged_output_stream.getvalue()


       merged_pdf_filename = f'Voucher_Merged_{request_id}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
       final_merged_pdf_url = upload_file_from_bytes(
           file_content=merged_pdf_bytes,
           file_name=merged_pdf_filename,
           mime_type='application/pdf'
       )
       if not final_merged_pdf_url:
           return jsonify({"success": False, "message": "Failed to upload merged voucher PDF to Drive"}), 500


       # Update Sheet1 with the voucher link and generation timestamp
       updated_sheet, update_msg = update_sheet_status(
           request_id,
           voucher_link=final_merged_pdf_url,
           voucher_generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           voucher_approved_by="",  # Cleared when voucher is generated
           voucher_rejection_reason="",
           voucher_prepared_by=selected_prepared_by_name  # This correctly updates the sheet
       )


       if not updated_sheet:
           logging.error(f"Failed to update sheet with voucher link for Request ID {request_id}: {update_msg}")
           # This is problematic: voucher generated and uploaded, but sheet not updated.
           # Consider how to handle this (e.g., manual alert).
           return jsonify({"success": False, "message": f"Voucher generated, but failed to update sheet: {update_msg}. Manual check required."}), 500


       logging.info(f"Successfully generated, merged, and uploaded voucher for {request_id}. Final PDF URL: {final_merged_pdf_url}")
       return jsonify({"success": True, "voucher_url": final_merged_pdf_url, "request_id": request_id})


   except HttpError as e: # Google API errors
       logging.error(f"Google API HttpError in generate_voucher_route: {e.resp.status} - {e._get_reason()}", exc_info=True)
       return jsonify({"success": False, "message": f"Google API Error: {e._get_reason()}"}), 500
   except ValueError as ve: # Type conversion errors
       logging.error(f"ValueError in generate_voucher_route: {ve}", exc_info=True)
       return jsonify({"success": False, "message": str(ve)}), 400
   except Exception as e: # Other unexpected errors
       import traceback
       logging.error(f"Unexpected error in generate_voucher_route: {traceback.format_exc()}")
       return jsonify({"success": False, "message": f"An unexpected server error occurred: {str(e)}"}), 500

if __name__ == '__main__':
   if not os.path.exists(UPLOAD_FOLDER):
       try:
           os.makedirs(UPLOAD_FOLDER);
           logging.info(f"Created upload folder: {UPLOAD_FOLDER}")
       except Exception as e:
           logging.error(f"Failed to create UPLOAD_FOLDER {UPLOAD_FOLDER}: {e}", exc_info=True)


   # For local development with Google OAuth, OAUTHLIB_INSECURE_TRANSPORT is often needed
   # if your redirect URI is http and not https. Render typically handles HTTPS.
   # Check if running locally or on Render before setting this.
   # if os.environ.get("RENDER") is None: # Example: only set if not on Render
   os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
   logging.warning("OAUTHLIB_INSECURE_TRANSPORT is enabled. DO NOT USE THIS IN A PRODUCTION ENVIRONMENT WITHOUT HTTPS.")


   port = int(os.environ.get('PORT', 10000)) # Use Render's port or default to 5000
   logging.info(f"Starting Flask development server on port {port}...")
   app.run(debug=True, host='0.0.0.0', port=port, use_reloader=True)
