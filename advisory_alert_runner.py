import os
import sys
import time
import hashlib
import webbrowser
from datetime import datetime
from email.message import EmailMessage
import openpyxl

# Ensure UTF-8 console output on all Windows machines to prevent charmap/cp1252 encoding crashes
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Try importing Windows COM for Classic Outlook & Excel Macro Auto-Refresh
try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

# =============================================================================
# CONFIGURATION
# =============================================================================
# Primary Tracker Workbook (.xlsm / .xlsx)
# Checks for 'Advisory_Tracker.xlsm' first, falls back to 'dummy_advisory.xlsx'
EXCEL_FILE = "Advisory_Tracker.xlsm" if os.path.exists("Advisory_Tracker.xlsm") else "dummy_advisory.xlsx"
SHEET_NAME = "Advisory"
MSG_OUTPUT_DIR = "output_msg"
TRACKING_FILE = ".processed_ids.txt"

# Automation Schedule: Set to 3 for 3-hour continuous check (0 = run once and exit)
SCHEDULE_INTERVAL_HOURS = 0

# Macro & Power Query Auto-Refresh before reading (Refreshes real-time SharePoint data)
ENABLE_EXCEL_AUTO_REFRESH = True
AUTO_REFRESH_WAIT_SECONDS = 20

# Safe Testing & Dispatch Controls:
# - Set AUTO_SEND = False to OPEN the email on screen in Outlook for manual review
# - Set AUTO_SEND = True to send immediately in the background
AUTO_SEND = True

# Maximum number of alerts to process per cycle (0 = unlimited, processes ALL new alerts found)
MAX_ALERTS_PER_CYCLE = 0

# ONE-TIME INITIAL BASELINE SEEDING:
# Set to True if your Excel sheet already has hundreds of old past alerts and you ONLY
# want to start alerting for NEW alerts added from today onwards (marks all existing as processed without sending emails).
SEED_ALL_EXISTING = False

# Recipient configuration for advisories and master trigger email
EMAIL_TO = "tejesh988@outlook.com"
EMAIL_CC = ""
# Optional: Specify sender or shared mailbox name (leave empty to use default logged-in Outlook profile)
EMAIL_SENT_ON_BEHALF = ""

# Excel Column Indices (1-based from tracker sheet)
COL_TITLE = 1             # Title
COL_TYPE = 2              # Advisory Type
COL_ADVISORY_NO = 4       # Advisory No (e.g. TI26-090)
COL_DATE = 5              # Advisory Preparation Date
COL_SUMMARY = 7           # Summary
COL_ATTACK_VECTOR = 8     # Attack Vector
COL_IMPACT_ANALYSIS = 9   # Impact Analysis
COL_IOCS = 10             # IOCs
COL_RECOMMENDATION = 11   # Recommendation / Vendor Solution
COL_REFERENCE = 12        # Reference
COL_CVE = 13              # CVE Id
COL_IMPACTED_ELEMENTS = 14# Impacted Elements
COL_ATTACK_TYPE = 15      # Attack type / Impact Type
COL_SEVERITY = 18         # Severity

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def get_processed_ids():
    if not os.path.exists(TRACKING_FILE):
        return set()
    with open(TRACKING_FILE, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())

def mark_as_processed(row_id):
    with open(TRACKING_FILE, "a", encoding="utf-8") as f:
        f.write(f"{row_id}\n")

def safe_str(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%b %dth %Y, %H:%M (CET)")
    return str(val).strip()

def get_row_hash(row_values):
    row_str = "|".join(safe_str(val) for val in row_values)
    return hashlib.md5(row_str.encode("utf-8")).hexdigest()

def format_bullet_list(text):
    if not text:
        return ""
    lines = [line.strip("- *• \t\r\n") for line in text.splitlines() if line.strip()]
    if not lines:
        return text
    items = "".join(f"<li style='margin-bottom: 5px;'>{line}</li>" for line in lines)
    return f"<ul style='margin: 0; padding-left: 20px;'>{items}</ul>"

def get_severity_metrics(severity_text):
    sev = severity_text.lower()
    if "critical" in sev:
        return "#C00000", "Critical", 10, 315
    elif "high" in sev:
        return "#E26B00", "High", 8, 212
    elif "medium" in sev:
        return "#FFC000", "Medium", 5, 137
    elif "low" in sev:
        return "#00B050", "Low", 2, 62
    return "#1F4E79", severity_text, 1, 62

# =============================================================================
# EXCEL MACRO & POWER QUERY AUTO-REFRESH (WINDOWS COM)
# =============================================================================
def trigger_excel_macro_refresh(file_path):
    """Triggers Excel 'Refresh All' to fetch real-time SharePoint data before scanning."""
    if not COM_AVAILABLE or not ENABLE_EXCEL_AUTO_REFRESH:
        return
    
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        return
        
    print(f"[*] Triggering Excel Macro / Power Query 'Refresh All' on {os.path.basename(file_path)}...")
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        wb = excel.Workbooks.Open(abs_path)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        
        print(f"[*] Waiting {AUTO_REFRESH_WAIT_SECONDS}s for real-time SharePoint synchronization...")
        time.sleep(AUTO_REFRESH_WAIT_SECONDS)
        
        wb.Save()
        wb.Close()
        excel.Quit()
        print("[+] Excel real-time data refreshed and saved successfully.")
    except Exception as e:
        print(f"[*] Note: Excel COM refresh skipped or not available on this host ({e}).")

# =============================================================================
# HEADER BANNER EMBEDDER (TI_BG.png with fallback & CID support)
# =============================================================================
def ensure_banner_file():
    # Check for original image files in the directory
    possible_names = ["TI_BG.png", "TI_BG.jpg", "TI_BG.jpeg", "header.png", "banner.png", "Capgemini_banner.png"]
    for fname in possible_names:
        if os.path.exists(fname):
            return fname
            
    # If missing on another machine, generate a branded banner using Pillow as fallback
    fallback_file = "TI_BG.png"
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("RGB", (860, 110), color="#001833")
        draw = ImageDraw.Draw(img)
        # Draw gradient/accent bars
        draw.rectangle([0, 0, 860, 6], fill="#0070AD")
        draw.rectangle([0, 104, 860, 110], fill="#005A9C")
        draw.text((25, 20), "Cyber Security Unit", fill="#FFFFFF")
        draw.text((25, 55), "Threat Intelligence", fill="#7EC8E3")
        img.save(fallback_file, "PNG")
        return fallback_file
    except Exception:
        return None

def get_header_banner_html(use_cid=False):
    banner_file = ensure_banner_file()
    if use_cid and banner_file and os.path.exists(banner_file):
        return """<div style="width: 100%; text-align: center; background-color: #001833; line-height: 0;">
            <img src="cid:header_banner" alt="Cyber Security Unit Threat Intelligence" style="width: 100%; max-height: 125px; display: block; border: 0;" />
        </div>"""
        
    if banner_file and os.path.exists(banner_file):
        try:
            import base64
            with open(banner_file, "rb") as f:
                b64_data = base64.b64encode(f.read()).decode("utf-8")
            return f"""<div style="width: 100%; text-align: center; background-color: #001833; line-height: 0;">
                <img src="data:image/png;base64,{b64_data}" alt="Cyber Security Unit Threat Intelligence" style="width: 100%; max-height: 125px; display: block; border: 0;" />
            </div>"""
        except Exception:
            pass
            
    # Clean CSS fallback if image file not available
    return """<div class="header-banner">
        <h1>Cyber Security Unit</h1>
        <h2>Threat Intelligence</h2>
    </div>"""

# =============================================================================
# HTML EMAIL TEMPLATE BUILDER (MATCHING EXACT CSU ADVISORY & DUAL THREAT METER)
# =============================================================================
def build_advisory_html(data, use_cid=True):
    sev_color, sev_label, threat_score, meter_pos = get_severity_metrics(data["severity"])
    header_banner_html = get_header_banner_html(use_cid=use_cid)
    
    impacted_html = format_bullet_list(data["impacted_elements"])
    solution_html = format_bullet_list(data["recommendation"])
    summary_html = data["summary"].replace("\n", "<br>")
    impact_html = data["impact_analysis"].replace("\n", "<br>")

    # Generate Reference Links
    ref_raw = data.get("reference", "")
    ref_links = []
    for line in ref_raw.splitlines():
        line = line.strip()
        if line.startswith("http://") or line.startswith("https://"):
            ref_links.append(f"<a href='{line}' target='_blank' style='color: #0066cc; text-decoration: underline;'>{line}</a>")
        elif line:
            ref_links.append(line)
    ref_html = "<br>".join(ref_links) if ref_links else "Refer to official vendor portal."

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
</head>
<body style="font-family: Arial, Calibri, sans-serif; font-size: 12px; color: #000000; background-color: #e9ecef; margin: 0; padding: 15px;">
    <div style="max-width: 860px; margin: 0 auto; background: #ffffff; border: 2px solid #1F4E79;">
        <!-- Banner Image -->
        {header_banner_html}

        <!-- Title Header Bar -->
        <div style="background-color: #EBF1F5; text-align: center; padding: 8px; border-top: 1px solid #1F4E79; border-bottom: 1px solid #1F4E79;">
            <h3 style="margin: 0; font-size: 15px; color: #1F4E79; font-weight: bold; font-family: Arial, sans-serif;">Information Security Advisory - Alert</h3>
            <p style="margin: 3px 0 0 0; font-size: 12px; color: #222222; font-weight: bold; font-family: Arial, sans-serif;">Date &amp; Time Issued: {data['date']}</p>
        </div>

        <!-- Main Metadata Table -->
        <table width="100%" cellpadding="8" cellspacing="0" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">
            <tr>
                <td width="22%" bgcolor="#EBF1F5" style="width: 22%; background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Advisory Number</td>
                <td width="78%" bgcolor="#FFFFFF" style="width: 78%; background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;"><b>{data['advisory_no']}</b></td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Title</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;"><b>{data['title']}</b></td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Impacted Elements</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{impacted_html or "Refer to technical analysis."}</td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Summary</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{summary_html}</td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Severity</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">
                    <span style="font-weight: bold; color: {sev_color}; font-size: 13px;">{data['severity']}</span>
                    <span style="display: inline-block; width: 32px; height: 7px; background-color: {sev_color}; margin-left: 8px; vertical-align: middle; border-radius: 2px;"></span>
                </td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Impact Type</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{data['attack_type'] or "Arbitrary code execution"}</td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Impact Analysis</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{impact_html}</td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Vendor Solution</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{solution_html or "Apply vendor patches immediately."}</td>
            </tr>
        </table>

        <!-- Threat Meter Header -->
        <div style="background-color: #EBF1F5; text-align: center; padding: 8px; border-top: 1px solid #1F4E79; border-bottom: 1px solid #1F4E79;">
            <h3 style="margin: 0; font-size: 15px; color: #1F4E79; font-weight: bold; font-family: Arial, sans-serif;">Threat Meter</h3>
        </div>

        <!-- DUAL THREAT METERS (VENDOR METER + GSOC METER) -->
        <div style="padding: 16px 10px 10px 10px; background-color: #ffffff; border-left: 1px solid #000; border-right: 1px solid #000;">
            <table width="100%" cellpadding="0" cellspacing="0" style="width: 100%; border: none; border-collapse: collapse;">
                <tr>
                    <!-- Left: Vendor Meter -->
                    <td style="width: 50%; vertical-align: top; text-align: center; border: none; padding: 0 10px;">
                        <svg width="340" height="95" viewBox="0 0 340 95" style="display: block; margin: 0 auto;">
                            <!-- Indicator Top Pointer -->
                            <text x="{meter_pos}" y="13" font-family="Arial" font-size="11" font-weight="bold" fill="#111" text-anchor="middle">{threat_score}</text>
                            <polygon points="{meter_pos-6},15 {meter_pos+6},15 {meter_pos},25" fill="#BFBFBF" stroke="#333333" stroke-width="1.5" />
                            
                            <!-- Color Blocks -->
                            <!-- Not Critical (1-4) -->
                            <rect x="25" y="26" width="75" height="24" fill="#0070C0" stroke="#000000" stroke-width="1.2" />
                            <text x="62" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Not Critical</text>
                            
                            <!-- Medium (4-7) -->
                            <rect x="100" y="26" width="75" height="24" fill="#ED7D31" stroke="#000000" stroke-width="1.2" />
                            <text x="137" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Medium</text>
                            
                            <!-- High (7-9) -->
                            <rect x="175" y="26" width="75" height="24" fill="#C00000" stroke="#000000" stroke-width="1.2" />
                            <text x="212" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">High</text>
                            
                            <!-- Critical (9-10) -->
                            <rect x="250" y="26" width="65" height="24" fill="#CC3399" stroke="#000000" stroke-width="1.2" />
                            <text x="282" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Critical</text>
                            
                            <!-- Base Axis Line -->
                            <line x1="12" y1="56" x2="330" y2="56" stroke="#000000" stroke-width="2.5" />
                            <polygon points="326,52 336,56 326,60" fill="#000000" />
                            
                            <!-- Axis Ticks & Scale Numbers -->
                            <line x1="25" y1="48" x2="25" y2="64" stroke="#000" stroke-width="2" />
                            <text x="25" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">1</text>
                            
                            <line x1="100" y1="48" x2="100" y2="64" stroke="#000" stroke-width="2" />
                            <text x="100" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">4</text>
                            
                            <line x1="175" y1="48" x2="175" y2="64" stroke="#000" stroke-width="2" />
                            <text x="175" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">7</text>
                            
                            <line x1="250" y1="48" x2="250" y2="64" stroke="#000" stroke-width="2" />
                            <text x="250" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">9</text>
                            
                            <line x1="315" y1="48" x2="315" y2="64" stroke="#000" stroke-width="2" />
                            <text x="315" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">10</text>
                        </svg>
                        <div style="font-family: Arial, sans-serif; font-weight: bold; font-size: 13px; color: #000; margin-top: 4px;">Vendor Meter</div>
                    </td>
                    
                    <!-- Right: GSOC Meter -->
                    <td style="width: 50%; vertical-align: top; text-align: center; border: none; padding: 0 10px;">
                        <svg width="340" height="95" viewBox="0 0 340 95" style="display: block; margin: 0 auto;">
                            <!-- Indicator Top Pointer -->
                            <text x="{meter_pos}" y="13" font-family="Arial" font-size="11" font-weight="bold" fill="#111" text-anchor="middle">{threat_score}</text>
                            <polygon points="{meter_pos-6},15 {meter_pos+6},15 {meter_pos},25" fill="#BFBFBF" stroke="#333333" stroke-width="1.5" />
                            
                            <!-- Color Blocks -->
                            <!-- Low (1-4) -->
                            <rect x="25" y="26" width="75" height="24" fill="#0070C0" stroke="#000000" stroke-width="1.2" />
                            <text x="62" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Low</text>
                            
                            <!-- Medium (4-7) -->
                            <rect x="100" y="26" width="75" height="24" fill="#ED7D31" stroke="#000000" stroke-width="1.2" />
                            <text x="137" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Medium</text>
                            
                            <!-- High (7-9) -->
                            <rect x="175" y="26" width="75" height="24" fill="#C00000" stroke="#000000" stroke-width="1.2" />
                            <text x="212" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">High</text>
                            
                            <!-- Critical (9-10) -->
                            <rect x="250" y="26" width="65" height="24" fill="#CC3399" stroke="#000000" stroke-width="1.2" />
                            <text x="282" y="42" font-family="Arial" font-size="10" font-weight="bold" fill="#ffffff" text-anchor="middle">Critical</text>
                            
                            <!-- Base Axis Line -->
                            <line x1="12" y1="56" x2="330" y2="56" stroke="#000000" stroke-width="2.5" />
                            <polygon points="326,52 336,56 326,60" fill="#000000" />
                            
                            <!-- Axis Ticks & Scale Numbers -->
                            <line x1="25" y1="48" x2="25" y2="64" stroke="#000" stroke-width="2" />
                            <text x="25" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">1</text>
                            
                            <line x1="100" y1="48" x2="100" y2="64" stroke="#000" stroke-width="2" />
                            <text x="100" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">4</text>
                            
                            <line x1="175" y1="48" x2="175" y2="64" stroke="#000" stroke-width="2" />
                            <text x="175" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">7</text>
                            
                            <line x1="250" y1="48" x2="250" y2="64" stroke="#000" stroke-width="2" />
                            <text x="250" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">9</text>
                            
                            <line x1="315" y1="48" x2="315" y2="64" stroke="#000" stroke-width="2" />
                            <text x="315" y="76" font-family="Arial" font-size="11" font-weight="bold" fill="#000" text-anchor="middle">10</text>
                        </svg>
                        <div style="font-family: Arial, sans-serif; font-weight: bold; font-size: 13px; color: #000; margin-top: 4px;">GSOC Meter</div>
                    </td>
                </tr>
            </table>
        </div>

        <!-- GSOC Assessment Section Header -->
        <div style="background-color: #EBF1F5; text-align: center; padding: 8px; border-top: 1px solid #1F4E79; border-bottom: 1px solid #1F4E79;">
            <h3 style="margin: 0; font-size: 15px; color: #1F4E79; font-weight: bold; font-family: Arial, sans-serif;">GSOC Assessment</h3>
        </div>

        <!-- GSOC Assessment Details Table -->
        <table width="100%" cellpadding="8" cellspacing="0" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">
            <tr>
                <td width="22%" bgcolor="#EBF1F5" style="width: 22%; background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">Exploitation Probability</td>
                <td width="78%" bgcolor="#FFFFFF" style="width: 78%; background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;"><b>{data['severity']}</b></td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">GSOC Risk Assessment</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">
                    Successful exploitation of these vulnerabilities could lead to {data['attack_type'] or "system compromise"}. Hence it has been categorized as <b>{data['severity']}</b>. No active exploitation detected so far.
                </td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">GSOC Recommendation</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">
                    <div>Apply the latest patch released by the vendor.</div>
                    <div style="margin-top: 6px;">{solution_html}</div>
                    <div style="background-color: #FFFF00; border: 1px solid #cccc00; color: #000000; font-weight: bold; font-size: 11px; padding: 4px 8px; margin-top: 8px; display: inline-block;">/* Test changes on non-production systems before applying on production systems */</div>
                </td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">References</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">{ref_html}</td>
            </tr>
            <tr>
                <td bgcolor="#EBF1F5" style="background-color: #EBF1F5; color: #1F4E79; font-weight: bold; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;">References CVE's</td>
                <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; color: #000000; padding: 8px 12px; border: 1px solid #000000; vertical-align: top;"><b>{data['cve'] or "N/A"}</b></td>
            </tr>
        </table>

        <!-- Capgemini CSU Footer -->
        <div style="background-color: #002855; color: #ffffff; text-align: center; padding: 10px 15px; font-size: 10.5px; line-height: 1.5; border-top: 1px solid #1F4E79; font-family: Arial, sans-serif;">
            <div>The information contained in this message is proprietary and confidential. It is for Capgemini and its customers only.</div>
            <div>Copyright &copy; 2024. All rights reserved by Capgemini.</div>
            <div style="font-style: italic; margin-top: 2px;">Collaborative Business Experience&trade;</div>
        </div>
    </div>
</body>
</html>"""
    return html

# =============================================================================
# OUTLOOK .MSG & TRIGGER EMAIL DISPATCHER
# =============================================================================
def create_individual_advisory(outlook, data, output_path):
    subject = f"CSU Threat Intelligence Notification Advisory {data['advisory_no']} - {data['title']} ({data['severity'].upper()})"
    banner_file = ensure_banner_file()
    
    # Always generate a clean HTML version for universal browser & webmail viewing
    html_path = output_path.replace(".msg", ".html")
    html_general = build_advisory_html(data, use_cid=False)
    with open(html_path, "w", encoding="utf-8") as hf:
        hf.write(html_general)
    
    # 1. If Classic Outlook COM is available, create native Outlook .msg item in Read/Received Mode
    if outlook is not None:
        try:
            html_content = build_advisory_html(data, use_cid=True)
            mail_item = outlook.CreateItem(0)
            mail_item.Subject = subject
            mail_item.To = EMAIL_TO
            if EMAIL_CC:
                mail_item.CC = EMAIL_CC
            if EMAIL_SENT_ON_BEHALF:
                mail_item.SentOnBehalfOfName = EMAIL_SENT_ON_BEHALF
            
            # Embed banner image as inline CID attachment for Outlook Desktop
            if banner_file and os.path.exists(banner_file):
                abs_banner = os.path.abspath(banner_file)
                att = mail_item.Attachments.Add(abs_banner, 1, 0, "Header Banner")
                att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "header_banner")
            
            mail_item.HTMLBody = html_content
            
            # Set PR_MESSAGE_FLAGS (0x0E070003) = 1 (MSGFLAG_READ)
            # This turns the unsent draft into a normal 'Received' email with Reply / Reply All buttons
            try:
                mail_item.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
            except Exception:
                pass
                
            abs_path = os.path.abspath(output_path)
            mail_item.SaveAs(abs_path, 3) # 3 = olMSG
            mail_item.Close(1)
            return abs_path
        except Exception as e:
            print(f"    [!] Classic Outlook COM notice ({e}), falling back to universal email format.")
    
    # 2. Universal RFC MIME email package (.msg & .eml for New Outlook, Web, and any system)
    msg_path = output_path if output_path.endswith(".msg") else output_path + ".msg"
    adv_email = EmailMessage()
    adv_email["Subject"] = subject
    adv_email["From"] = EMAIL_SENT_ON_BEHALF or "CSU Threat Intelligence <csu-alerts@capgemini.com>"
    adv_email["To"] = EMAIL_TO
    if EMAIL_CC:
        adv_email["Cc"] = EMAIL_CC
    adv_email["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0000")
    adv_email.set_content(f"CSU Threat Intelligence Advisory {data['advisory_no']}: {data['title']}")
    adv_email.add_alternative(html_general, subtype="html")
    
    abs_path = os.path.abspath(msg_path)
    with open(abs_path, "wb") as f:
        f.write(adv_email.as_bytes())
        
    # Also save .eml version for default Windows Mail / New Outlook
    eml_path = output_path.replace(".msg", ".eml")
    with open(eml_path, "wb") as f:
        f.write(adv_email.as_bytes())
        
    return abs_path

def send_master_trigger_email(outlook, generated_files):
    if not generated_files:
        return

    date_str = datetime.now().strftime("%d %b %Y %H:%M")
    count = len(generated_files)
    
    file_list_html = "".join(f"<li style='margin-bottom: 5px;'><b>{os.path.basename(f)}</b></li>" for f in generated_files)
    
    subject = f"CSU Threat Intelligence Notification Advisory - {count} New Alert(s) [{date_str}]"
    
    trigger_html = f"""<!DOCTYPE html>
<html>
<body style="font-family: Arial, Calibri, sans-serif; font-size: 13px; color: #333333; line-height: 1.6;">
    <p>Hi Team,</p>
    <p>Please find attached <b>{count} new Threat Intelligence Security Advisories</b> generated by the CSU Threat Intelligence Automation Engine on <b>{date_str}</b>.</p>
    <ul style="margin: 12px 0; padding-left: 25px;">
        {file_list_html}
    </ul>
    <br>
    <p>Best regards,<br>
    <strong style="color: #1F4E79;">Cyber Security Unit (CSU)</strong><br>
    Threat Intelligence Team</p>
</body>
</html>
"""
    
    # Build the full MIME email message
    eml_msg = EmailMessage()
    eml_msg["To"] = EMAIL_TO
    if EMAIL_CC:
        eml_msg["Cc"] = EMAIL_CC
    eml_msg["Subject"] = subject
    eml_msg.set_content("Please find attached the new Threat Intelligence Security Advisories.")
    eml_msg.add_alternative(trigger_html, subtype="html")

    # Attach all generated .msg advisory files
    for f in generated_files:
        with open(f, "rb") as fp:
            file_data = fp.read()
            eml_msg.add_attachment(
                file_data, 
                maintype="application", 
                subtype="vnd.ms-outlook", 
                filename=os.path.basename(f)
            )

    # 1. OPTIONAL: Direct SMTP Sending (If smtp_config.py is configured with credentials)
    try:
        import smtp_config
        smtp_user = getattr(smtp_config, "SMTP_USERNAME", "").strip()
        smtp_pass = getattr(smtp_config, "SMTP_PASSWORD", "").strip()
        if smtp_user and smtp_pass:
            import smtplib
            eml_msg["From"] = smtp_user
            smtp_server = getattr(smtp_config, "SMTP_SERVER", "smtp-mail.outlook.com")
            smtp_port = getattr(smtp_config, "SMTP_PORT", 587)
            
            print(f"[*] Transmitting live email via SMTP ({smtp_server}:{smtp_port}) to {EMAIL_TO}...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.ehlo()
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.send_message(eml_msg)
            server.quit()
            print(f"\n[+] SUCCESS: Live email transmitted via SMTP to {EMAIL_TO} with {count} .msg attachment(s)!")
            return
    except ImportError:
        pass
    except Exception as e:
        print(f"[!] SMTP sending note: {e}")

    # 2. Dispatch via Classic Outlook COM if available (Production host)
    if outlook is not None:
        try:
            trigger_mail = outlook.CreateItem(0)
            trigger_mail.To = EMAIL_TO
            if EMAIL_CC:
                trigger_mail.CC = EMAIL_CC
            if EMAIL_SENT_ON_BEHALF:
                trigger_mail.SentOnBehalfOfName = EMAIL_SENT_ON_BEHALF
            trigger_mail.Subject = subject
            trigger_mail.HTMLBody = trigger_html
            for f in generated_files:
                trigger_mail.Attachments.Add(os.path.abspath(f))
            if AUTO_SEND:
                trigger_mail.Send()
                print(f"\n[+] Master trigger email SENT automatically via Outlook to {EMAIL_TO} with {count} .msg attachment(s)!")
            else:
                trigger_mail.Display()
                print(f"\n[+] Master trigger email opened in Classic Outlook with {count} .msg attachment(s)!")
            return
        except Exception as e:
            print(f"[!] Note: Classic Outlook COM dispatch: {e}")

    # 3. Universal Trigger Package (Local testing with New Outlook / Desktop Mail)
    eml_path = os.path.join(MSG_OUTPUT_DIR, "Master_Trigger_Email.eml")
    with open(eml_path, "wb") as fp:
        fp.write(eml_msg.as_bytes())

    summary_path = os.path.join(MSG_OUTPUT_DIR, "Master_Trigger_Email_Summary.html")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(trigger_html)

    print(f"\n[+] Master Trigger email generated with .msg attachments: {eml_path}")
    try:
        os.startfile(os.path.abspath(eml_path))
        print(f"[+] Opened trigger email in Outlook: {eml_path}")
    except Exception:
        webbrowser.open(os.path.abspath(summary_path))

# =============================================================================
# CYCLE EXECUTION FLOW (HANDLES 25K+ ROWS EFFICIENTLY WITH PROGRESS BAR)
# =============================================================================
def draw_progress_bar(current, total, bar_length=20):
    fraction = current / total if total > 0 else 1
    filled = int(bar_length * fraction)
    bar = "=" * filled + "." * (bar_length - filled)
    percent = int(fraction * 100)
    return f"[{bar}] {percent:>3}%"

def run_cycle():
    start_time = time.time()
    now_str = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
    print("\n" + "=" * 65)
    print(f"  CSU Threat Intelligence Alert Automation Cycle  [{now_str}]")
    print("=" * 65)
    
    ensure_dir(MSG_OUTPUT_DIR)
    processed_ids = get_processed_ids()
    
    # 1. Trigger live SharePoint macro refresh
    trigger_excel_macro_refresh(EXCEL_FILE)
    
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] Excel file '{EXCEL_FILE}' not found.")
        return 0

    print(f"[*] Scanning workbook '{EXCEL_FILE}' (fast read-only mode for 25k+ rows)...")
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    except Exception as e:
        print(f"[ERROR] Failed to load Excel workbook: {e}")
        return 0
    
    if SHEET_NAME not in wb.sheetnames:
        print(f"[ERROR] Sheet '{SHEET_NAME}' not found in workbook.")
        wb.close()
        return 0
        
    sheet = wb[SHEET_NAME]
    
    outlook = None
    if COM_AVAILABLE:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception:
            outlook = None

    unprocessed_alerts = []
    total_scanned = 0
    total_alerts_found = 0
    
    # Step 1: Fast scan to identify all new Alert records
    for row in sheet.iter_rows(min_row=2, values_only=True):
        total_scanned += 1
        if not row or len(row) < 2:
            continue
            
        advisory_type = safe_str(row[COL_TYPE - 1]).lower()
        
        # TARGET ONLY ALERTS
        if advisory_type != "alert":
            continue
            
        total_alerts_found += 1
        row_hash = get_row_hash(row)
        if row_hash in processed_ids:
            continue # Already processed in previous cycle
            
        title = safe_str(row[COL_TITLE - 1])
        if not title:
            continue
            
        advisory_no = safe_str(row[COL_ADVISORY_NO - 1]) or "TI-Alert"
        
        data = {
            "title": title,
            "advisory_no": advisory_no,
            "date": safe_str(row[COL_DATE - 1]) or datetime.now().strftime("%b %dth %Y, %H:%M (CET)"),
            "severity": safe_str(row[COL_SEVERITY - 1]) or "Medium",
            "cve": safe_str(row[COL_CVE - 1]),
            "summary": safe_str(row[COL_SUMMARY - 1]),
            "impact_analysis": safe_str(row[COL_IMPACT_ANALYSIS - 1]),
            "recommendation": safe_str(row[COL_RECOMMENDATION - 1]),
            "reference": safe_str(row[COL_REFERENCE - 1]),
            "impacted_elements": safe_str(row[COL_IMPACTED_ELEMENTS - 1]),
            "attack_type": safe_str(row[COL_ATTACK_TYPE - 1]),
            "row_hash": row_hash
        }
        unprocessed_alerts.append(data)

    wb.close()
    elapsed = time.time() - start_time
    print(f"[+] Scanned {total_scanned:,} total rows in {elapsed:.2f}s.")
    print(f"[*] Total 'Alert' rows in tracker: {total_alerts_found} | New unprocessed alerts: {len(unprocessed_alerts)}")

    if not unprocessed_alerts:
        print("\n[*] All alerts are up-to-date. No new emails to send.")
        return 0

    # Step 1.5: Handle One-Time Baseline Seeding (Skip old history)
    if SEED_ALL_EXISTING:
        print(f"\n[!] SEED MODE ACTIVE: Marking all {len(unprocessed_alerts)} existing alerts as processed...")
        for data in unprocessed_alerts:
            mark_as_processed(data["row_hash"])
        print(f"[+] All {len(unprocessed_alerts)} historical alerts recorded in {TRACKING_FILE}.")
        print("[*] Baseline seeding complete! Now set SEED_ALL_EXISTING = False to begin live alerting.")
        return 0

    # Step 1.6: Apply Batch Limit (Prevents attaching 3k files to one email)
    if MAX_ALERTS_PER_CYCLE > 0 and len(unprocessed_alerts) > MAX_ALERTS_PER_CYCLE:
        print(f"[*] Batch Limit Active: Processing first {MAX_ALERTS_PER_CYCLE} of {len(unprocessed_alerts)} new alerts.")
        print(f"    (Remaining {len(unprocessed_alerts) - MAX_ALERTS_PER_CYCLE} alerts will automatically process in subsequent cycles).")
        unprocessed_alerts = unprocessed_alerts[:MAX_ALERTS_PER_CYCLE]

    # Step 2: Process each new alert with a visual progress bar
    print(f"\n[*] Processing {len(unprocessed_alerts)} alert advisory package(s):")
    print("-" * 65)
    
    generated_files = []
    total_to_process = len(unprocessed_alerts)
    
    for idx, data in enumerate(unprocessed_alerts, start=1):
        progress = draw_progress_bar(idx, total_to_process)
        clean_title = "".join(c for c in data["title"] if c.isalnum() or c in " _-")[:35].strip()
        msg_filename = f"{data['advisory_no']}_Alert_{clean_title}.msg"
        msg_path = os.path.join(MSG_OUTPUT_DIR, msg_filename)
        
        print(f"  {progress} [{idx}/{total_to_process}] Generating: {data['advisory_no']} - {clean_title}...")
        
        try:
            saved_file = create_individual_advisory(outlook, data, msg_path)
            generated_files.append(saved_file)
            mark_as_processed(data["row_hash"])
            processed_ids.add(data["row_hash"])
        except Exception as e:
            print(f"    [!] Error creating {data['advisory_no']}: {e}")

    # Step 3: Trigger master notification email with all packages attached
    if generated_files:
        print("-" * 65)
        print(f"[+] Successfully generated {len(generated_files)} advisory package(s).")
        print(f"[*] Transmitting master trigger notification email to {EMAIL_TO}...")
        send_master_trigger_email(outlook, generated_files)
        first_file = os.path.abspath(generated_files[0])
        print(f"[*] Opening preview: {first_file}")
        webbrowser.open(first_file)

    return len(generated_files)

def main():
    if SCHEDULE_INTERVAL_HOURS > 0:
        print(f"[*] Starting continuous automation engine (recurring every {SCHEDULE_INTERVAL_HOURS} hour(s))...")
        print("[*] Press Ctrl+C at any time to stop.\n")
        try:
            while True:
                run_cycle()
                next_check = datetime.fromtimestamp(time.time() + SCHEDULE_INTERVAL_HOURS * 3600).strftime("%H:%M:%S")
                print(f"\n" + "=" * 65)
                print(f"[*] Cycle finished. Next automated check at: {next_check} (sleeping {SCHEDULE_INTERVAL_HOURS}h)")
                print(f"[*] Keep this window open or minimized to run continuously.")
                print("=" * 65 + "\n")
                time.sleep(SCHEDULE_INTERVAL_HOURS * 3600)
        except KeyboardInterrupt:
            print("\n[*] Automation scheduler stopped by user.")
    else:
        run_cycle()
        print("\n" + "=" * 65)
        print("  Automation cycle completed.")
        print("=" * 65)

if __name__ == "__main__":
    main()
