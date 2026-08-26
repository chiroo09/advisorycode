import os
import sys
import time
import hashlib
import webbrowser
from datetime import datetime
from email.message import EmailMessage
import openpyxl

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
SCHEDULE_INTERVAL_HOURS = 3

# Macro & Power Query Auto-Refresh before reading (Refreshes real-time SharePoint data)
ENABLE_EXCEL_AUTO_REFRESH = True
AUTO_REFRESH_WAIT_SECONDS = 20

# Set to True to send automatically without manual click on production host
AUTO_SEND = True

# Recipient configuration for the master trigger email
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
# HEADER BANNER EMBEDDER (TI_BG.png)
# =============================================================================
def get_header_banner_html():
    banner_file = "TI_BG.png"
    if os.path.exists(banner_file):
        try:
            import base64
            with open(banner_file, "rb") as f:
                b64_data = base64.b64encode(f.read()).decode("utf-8")
            return f"""<div style="width: 100%; text-align: center; background-color: #001833; line-height: 0;">
                <img src="data:image/png;base64,{b64_data}" alt="Cyber Security Unit Threat Intelligence" style="width: 100%; max-height: 125px; display: block;" />
            </div>"""
        except Exception:
            pass
            
    # Clean CSS fallback if image file not found
    return """<div class="header-banner">
        <h1>Cyber Security Unit</h1>
        <h2>Threat Intelligence</h2>
    </div>"""

# =============================================================================
# HTML EMAIL TEMPLATE BUILDER (MATCHING EXACT CSU ADVISORY & DUAL THREAT METER)
# =============================================================================
def build_advisory_html(data):
    sev_color, sev_label, threat_score, meter_pos = get_severity_metrics(data["severity"])
    header_banner_html = get_header_banner_html()
    
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
    <style>
        body {{
            font-family: Arial, Calibri, sans-serif;
            font-size: 12px;
            color: #000000;
            background-color: #e9ecef;
            margin: 0;
            padding: 15px;
        }}
        .container {{
            max-width: 860px;
            margin: 0 auto;
            background: #ffffff;
            border: 2px solid #1F4E79;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }}
        .header-banner {{
            background: linear-gradient(135deg, #001833 0%, #003DA5 55%, #001833 100%);
            color: #ffffff;
            padding: 18px 24px;
        }}
        .header-banner h1 {{
            margin: 0;
            font-size: 24px;
            font-weight: bold;
            letter-spacing: 0.5px;
        }}
        .header-banner h2 {{
            margin: 3px 0 0 0;
            font-size: 17px;
            font-weight: normal;
            color: #7EC8E3;
        }}
        .section-bar {{
            background-color: #EBF1F5;
            text-align: center;
            padding: 8px;
            border-top: 1px solid #1F4E79;
            border-bottom: 1px solid #1F4E79;
        }}
        .section-bar h3 {{
            margin: 0;
            font-size: 15px;
            color: #1F4E79;
            font-weight: bold;
        }}
        .section-bar p {{
            margin: 3px 0 0 0;
            font-size: 12px;
            color: #222222;
            font-weight: bold;
        }}
        table.advisory-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        table.advisory-table td {{
            padding: 8px 12px;
            border: 1px solid #000000;
            vertical-align: top;
            font-size: 12px;
            line-height: 1.45;
        }}
        td.label-col {{
            width: 22%;
            background-color: #EBF1F5;
            color: #1F4E79;
            font-weight: bold;
        }}
        td.data-col {{
            width: 78%;
            background-color: #FFFFFF;
            color: #000000;
        }}
        .sev-badge {{
            font-weight: bold;
            color: {sev_color};
            font-size: 13px;
        }}
        .yellow-note {{
            background-color: #FFFF00;
            border: 1px solid #cccc00;
            color: #000000;
            font-weight: bold;
            font-size: 11px;
            padding: 4px 8px;
            margin-top: 8px;
            display: inline-block;
        }}
        .footer-banner {{
            background-color: #002855;
            color: #ffffff;
            text-align: center;
            padding: 10px 15px;
            font-size: 10.5px;
            line-height: 1.5;
            border-top: 1px solid #1F4E79;
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Banner Image -->
        {header_banner_html}

        <!-- Title Header Bar -->
        <div class="section-bar">
            <h3>Information Security Advisory - Alert</h3>
            <p>Date &amp; Time Issued: {data['date']}</p>
        </div>

        <!-- Main Metadata Table -->
        <table class="advisory-table">
            <tr>
                <td class="label-col">Advisory Number</td>
                <td class="data-col"><b>{data['advisory_no']}</b></td>
            </tr>
            <tr>
                <td class="label-col">Title</td>
                <td class="data-col"><b>{data['title']}</b></td>
            </tr>
            <tr>
                <td class="label-col">Impacted Elements</td>
                <td class="data-col">{impacted_html or "Refer to technical analysis."}</td>
            </tr>
            <tr>
                <td class="label-col">Summary</td>
                <td class="data-col">{summary_html}</td>
            </tr>
            <tr>
                <td class="label-col">Severity</td>
                <td class="data-col">
                    <span class="sev-badge">{data['severity']}</span>
                    <span style="display: inline-block; width: 32px; height: 7px; background-color: {sev_color}; margin-left: 8px; vertical-align: middle; border-radius: 2px;"></span>
                </td>
            </tr>
            <tr>
                <td class="label-col">Impact Type</td>
                <td class="data-col">{data['attack_type'] or "Arbitrary code execution"}</td>
            </tr>
            <tr>
                <td class="label-col">Impact Analysis</td>
                <td class="data-col">{impact_html}</td>
            </tr>
            <tr>
                <td class="label-col">Vendor Solution</td>
                <td class="data-col">{solution_html or "Apply vendor patches immediately."}</td>
            </tr>
        </table>

        <!-- Threat Meter Header -->
        <div class="section-bar">
            <h3>Threat Meter</h3>
        </div>

        <!-- DUAL THREAT METERS (VENDOR METER + GSOC METER) -->
        <div style="padding: 16px 10px 10px 10px; background-color: #ffffff; border-left: 1px solid #000; border-right: 1px solid #000;">
            <table style="width: 100%; border: none; border-collapse: collapse;">
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
                        <div style="font-family: Arial; font-weight: bold; font-size: 13px; color: #000; margin-top: 4px;">Vendor Meter</div>
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
                        <div style="font-family: Arial; font-weight: bold; font-size: 13px; color: #000; margin-top: 4px;">GSOC Meter</div>
                    </td>
                </tr>
            </table>
        </div>

        <!-- GSOC Assessment Section Header -->
        <div class="section-bar">
            <h3>GSOC Assessment</h3>
        </div>

        <!-- GSOC Assessment Details Table -->
        <table class="advisory-table">
            <tr>
                <td class="label-col">Exploitation Probability</td>
                <td class="data-col"><b>{data['severity']}</b></td>
            </tr>
            <tr>
                <td class="label-col">GSOC Risk Assessment</td>
                <td class="data-col">
                    Successful exploitation of these vulnerabilities could lead to {data['attack_type'] or "system compromise"}. Hence it has been categorized as <b>{data['severity']}</b>. No active exploitation detected so far.
                </td>
            </tr>
            <tr>
                <td class="label-col">GSOC Recommendation</td>
                <td class="data-col">
                    <div>Apply the latest patch released by the vendor.</div>
                    <div style="margin-top: 6px;">{solution_html}</div>
                    <div class="yellow-note">/* Test changes on non-production systems before applying on production systems */</div>
                </td>
            </tr>
            <tr>
                <td class="label-col">References</td>
                <td class="data-col">{ref_html}</td>
            </tr>
            <tr>
                <td class="label-col">References CVE's</td>
                <td class="data-col"><b>{data['cve'] or "N/A"}</b></td>
            </tr>
        </table>

        <!-- Capgemini CSU Footer -->
        <div class="footer-banner">
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
    html_content = build_advisory_html(data)
    
    # 1. If Classic Outlook COM is available, save via Outlook COM API
    if outlook is not None:
        try:
            mail_item = outlook.CreateItem(0)
            mail_item.Subject = f"CSU Threat Intelligence Notification Advisory {data['advisory_no']} - {data['title']} ({data['severity'].upper()})"
            mail_item.HTMLBody = html_content
            abs_path = os.path.abspath(output_path)
            mail_item.SaveAs(abs_path, 3) # 3 = olMSG
            mail_item.Close(1)
            return abs_path
        except Exception:
            pass
    
    # 2. Universal Outlook .msg package creation
    msg_path = output_path if output_path.endswith(".msg") else output_path + ".msg"
    adv_email = EmailMessage()
    adv_email["Subject"] = f"CSU Threat Intelligence Notification Advisory {data['advisory_no']} - {data['title']} ({data['severity'].upper()})"
    adv_email["From"] = "CSU Threat Intelligence <csu-alerts@capgemini.com>"
    adv_email["To"] = EMAIL_TO
    adv_email.set_content(f"CSU Threat Intelligence Advisory {data['advisory_no']}: {data['title']}")
    adv_email.add_alternative(html_content, subtype="html")
    
    abs_path = os.path.abspath(msg_path)
    with open(abs_path, "wb") as f:
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
# CYCLE EXECUTION FLOW (HANDLES 25K+ ROWS EFFICIENTLY)
# =============================================================================
def run_cycle():
    now_str = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
    print("\n" + "=" * 60)
    print(f"  CSU Alert Automation Cycle  [{now_str}]")
    print("=" * 60)
    
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

    generated_files = []
    total_scanned = 0
    
    # Fast row iterator
    for row in sheet.iter_rows(min_row=2, values_only=True):
        total_scanned += 1
        if not row or len(row) < 2:
            continue
            
        advisory_type = safe_str(row[COL_TYPE - 1]).lower()
        
        # TARGET ONLY ALERTS
        if advisory_type != "alert":
            continue
            
        row_hash = get_row_hash(row)
        if row_hash in processed_ids:
            continue # Already processed in previous cycle
            
        title = safe_str(row[COL_TITLE - 1])
        if not title:
            continue
            
        advisory_no = safe_str(row[COL_ADVISORY_NO - 1]) or "TI-Alert"
        print(f"\n[>] Found New Alert: {advisory_no} - {title}")
        
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
        }
        
        clean_title = "".join(c for c in title if c.isalnum() or c in " _-")[:40].strip()
        msg_filename = f"{advisory_no}_Alert_{clean_title}.msg"
        msg_path = os.path.join(MSG_OUTPUT_DIR, msg_filename)
        
        try:
            saved_file = create_individual_advisory(outlook, data, msg_path)
            print(f"    -> Generated Advisory: {saved_file}")
            
            generated_files.append(saved_file)
            mark_as_processed(row_hash)
            processed_ids.add(row_hash)
            
        except Exception as e:
            print(f"    [!] Error generating advisory for '{title}': {e}")

    wb.close()
    print(f"[*] Scanned {total_scanned} total rows.")

    if generated_files:
        print(f"\n[*] Triggering notification email for {len(generated_files)} new advisory package(s)...")
        send_master_trigger_email(outlook, generated_files)
        first_file = os.path.abspath(generated_files[0])
        print(f"[*] Opening preview in your web browser: {first_file}")
        webbrowser.open(first_file)
    else:
        print("\n[*] No new unprocessed 'Alert' records found in Excel.")

    return len(generated_files)

def main():
    if SCHEDULE_INTERVAL_HOURS > 0:
        print(f"[*] Starting continuous automation scheduler (running every {SCHEDULE_INTERVAL_HOURS} hour(s))...")
        print("[*] Press Ctrl+C at any time to stop.\n")
        try:
            while True:
                run_cycle()
                next_check = datetime.fromtimestamp(time.time() + SCHEDULE_INTERVAL_HOURS * 3600).strftime("%H:%M:%S")
                print(f"\n[*] Waiting for next check at {next_check} (sleeping {SCHEDULE_INTERVAL_HOURS}h)...")
                time.sleep(SCHEDULE_INTERVAL_HOURS * 3600)
        except KeyboardInterrupt:
            print("\n[*] Automation scheduler stopped by user.")
    else:
        run_cycle()
        print("\n" + "=" * 60)
        print("  Automation cycle completed.")
        print("=" * 60)

if __name__ == "__main__":
    main()
