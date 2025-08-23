# fixed_smtp_server.py - Enhanced Email Server (SMTP + Gmail + Outlook + Inline Images)
from flask import Flask, request, jsonify
from flask_cors import CORS
import smtplib
import os
import base64
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Try to import win32com, but don't fail if not available
try:
    import win32com.client as win32
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    print("⚠️ win32com not available - Outlook disabled, SMTP/Gmail only")

app = Flask(__name__)
CORS(app)

# SMTP server configurations
SMTP_CONFIGS = {
    'hotmail.com': {'server': 'smtp-mail.outlook.com', 'port': 587, 'tls': True, 'name': 'Hotmail/Outlook'},
    'outlook.com': {'server': 'smtp-mail.outlook.com', 'port': 587, 'tls': True, 'name': 'Hotmail/Outlook'},
    'gmail.com': {'server': 'smtp.gmail.com', 'port': 587, 'tls': True, 'name': 'Gmail'},
    'yahoo.com': {'server': 'smtp.mail.yahoo.com', 'port': 587, 'tls': True, 'name': 'Yahoo'},
    'office365.com': {'server': 'smtp.office365.com', 'port': 587, 'tls': True, 'name': 'Office 365'}
}


# -------------------------------------------------------------------
# Utility: embed inline image (base64)
# -------------------------------------------------------------------
def embed_inline_image(msg, html_body, image_data, cid="embedded_image"):
    """Embed base64 image into email, replace placeholder with cid."""
    if not image_data:
        return html_body

    try:
        if "," in image_data:  # strip header if present
            _, encoded = image_data.split(",", 1)
        else:
            encoded = image_data
        image_bytes = base64.b64decode(encoded)

        image = MIMEImage(image_bytes)
        image.add_header("Content-ID", f"<{cid}>")
        image.add_header("Content-Disposition", "inline")
        msg.attach(image)

        # Replace placeholder
        return html_body.replace("{{IMAGE_PLACEHOLDER}}", f"cid:{cid}")
    except Exception as e:
        print(f"⚠️ Image embedding failed: {e}")
        return html_body.replace(
            "{{IMAGE_PLACEHOLDER}}",
            "https://via.placeholder.com/400x200?text=Image+Error",
        )


# -------------------------------------------------------------------
# Health Check
# -------------------------------------------------------------------
@app.route('/health', methods=['GET'])
def health_check():
    health_info = {
        'status': 'ok',
        'outlook': 'available' if OUTLOOK_AVAILABLE else 'unavailable',
        'smtp': True
    }

    if OUTLOOK_AVAILABLE:
        try:
            outlook = win32.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            accounts = namespace.Accounts
            health_info['outlook_accounts'] = accounts.Count
        except Exception as e:
            health_info['outlook_error'] = str(e)

    return jsonify(health_info), 200


# -------------------------------------------------------------------
# Send via Generic SMTP
# -------------------------------------------------------------------
@app.route('/send-smtp-email', methods=['POST'])
def send_smtp_email():
    try:
        data = request.json
        receiver = data.get('to', '').strip()
        subject = data.get('subject', '').strip()
        html_body = data.get('body', '').strip()
        smtp_email = data.get('smtpEmail', '').strip()
        smtp_password = data.get('smtpPassword', '').strip()
        image_data = data.get('imageData')

        if not all([receiver, subject, html_body, smtp_email, smtp_password]):
            return jsonify({'success': False, 'error': 'Missing required fields'}), 400

        # Determine SMTP config
        domain = smtp_email.lower().split('@')[-1]
        smtp_config = None
        for d, config in SMTP_CONFIGS.items():
            if d in domain:
                smtp_config = config
                break
        if not smtp_config:
            smtp_config = SMTP_CONFIGS['office365.com']

        # Create email
        msg = MIMEMultipart('related')
        msg['From'] = smtp_email
        msg['To'] = receiver
        msg['Subject'] = subject

        # Inline image
        html_body = embed_inline_image(msg, html_body, image_data)

        # Attach HTML
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'], timeout=30)
        if smtp_config['tls']:
            server.starttls()
        server.login(smtp_email, smtp_password)
        server.sendmail(smtp_email, [receiver], msg.as_string())
        server.quit()

        return jsonify({'success': True, 'message': f'Email sent via {smtp_config["name"]}'}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# -------------------------------------------------------------------
# Send via Gmail
# -------------------------------------------------------------------
@app.route('/send-gmail-email', methods=['POST'])
def send_gmail_email():
    try:
        data = request.json
        sender_email = data.get('gmailEmail', '').strip()
        sender_password = data.get('gmailPassword', '').strip()
        receiver = data.get('to', '').strip()
        subject = data.get('subject', '').strip()
        html_body = data.get('body', '').strip()
        image_data = data.get('imageData')

        if not all([sender_email, sender_password, receiver, subject, html_body]):
            return jsonify({'success': False, 'error': 'Missing required fields'}), 400

        if '@gmail.com' not in sender_email.lower():
            return jsonify({'success': False, 'error': 'Sender must be a Gmail address'}), 400

        msg = MIMEMultipart('related')
        msg['From'] = sender_email
        msg['To'] = receiver
        msg['Subject'] = subject

        # Inline image
        html_body = embed_inline_image(msg, html_body, image_data, cid="gmail_img")

        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [receiver], msg.as_string())
        server.quit()

        return jsonify({'success': True, 'message': 'Email sent via Gmail'}), 200
    except smtplib.SMTPAuthenticationError:
        return jsonify({'success': False, 'error': 'Authentication failed. Use Gmail App Password instead of regular password.'}), 401
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# -------------------------------------------------------------------
# Send via Outlook (COM API)
# -------------------------------------------------------------------
@app.route('/send-outlook-email', methods=['POST'])
def send_outlook_email():
    if not OUTLOOK_AVAILABLE:
        return jsonify({'success': False, 'error': 'Outlook COM not available. Install pywin32 on Windows.'}), 400

    try:
        data = request.json
        receiver = data.get('to', '').strip()
        subject = data.get('subject', '').strip()
        html_body = data.get('body', '').strip()
        image_data = data.get('imageData')

        if not all([receiver, subject, html_body]):
            return jsonify({'success': False, 'error': 'Missing required fields'}), 400

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = receiver
        mail.Subject = subject

        # Handle inline image
        if image_data:
            try:
                if "," in image_data:
                    _, encoded = image_data.split(",", 1)
                else:
                    encoded = image_data
                image_bytes = base64.b64decode(encoded)

                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(image_bytes)
                    image_path = tmp.name

                attachment = mail.Attachments.Add(image_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "OutlookImg1"
                )
                html_body = html_body.replace("{{IMAGE_PLACEHOLDER}}", "cid:OutlookImg1")

                # Clean later
                try:
                    os.unlink(image_path)
                except:
                    pass
            except Exception as e:
                print(f"⚠️ Outlook image error: {e}")

        mail.HTMLBody = html_body
        mail.Send()

        return jsonify({'success': True, 'message': 'Email sent via Outlook'}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# -------------------------------------------------------------------
# Main Entry Point
# -------------------------------------------------------------------
if __name__ == "__main__":
    print("🚀 Starting Enhanced Email Server")
    print("🌐 Running on http://localhost:5000")
    if OUTLOOK_AVAILABLE:
        print("✅ Outlook COM available")
    else:
        print("⚠️ Outlook COM not available (install pywin32 if needed)")
    app.run(host="localhost", port=5000, debug=True)
