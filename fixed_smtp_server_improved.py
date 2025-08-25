# fixed_smtp_server.py - Enhanced Email Server (SMTP + Gmail + Outlook + Inline Images)
from flask import Flask, request, jsonify
from flask_cors import CORS
import smtplib
import os
import base64
import tempfile
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import make_msgid

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
# Utility: Create proper MIME structure with inline images
# -------------------------------------------------------------------
def create_email_with_inline_images(sender, receiver, subject, html_body, image_data=None):
    """Create properly structured MIME email with inline images."""
    
    # Create root message
    msg = MIMEMultipart('mixed')
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject
    
    # Create alternative container (for text/html alternatives)
    msg_alternative = MIMEMultipart('alternative')
    
    # Create plain text version (strip HTML tags)
    plain_text = re.sub('<[^<]+?>', '', html_body)
    plain_text = re.sub(r'\s+', ' ', plain_text).strip()
    if '{{IMAGE_PLACEHOLDER}}' in plain_text:
        plain_text = plain_text.replace('{{IMAGE_PLACEHOLDER}}', '[Image]')
    
    text_part = MIMEText(plain_text, 'plain', 'utf-8')
    msg_alternative.attach(text_part)
    
    # Handle inline images
    if image_data and '{{IMAGE_PLACEHOLDER}}' in html_body:
        # Create related container for HTML + images
        msg_related = MIMEMultipart('related')
        
        # Generate unique Content-ID
        img_cid = make_msgid(domain='emailserver.local')[1:-1]  # Remove < >
        
        try:
            # Process image data
            if "," in image_data:
                header, encoded = image_data.split(",", 1)
                # Extract image type from header (e.g., data:image/png;base64,)
                if 'image/' in header:
                    img_type = header.split('image/')[1].split(';')[0].lower()
                else:
                    img_type = 'png'  # default
            else:
                encoded = image_data
                img_type = 'png'  # default
            
            image_bytes = base64.b64decode(encoded)
            
            # Create image attachment
            if img_type.lower() in ['jpg', 'jpeg']:
                img_mime = MIMEImage(image_bytes, 'jpeg')
            elif img_type.lower() == 'gif':
                img_mime = MIMEImage(image_bytes, 'gif')
            else:
                img_mime = MIMEImage(image_bytes, 'png')
            
            # Set proper headers
            img_mime.add_header('Content-ID', f'<{img_cid}>')
            img_mime.add_header('Content-Disposition', 'inline', filename=f'image.{img_type}')
            
            # Replace placeholder in HTML
            html_body = html_body.replace('{{IMAGE_PLACEHOLDER}}', f'cid:{img_cid}')
            
            # Attach HTML first, then image
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg_related.attach(html_part)
            msg_related.attach(img_mime)
            
            # Attach related to alternative
            msg_alternative.attach(msg_related)
            
        except Exception as e:
            print(f"⚠️ Image processing failed: {e}")
            # Fallback: just attach HTML without image
            html_body = html_body.replace('{{IMAGE_PLACEHOLDER}}', '<p><em>[Image could not be loaded]</em></p>')
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg_alternative.attach(html_part)
    else:
        # No image, just attach HTML
        html_body = html_body.replace('{{IMAGE_PLACEHOLDER}}', '')
        html_part = MIMEText(html_body, 'html', 'utf-8')
        msg_alternative.attach(html_part)
    
    # Attach alternative to root message
    msg.attach(msg_alternative)
    
    return msg


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

        # Create properly structured email
        msg = create_email_with_inline_images(smtp_email, receiver, subject, html_body, image_data)

        # Send email
        with smtplib.SMTP(smtp_config['server'], smtp_config['port'], timeout=30) as server:
            if smtp_config['tls']:
                server.starttls()
            server.login(smtp_email, smtp_password)
            server.sendmail(smtp_email, [receiver], msg.as_string())

        return jsonify({'success': True, 'message': f'Email sent via {smtp_config["name"]}'}), 200
    except smtplib.SMTPAuthenticationError as e:
        return jsonify({'success': False, 'error': 'SMTP Authentication failed. Check credentials or use App Password.'}), 401
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

        # Create properly structured email
        msg = create_email_with_inline_images(sender_email, receiver, subject, html_body, image_data)

        # Send via Gmail SMTP
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, [receiver], msg.as_string())

        return jsonify({'success': True, 'message': 'Email sent via Gmail'}), 200
    except smtplib.SMTPAuthenticationError:
        return jsonify({'success': False, 'error': 'Gmail Authentication failed. Use App Password instead of regular password.'}), 401
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
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = receiver
        mail.Subject = subject

        # Handle inline image for Outlook
        if image_data and '{{IMAGE_PLACEHOLDER}}' in html_body:
            try:
                # Process base64 image
                if "," in image_data:
                    _, encoded = image_data.split(",", 1)
                else:
                    encoded = image_data
                image_bytes = base64.b64decode(encoded)

                # Create temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(image_bytes)
                    image_path = tmp.name

                # Add as inline attachment
                attachment = mail.Attachments.Add(image_path)
                # Set Content-ID for inline display
                cid_property = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                attachment.PropertyAccessor.SetProperty(cid_property, "OutlookImg1")
                
                # Replace placeholder
                html_body = html_body.replace("{{IMAGE_PLACEHOLDER}}", "cid:OutlookImg1")

                # Clean up temp file
                try:
                    os.unlink(image_path)
                except:
                    pass
            except Exception as e:
                print(f"⚠️ Outlook image error: {e}")
                html_body = html_body.replace("{{IMAGE_PLACEHOLDER}}", "<p><em>[Image could not be loaded]</em></p>")
        else:
            html_body = html_body.replace("{{IMAGE_PLACEHOLDER}}", "")

        mail.HTMLBody = html_body
        mail.Send()

        return jsonify({'success': True, 'message': 'Email sent via Outlook'}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': f'Outlook error: {str(e)}'}), 500


# -------------------------------------------------------------------
# Test endpoint to validate email structure
# -------------------------------------------------------------------
@app.route('/test-email-structure', methods=['POST'])
def test_email_structure():
    """Test endpoint to see the generated email structure without sending."""
    try:
        data = request.json
        sender = data.get('sender', 'test@example.com')
        receiver = data.get('to', 'recipient@example.com')
        subject = data.get('subject', 'Test Email')
        html_body = data.get('body', '<p>Test body with {{IMAGE_PLACEHOLDER}}</p>')
        image_data = data.get('imageData')

        msg = create_email_with_inline_images(sender, receiver, subject, html_body, image_data)
        
        # Return email structure info
        structure_info = {
            'headers': dict(msg.items()),
            'content_type': msg.get_content_type(),
            'parts': [],
            'raw_preview': msg.as_string()[:1000] + "..." if len(msg.as_string()) > 1000 else msg.as_string()
        }
        
        def analyze_part(part, level=0):
            part_info = {
                'level': level,
                'content_type': part.get_content_type(),
                'headers': dict(part.items()) if hasattr(part, 'items') else {}
            }
            if hasattr(part, 'get_payload') and hasattr(part, 'is_multipart'):
                if part.is_multipart():
                    part_info['subparts'] = []
                    for subpart in part.get_payload():
                        part_info['subparts'].append(analyze_part(subpart, level + 1))
                else:
                    payload = part.get_payload()
                    if isinstance(payload, str) and len(payload) > 100:
                        part_info['payload_preview'] = payload[:100] + "..."
                    else:
                        part_info['payload_preview'] = str(payload)[:100]
            return part_info
        
        structure_info['structure'] = analyze_part(msg)
        
        return jsonify({'success': True, 'structure': structure_info}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# -------------------------------------------------------------------
# Main Entry Point
# -------------------------------------------------------------------
if __name__ == "__main__":
    print("🚀 Starting Enhanced Email Server with Fixed MIME Structure")
    print("🌐 Running on http://localhost:5000")
    print("🔧 New endpoint: POST /test-email-structure (for debugging)")
    if OUTLOOK_AVAILABLE:
        print("✅ Outlook COM available")
    else:
        print("⚠️ Outlook COM not available (install pywin32 if needed)")
    app.run(host="localhost", port=5000, debug=True)