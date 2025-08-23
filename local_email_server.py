# local_email_server.py - Run this on user's Windows machine
from flask import Flask, request, jsonify
from flask_cors import CORS
import win32com.client as win32
import os
import base64
import tempfile
import json

app = Flask(__name__)
CORS(app)  # Allow requests from web browser

@app.route('/health', methods=['GET'])
def health_check():
    """Enhanced health check endpoint with detailed diagnostics"""
    health_info = {
        'server_status': 'running',
        'outlook_available': False,
        'outlook_running': False,
        'outlook_configured': False,
        'error': None,
        'diagnostics': [],
        'accounts': []  # Add accounts to health check
    }
    
    try:
        # Step 1: Test if we can create Outlook application object
        try:
            outlook = win32.Dispatch("Outlook.Application")
            health_info['outlook_available'] = True
            health_info['diagnostics'].append("✅ Outlook COM object created successfully")
        except Exception as com_error:
            health_info['diagnostics'].append(f"❌ Failed to create Outlook COM object: {str(com_error)}")
            health_info['error'] = f"Outlook COM Error: {str(com_error)}"
            health_info['status'] = 'error'
            return jsonify(health_info), 200
        
        # Step 2: Test if Outlook is actually running
        try:
            # Try to get the namespace (this requires Outlook to be running)
            namespace = outlook.GetNamespace("MAPI")
            health_info['outlook_running'] = True
            health_info['diagnostics'].append("✅ Outlook application is running")
        except Exception as running_error:
            health_info['diagnostics'].append(f"⚠️ Outlook may not be running: {str(running_error)}")
            # Don't fail here - Outlook can auto-start
        
        # Step 3: Test if we can access email accounts and collect them
        try:
            namespace = outlook.GetNamespace("MAPI")
            accounts = namespace.Accounts
            if accounts.Count > 0:
                health_info['outlook_configured'] = True
                health_info['diagnostics'].append(f"✅ Found {accounts.Count} configured email account(s)")
                
                # Collect account information
                account_list = []
                for i in range(1, accounts.Count + 1):
                    try:
                        account = accounts.Item(i)
                        account_info = {
                            'index': i,
                            'displayName': account.DisplayName,
                            'smtpAddress': account.SmtpAddress,
                            'accountType': getattr(account, 'AccountType', 'Unknown'),
                            'isDefault': False
                        }
                        account_list.append(account_info)
                    except Exception as acc_error:
                        health_info['diagnostics'].append(f"⚠️ Could not read account {i}: {str(acc_error)}")
                
                health_info['accounts'] = account_list
                
                # Try to determine default account
                try:
                    default_store = namespace.DefaultStore
                    for account_info in account_list:
                        if account_info['displayName'] == default_store.DisplayName:
                            account_info['isDefault'] = True
                            break
                except:
                    # If first method fails, mark first account as default
                    if account_list:
                        account_list[0]['isDefault'] = True
                
            else:
                health_info['diagnostics'].append("❌ No email accounts configured in Outlook")
                health_info['error'] = "No email accounts found in Outlook"
        except Exception as config_error:
            health_info['diagnostics'].append(f"❌ Failed to check Outlook configuration: {str(config_error)}")
            health_info['error'] = f"Outlook configuration error: {str(config_error)}"
        
        # Step 4: Test creating a mail item (without sending)
        try:
            test_mail = outlook.CreateItem(0)  # 0 = olMailItem
            test_mail = None  # Clean up
            health_info['diagnostics'].append("✅ Can create mail items")
        except Exception as mail_error:
            health_info['diagnostics'].append(f"❌ Cannot create mail items: {str(mail_error)}")
            health_info['error'] = f"Mail creation error: {str(mail_error)}"
        
        # Final status determination
        if health_info['outlook_available'] and health_info['outlook_configured']:
            health_info['status'] = 'healthy'
            health_info['message'] = 'Outlook is ready for sending emails'
            return jsonify(health_info), 200
        elif health_info['outlook_available']:
            health_info['status'] = 'partial'
            health_info['message'] = 'Outlook available but may need configuration'
            return jsonify(health_info), 200
        else:
            health_info['status'] = 'error'
            health_info['message'] = 'Outlook is not available'
            return jsonify(health_info), 200
            
    except Exception as e:
        health_info['status'] = 'error'
        health_info['error'] = str(e)
        health_info['message'] = f'Health check failed: {str(e)}'
        health_info['diagnostics'].append(f"💥 Unexpected error: {str(e)}")
        return jsonify(health_info), 200

@app.route('/accounts', methods=['GET'])
def get_accounts():
    """Get all available Outlook accounts"""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        accounts = session.Accounts
        
        account_list = []
        default_account = None
        
        # Get all accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            account_info = {
                'index': i,
                'displayName': account.DisplayName,
                'smtpAddress': account.SmtpAddress,
                'accountType': getattr(account, 'AccountType', 'Unknown'),
                'isDefault': False
            }
            account_list.append(account_info)
        
        # Try to determine default account
        try:
            # Method 1: Check default store
            default_store = session.DefaultStore
            for account_info in account_list:
                if account_info['displayName'] == default_store.DisplayName:
                    account_info['isDefault'] = True
                    default_account = account_info
                    break
        except:
            pass
        
        # Method 2: Check default sending account via test mail
        if not default_account:
            try:
                test_mail = outlook.CreateItem(0)
                sending_account = test_mail.SendUsingAccount
                if sending_account:
                    for account_info in account_list:
                        if account_info['smtpAddress'].lower() == sending_account.SmtpAddress.lower():
                            account_info['isDefault'] = True
                            default_account = account_info
                            break
                test_mail = None  # Clean up
            except:
                pass
        
        # If still no default found, mark first as default
        if not default_account and account_list:
            account_list[0]['isDefault'] = True
            default_account = account_list[0]
        
        return jsonify({
            'success': True,
            'accounts': account_list,
            'defaultAccount': default_account,
            'totalAccounts': len(account_list)
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Failed to get accounts: {str(e)}',
            'accounts': [],
            'totalAccounts': 0
        }), 500

@app.route('/send-outlook-email', methods=['POST'])
def send_outlook_email():
    """API endpoint to send email via Outlook with account selection"""
    try:
        data = request.json
        
        # Extract email data
        receiver = data.get('to')
        cc = data.get('cc', '')
        bcc = data.get('bcc', '')
        subject = data.get('subject')
        html_body = data.get('body')
        image_data = data.get('imageData')
        preferred_account = data.get('fromAccount')  # New: specify sending account
        
        # Validation
        if not all([receiver, subject, html_body]):
            return jsonify({'success': False, 'error': 'Missing required fields'}), 400
        
        # Create Outlook application object
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        # Try to set specific sending account if specified
        if preferred_account:
            try:
                session = outlook.Session
                accounts = session.Accounts
                target_account = None
                
                # Find the specified account
                for i in range(1, accounts.Count + 1):
                    account = accounts.Item(i)
                    if (account.SmtpAddress.lower() == preferred_account.lower() or 
                        account.DisplayName.lower() == preferred_account.lower()):
                        target_account = account
                        break
                
                if target_account:
                    # Set the sending account
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, target_account))
                    print(f"Using account: {target_account.DisplayName} ({target_account.SmtpAddress})")
                else:
                    print(f"Warning: Account '{preferred_account}' not found, using default")
                    
            except Exception as account_error:
                print(f"Warning: Could not set specific account: {account_error}")
                # Continue with default account
        
        # Set email properties
        mail.To = receiver
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        
        # Handle inline image if provided
        image_path = None
        if image_data:
            try:
                # Decode base64 image
                header, encoded = image_data.split(',', 1)
                image_bytes = base64.b64decode(encoded)
                
                # Determine file extension from mime type
                mime_type = header.split(':')[1].split(';')[0]
                extension = mime_type.split('/')[1]
                
                # Save temporary image file
                with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{extension}') as temp_file:
                    temp_file.write(image_bytes)
                    image_path = temp_file.name
                
                # Add inline image attachment
                attachment = mail.Attachments.Add(image_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
                    "MyId1"
                )
                
                # Replace placeholder with cid reference
                html_body = html_body.replace('{{IMAGE_PLACEHOLDER}}', 'cid:MyId1')
                
            except Exception as img_error:
                return jsonify({'success': False, 'error': f'Image processing error: {str(img_error)}'}), 400
        
        # Set HTML body
        mail.HTMLBody = html_body
        
        # Send email
        mail.Send()
        
        # Clean up temporary file
        if image_path and os.path.exists(image_path):
            try:
                os.unlink(image_path)
            except:
                pass  # File cleanup failure is not critical
        
        return jsonify({
            'success': True, 
            'message': 'Email sent successfully via Outlook!',
            'account_used': preferred_account or 'default'
        })
        
    except Exception as e:
        return jsonify({
            'success': False, 
            'error': f'Failed to send email: {str(e)}'
        }), 500

if __name__ == '__main__':
    print("🚀 Starting Local Email Server...")
    print("📧 Outlook integration ready!")
    print("🌐 Server running on http://localhost:5000")
    print("=" * 50)
    
    # Pre-flight check
    print("🔍 Performing pre-flight checks...")
    try:
        outlook = win32.Dispatch("Outlook.Application")
        print("✅ Outlook COM object created successfully")
        
        try:
            namespace = outlook.GetNamespace("MAPI")
            accounts = namespace.Accounts
            print(f"✅ Found {accounts.Count} email account(s) configured")
            
            if accounts.Count > 0:
                print("📧 Available accounts:")
                for i in range(1, min(accounts.Count + 1, 4)):  # Show first 3 accounts
                    try:
                        account = accounts.Item(i)
                        print(f"   {i}. {account.DisplayName} ({account.SmtpAddress})")
                    except:
                        print(f"   {i}. Account {i} (could not get details)")
                
                if accounts.Count > 3:
                    print(f"   ... and {accounts.Count - 3} more accounts")
            else:
                print("⚠️ WARNING: No email accounts configured in Outlook!")
                print("   Please set up at least one email account in Outlook")
                
        except Exception as ns_error:
            print(f"⚠️ WARNING: Could not access Outlook accounts: {str(ns_error)}")
            print("   Try opening Outlook manually first")
            
    except Exception as e:
        print(f"❌ ERROR: Cannot connect to Outlook: {str(e)}")
        print("\n🔧 Troubleshooting steps:")
        print("1. Make sure Microsoft Outlook is installed")
        print("2. Open Outlook manually and complete setup")
        print("3. Close and restart this server")
        print("4. Check Windows permissions for COM objects")
    
    print("=" * 50)
    print("⚠️  Make sure Microsoft Outlook is installed and configured")
    print("🔗 Web interface will be available after starting the server")
    
    # Run server
    app.run(host='localhost', port=5000, debug=True)

# requirements.txt
"""
Flask==2.3.3
Flask-CORS==4.0.0
pywin32==306
"""

# To install dependencies:
# pip install Flask Flask-CORS pywin32