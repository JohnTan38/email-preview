# outlook_diagnostics.py - Run this to diagnose Outlook issues
import sys
import os

def check_imports():
    """Check if all required libraries are available"""
    print("🔍 Checking Python dependencies...")
    
    try:
        import win32com.client
        print("✅ win32com.client imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import win32com.client: {e}")
        print("   Fix: pip install pywin32")
        return False
    
    try:
        import flask
        print(f"✅ Flask {flask.__version__} imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import Flask: {e}")
        print("   Fix: pip install Flask")
        return False
    
    try:
        import flask_cors
        print("✅ Flask-CORS imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import Flask-CORS: {e}")
        print("   Fix: pip install Flask-CORS")
        return False
    
    return True

def check_outlook_installation():
    """Check if Outlook is properly installed"""
    print("\n🔍 Checking Outlook installation...")
    
    try:
        import win32com.client
        
        # Try to create Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✅ Outlook COM object created successfully")
        
        # Check Outlook version
        try:
            version = outlook.Version
            print(f"📧 Outlook version: {version}")
        except Exception as e:
            print(f"⚠️ Could not get Outlook version: {e}")
        
        return True
        
    except Exception as e:
        print(f"❌ Failed to create Outlook COM object: {e}")
        print("\n🔧 Possible causes:")
        print("   - Microsoft Outlook is not installed")
        print("   - Outlook is not properly registered in Windows")
        print("   - COM security settings are blocking access")
        print("   - Outlook needs to be run as Administrator")
        return False

def check_outlook_configuration():
    """Check if Outlook is configured with email accounts"""
    print("\n🔍 Checking Outlook configuration...")
    
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Try to access MAPI namespace
        try:
            namespace = outlook.GetNamespace("MAPI")
            print("✅ MAPI namespace accessible")
        except Exception as e:
            print(f"❌ Cannot access MAPI namespace: {e}")
            print("   This usually means Outlook is not running or not configured")
            print("   Try opening Outlook manually first")
            return False
        
        # Check email accounts
        try:
            accounts = namespace.Accounts
            account_count = accounts.Count
            print(f"📧 Found {account_count} email account(s)")
            
            if account_count == 0:
                print("⚠️ No email accounts configured!")
                print("   Please set up at least one email account in Outlook")
                return False
            
            # List all accounts with details
            personal_accounts = []
            corporate_accounts = []
            
            for i in range(1, account_count + 1):
                try:
                    account = accounts.Item(i)
                    account_type = getattr(account, 'AccountType', 'Unknown')
                    display_name = account.DisplayName
                    smtp_address = account.SmtpAddress
                    
                    print(f"   📧 Account {i}: {display_name}")
                    print(f"      Email: {smtp_address}")
                    print(f"      Type: {account_type}")
                    
                    # Categorize accounts
                    if any(domain in smtp_address.lower() for domain in ['hotmail.com', 'outlook.com', 'gmail.com', 'yahoo.com']):
                        personal_accounts.append((display_name, smtp_address))
                        print(f"      Category: 🏠 Personal Account")
                    else:
                        corporate_accounts.append((display_name, smtp_address))
                        print(f"      Category: 🏢 Corporate Account")
                    
                    print()
                    
                except Exception as e:
                    print(f"   ❌ Account {i}: Could not get details ({e})")
            
            # Check default account
            print("🔍 Checking default account...")
            try:
                default_store = namespace.DefaultStore
                print(f"   Default Store: {default_store.DisplayName}")
                
                # Try to determine which account is default for sending
                test_mail = outlook.CreateItem(0)
                sending_account = test_mail.SendUsingAccount
                if sending_account:
                    print(f"   Default Sending: {sending_account.DisplayName} ({sending_account.SmtpAddress})")
                    
                    # Check if default is personal account
                    if any(domain in sending_account.SmtpAddress.lower() for domain in ['hotmail.com', 'outlook.com', 'gmail.com']):
                        print("   ✅ Default is personal account - Good for automation!")
                    else:
                        print("   ⚠️ Default is corporate account - May have restrictions")
                        if personal_accounts:
                            print(f"   💡 Consider changing default to: {personal_accounts[0][1]}")
                else:
                    print("   ⚠️ No default sending account found")
                    
                test_mail = None  # Clean up
                
            except Exception as default_error:
                print(f"   ❌ Could not determine default account: {default_error}")
            
            # Summary
            print("\n📋 Account Summary:")
            print(f"   🏠 Personal accounts: {len(personal_accounts)}")
            if personal_accounts:
                for name, email in personal_accounts:
                    print(f"      - {email}")
            
            print(f"   🏢 Corporate accounts: {len(corporate_accounts)}")
            if corporate_accounts:
                for name, email in corporate_accounts[:2]:  # Show first 2
                    print(f"      - {email}")
                if len(corporate_accounts) > 2:
                    print(f"      - ... and {len(corporate_accounts) - 2} more")
            
            # Recommendations
            print("\n💡 Recommendations:")
            if personal_accounts:
                print("   ✅ You have personal accounts - these are best for automation")
                print(f"   🎯 Recommended account: {personal_accounts[0][1]}")
                if not any('azx1818@hotmail.com' in email for _, email in personal_accounts):
                    print("   ⚠️ azx1818@hotmail.com not found - please add this account")
            else:
                print("   ⚠️ No personal accounts found")
                print("   📝 Add azx1818@hotmail.com to Outlook for better automation")
            
            return True
            
        except Exception as e:
            print(f"❌ Cannot access email accounts: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Configuration check failed: {e}")
        return False

def test_email_creation():
    """Test if we can create email items"""
    print("\n🔍 Testing email creation...")
    
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Try to create a mail item
        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
        print("✅ Can create mail items")
        
        # Test setting properties
        mail_item.Subject = "Test Subject"
        mail_item.Body = "Test Body"
        print("✅ Can set mail properties")
        
        # Clean up (don't save or send)
        mail_item = None
        return True
        
    except Exception as e:
        print(f"❌ Cannot create mail items: {e}")
        return False

def check_permissions():
    """Check Windows permissions and security"""
    print("\n🔍 Checking permissions and security...")
    
    # Check if running as administrator
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if is_admin:
            print("✅ Running with Administrator privileges")
        else:
            print("⚠️ Not running as Administrator")
            print("   Some COM operations may require elevated privileges")
    except Exception as e:
        print(f"⚠️ Could not check admin status: {e}")
    
    # Check Python architecture
    import platform
    arch = platform.architecture()[0]
    python_version = sys.version
    print(f"🐍 Python: {python_version}")
    print(f"🏗️ Architecture: {arch}")
    
    if arch == '64bit':
        print("✅ Using 64-bit Python (recommended for modern Outlook)")
    else:
        print("⚠️ Using 32-bit Python (may have compatibility issues)")

def main():
    """Run all diagnostic checks"""
    print("=" * 60)
    print("🔬 OUTLOOK INTEGRATION DIAGNOSTIC TOOL")
    print("=" * 60)
    
    all_passed = True
    
    # Run all checks
    all_passed &= check_imports()
    all_passed &= check_outlook_installation()
    all_passed &= check_outlook_configuration()
    all_passed &= test_email_creation()
    check_permissions()  # This is informational
    
    print("\n" + "=" * 60)
    
    if all_passed:
        print("🎉 ALL CHECKS PASSED!")
        print("✅ Outlook integration should work properly")
        print("🚀 Try starting the email server: python local_email_server.py")
    else:
        print("❌ SOME ISSUES FOUND")
        print("🔧 Please fix the issues above and run this diagnostic again")
        print("\n💡 Common solutions:")
        print("   1. Open Microsoft Outlook manually")
        print("   2. Complete Outlook setup wizard if prompted")
        print("   3. Add at least one email account")
        print("   4. Try running Python as Administrator")
        print("   5. Restart Windows after Outlook installation")
    
    print("=" * 60)
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()