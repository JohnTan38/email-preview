# change_default_account.py - Script to help change default Outlook account
import win32com.client as win32

def list_all_accounts():
    """List all available Outlook accounts"""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        accounts = session.Accounts
        
        print("📧 Available Outlook Accounts:")
        print("=" * 50)
        
        account_list = []
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            account_info = {
                'index': i,
                'name': account.DisplayName,
                'email': account.SmtpAddress,
                'type': getattr(account, 'AccountType', 'Unknown')
            }
            account_list.append(account_info)
            
            # Categorize account
            if any(domain in account.SmtpAddress.lower() for domain in ['hotmail.com', 'outlook.com', 'gmail.com']):
                category = "🏠 Personal"
            else:
                category = "🏢 Corporate"
            
            print(f"{i}. {account.DisplayName}")
            print(f"   📧 {account.SmtpAddress}")
            print(f"   🏷️ {category}")
            print()
        
        return account_list
    except Exception as e:
        print(f"❌ Error listing accounts: {e}")
        return []

def get_current_default():
    """Get current default account"""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        
        # Method 1: Check default store
        try:
            default_store = session.DefaultStore
            print(f"🗂️ Default Store: {default_store.DisplayName}")
        except:
            print("⚠️ Could not determine default store")
        
        # Method 2: Check default sending account
        try:
            test_mail = outlook.CreateItem(0)
            sending_account = test_mail.SendUsingAccount
            if sending_account:
                print(f"📤 Default Sending Account: {sending_account.DisplayName} ({sending_account.SmtpAddress})")
                return sending_account.SmtpAddress
            else:
                print("⚠️ No default sending account set")
        except Exception as e:
            print(f"⚠️ Could not check default sending account: {e}")
        
        return None
        
    except Exception as e:
        print(f"❌ Error checking current default: {e}")
        return None

def test_account_sending(account_email):
    """Test if we can send emails with specific account"""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        accounts = session.Accounts
        
        # Find the target account
        target_account = None
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            if account.SmtpAddress.lower() == account_email.lower():
                target_account = account
                break
        
        if not target_account:
            print(f"❌ Account {account_email} not found")
            return False
        
        # Test creating mail with this account
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, target_account))  # Set SendUsingAccount
        mail.Subject = "Test Email"
        mail.Body = "This is a test email to verify account access"
        
        # Don't actually send, just test creation
        print(f"✅ Can create emails with {target_account.DisplayName}")
        mail = None  # Clean up
        return True
        
    except Exception as e:
        print(f"❌ Error testing account {account_email}: {e}")
        return False

def provide_manual_instructions():
    """Provide manual instructions for changing default account"""
    print("\n" + "=" * 60)
    print("📋 MANUAL INSTRUCTIONS TO CHANGE DEFAULT ACCOUNT")
    print("=" * 60)
    
    print("\n🎯 To set azx1818@hotmail.com as default:")
    print("1. Open Microsoft Outlook")
    print("2. File → Account Settings → Account Settings...")
    print("3. On the 'Email' tab, select: azx1818@hotmail.com")
    print("4. Click 'Set as Default' button")
    print("5. Click 'Close'")
    print("6. Restart Outlook completely")
    print("7. Restart the Python email server")
    
    print("\n🔧 If 'Set as Default' is grayed out:")
    print("- This means corporate policy is preventing the change")
    print("- Solution: Use the account selector in the web interface")
    print("- The server will override the default programmatically")
    
    print("\n📧 If azx1818@hotmail.com is missing:")
    print("1. In Outlook: File → Add Account")
    print("2. Enter: azx1818@hotmail.com")
    print("3. Enter your password")
    print("4. Complete the setup")
    print("5. Then follow the steps above to set as default")

def main():
    """Main function"""
    print("🔧 OUTLOOK DEFAULT ACCOUNT MANAGER")
    print("=" * 60)
    
    # Step 1: List all accounts
    accounts = list_all_accounts()
    if not accounts:
        print("❌ No accounts found or Outlook not accessible")
        return
    
    # Step 2: Show current default
    print("🔍 Current Default Account:")
    print("-" * 30)
    current_default = get_current_default()
    
    # Step 3: Check if target account exists
    print("\n🎯 Looking for azx1818@hotmail.com...")
    target_found = False
    for account in accounts:
        if account['email'].lower() == 'azx1818@hotmail.com':
            target_found = True
            print(f"✅ Found: {account['name']} ({account['email']})")
            
            # Test if we can use this account
            if test_account_sending(account['email']):
                print("✅ Account is ready for email automation")
            break
    
    if not target_found:
        print("❌ azx1818@hotmail.com not found in Outlook")
        print("📝 You need to add this account first")
    
    # Step 4: Provide instructions
    provide_manual_instructions()
    
    print("\n" + "=" * 60)
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()