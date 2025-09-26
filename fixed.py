#!/usr/bin/env python3
"""
Exchange Online Mailbox Access using Microsoft Graph API
This script demonstrates how to authenticate and access your Exchange Online mailbox
using the Microsoft Graph API with username/password authentication.
"""

import requests
import json
from datetime import datetime
import urllib.parse

class ExchangeGraphClient:
    def __init__(self, tenant_id, client_id, username, password):
        """
        Initialize the Graph API client
        
        Args:
            tenant_id (str): Your Azure AD tenant ID
            client_id (str): Your registered application's client ID
            username (str): Your email address
            password (str): Your password
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.username = username
        self.password = password
        self.access_token = None
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"
        
    def authenticate(self):
        """
        Authenticate using Resource Owner Password Credentials (ROPC) flow
        Note: This requires specific Azure AD configuration
        """
        auth_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        # Try with specific scopes first
        data = {
            'client_id': self.client_id,
            'scope': 'https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read',
            'username': self.username,
            'password': self.password,
            'grant_type': 'password'
        }
        
        try:
            response = requests.post(auth_url, headers=headers, data=data)
            
            if response.status_code == 400:
                # If specific scopes fail, try with .default scope
                print("🔄 Trying alternative authentication method...")
                data['scope'] = 'https://graph.microsoft.com/.default'
                response = requests.post(auth_url, headers=headers, data=data)
            
            if response.status_code == 400:
                # Parse the error response
                error_data = response.json()
                error_code = error_data.get('error', '')
                error_description = error_data.get('error_description', '')
                
                if 'AADSTS65001' in error_description:
                    print("❌ Consent Required Error:")
                    print("   The application needs admin consent for the requested permissions.")
                    print("   Please follow these steps:")
                    print("   1. Go to Azure Portal > Azure Active Directory > App registrations")
                    print("   2. Find your application and go to 'API permissions'")
                    print("   3. Click 'Grant admin consent for [your organization]'")
                    print("   4. Alternatively, use the admin consent URL below:")
                    admin_consent_url = f"https://login.microsoftonline.com/{self.tenant_id}/adminconsent?client_id={self.client_id}"
                    print(f"   {admin_consent_url}")
                    return False
                elif 'AADSTS50076' in error_description:
                    print("❌ MFA Required:")
                    print("   Multi-factor authentication is required but not supported with password flow.")
                    print("   Please disable MFA for this account or use interactive authentication.")
                    return False
                elif 'AADSTS50034' in error_description:
                    print("❌ User Not Found:")
                    print("   The username doesn't exist in this tenant.")
                    return False
                elif 'AADSTS50126' in error_description:
                    print("❌ Invalid Credentials:")
                    print("   Username or password is incorrect.")
                    return False
                else:
                    print(f"❌ Authentication error: {error_code}")
                    print(f"   Description: {error_description}")
                    return False
            
            response.raise_for_status()
            token_data = response.json()
            self.access_token = token_data.get('access_token')
            
            if self.access_token:
                print("✅ Authentication successful!")
                return True
            else:
                print("❌ Authentication failed: No access token received")
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"❌ Authentication error: {e}")
            if hasattr(e, 'response') and e.response and e.response.text:
                try:
                    error_data = e.response.json()
                    print(f"Error details: {error_data}")
                except:
                    print(f"Response: {e.response.text}")
            return False
    
    def _make_request(self, endpoint, method='GET', data=None):
        """
        Make an authenticated request to the Graph API
        """
        if not self.access_token:
            print("❌ No access token. Please authenticate first.")
            return None
            
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.graph_endpoint}{endpoint}"
        
        try:
            if method == 'GET':
                response = requests.get(url, headers=headers)
            elif method == 'POST':
                response = requests.post(url, headers=headers, json=data)
            elif method == 'PATCH':
                response = requests.patch(url, headers=headers, json=data)
            elif method == 'DELETE':
                response = requests.delete(url, headers=headers)
                
            response.raise_for_status()
            
            if response.content:
                return response.json()
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"❌ API request failed: {e}")
            if hasattr(e.response, 'text'):
                print(f"Response: {e.response.text}")
            return None
    
    def get_user_profile(self):
        """Get the current user's profile information"""
        print("📋 Fetching user profile...")
        result = self._make_request("/me")
        
        if result:
            print(f"Name: {result.get('displayName')}")
            print(f"Email: {result.get('mail')}")
            print(f"Job Title: {result.get('jobTitle', 'N/A')}")
            print(f"Office Location: {result.get('officeLocation', 'N/A')}")
        
        return result
    
    def get_mailbox_folders(self):
        """Get all mailbox folders"""
        print("📁 Fetching mailbox folders...")
        result = self._make_request("/me/mailFolders")
        
        if result and 'value' in result:
            folders = result['value']
            print(f"Found {len(folders)} folders:")
            for folder in folders:
                print(f"  - {folder['displayName']} ({folder['totalItemCount']} items)")
        
        return result
    
    def get_messages(self, folder_id='inbox', top=10):
        """
        Get messages from a specific folder
        
        Args:
            folder_id (str): Folder ID or name ('inbox', 'sentitems', etc.)
            top (int): Number of messages to retrieve
        """
        print(f"📧 Fetching {top} messages from {folder_id}...")
        
        endpoint = f"/me/mailFolders/{folder_id}/messages"
        if top:
            endpoint += f"?$top={top}&$select=subject,from,receivedDateTime,isRead,bodyPreview"
            
        result = self._make_request(endpoint)
        
        if result and 'value' in result:
            messages = result['value']
            print(f"Found {len(messages)} messages:")
            
            for i, msg in enumerate(messages, 1):
                from_addr = msg.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
                subject = msg.get('subject', 'No Subject')
                received = msg.get('receivedDateTime', '')
                is_read = msg.get('isRead', False)
                preview = msg.get('bodyPreview', '')[:100] + '...' if msg.get('bodyPreview') else ''
                
                status = "📖" if is_read else "📩"
                
                print(f"\n{i}. {status} From: {from_addr}")
                print(f"   Subject: {subject}")
                print(f"   Received: {received}")
                print(f"   Preview: {preview}")
        
        return result
    
    def search_messages(self, query, top=10):
        """
        Search for messages
        
        Args:
            query (str): Search query
            top (int): Number of results to return
        """
        print(f"🔍 Searching for messages with query: '{query}'...")
        
        encoded_query = urllib.parse.quote(query)
        endpoint = f"/me/messages?$search=\"{encoded_query}\"&$top={top}&$select=subject,from,receivedDateTime,bodyPreview"
        
        result = self._make_request(endpoint)
        
        if result and 'value' in result:
            messages = result['value']
            print(f"Found {len(messages)} matching messages:")
            
            for i, msg in enumerate(messages, 1):
                from_addr = msg.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
                subject = msg.get('subject', 'No Subject')
                received = msg.get('receivedDateTime', '')
                
                print(f"\n{i}. From: {from_addr}")
                print(f"   Subject: {subject}")
                print(f"   Received: {received}")
        
        return result
    
    def send_message(self, to_email, subject, body, body_type='Text'):
        """
        Send an email message
        
        Args:
            to_email (str): Recipient email address
            subject (str): Email subject
            body (str): Email body content
            body_type (str): 'Text' or 'HTML'
        """
        print(f"📤 Sending message to {to_email}...")
        
        message_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            }
        }
        
        result = self._make_request("/me/sendMail", method='POST', data=message_data)
        
        if result is not None:
            print("✅ Message sent successfully!")
        
        return result


def main():
    """
    Main function to demonstrate the Exchange Graph API client
    """
    print("🔐 Exchange Online Graph API Client")
    print("=" * 40)
    
    # Configuration - Replace with your actual values
    TENANT_ID = "your-tenant-id"  # Your Azure AD tenant ID
    CLIENT_ID = "your-client-id"  # Your registered app's client ID
    USERNAME = "your-email@yourdomain.com"  # Your email address
    PASSWORD = "your-password"  # Your password
    
    # Check if configuration is set
    if any(val.startswith("your-") for val in [TENANT_ID, CLIENT_ID, USERNAME, PASSWORD]):
        print("⚠️  Please update the configuration variables with your actual values:")
        print("   - TENANT_ID: Your Azure AD tenant ID")
        print("   - CLIENT_ID: Your registered application's client ID")
        print("   - USERNAME: Your email address")
        print("   - PASSWORD: Your password")
        print("\n📖 See the setup instructions in the comments above.")
        return
    
    # Initialize client
    client = ExchangeGraphClient(TENANT_ID, CLIENT_ID, USERNAME, PASSWORD)
    
    # Authenticate
    if not client.authenticate():
        print("❌ Authentication failed. Please check your credentials.")
        return
    
    try:
        # Get user profile
        print("\n" + "="*40)
        client.get_user_profile()
        
        # Get mailbox folders
        print("\n" + "="*40)
        client.get_mailbox_folders()
        
        # Get recent messages from inbox
        print("\n" + "="*40)
        client.get_messages('inbox', top=5)
        
        # Example: Search for messages
        print("\n" + "="*40)
        search_query = input("Enter search query (or press Enter to skip): ").strip()
        if search_query:
            client.search_messages(search_query, top=5)
        
        # Example: Send a message (uncomment to use)
        # send_to = input("Send test email to (or press Enter to skip): ").strip()
        # if send_to:
        #     client.send_message(
        #         to_email=send_to,
        #         subject="Test from Graph API",
        #         body="This is a test message sent via Microsoft Graph API!"
        #     )
        
    except KeyboardInterrupt:
        print("\n\n👋 Exiting...")
    except Exception as e:
        print(f"❌ An error occurred: {e}")


if __name__ == "__main__":
    main()


"""
SETUP INSTRUCTIONS:
==================

1. Register an Application in Azure AD:
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Click "New registration"
   - Name: "Exchange API Client" (or any name)
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

2. Configure Application Permissions:
   - Go to "API permissions" in your app
   - Click "Add a permission" > "Microsoft Graph" > "Delegated permissions"
   - Add these permissions:
     * Mail.ReadWrite
     * Mail.Send
     * User.Read
   - ⚠️  IMPORTANT: Click "Grant admin consent for [Your Organization]" button
   - The status should show green checkmarks for all permissions

3. Enable Public Client:
   - Go to "Authentication" in your app
   - Under "Advanced settings" > "Allow public client flows" > Select "Yes"
   - Click "Save"

4. Alternative: Use Admin Consent URL (if step 2 doesn't work):
   - Replace YOUR_TENANT_ID and YOUR_CLIENT_ID in this URL:
   - https://login.microsoftonline.com/YOUR_TENANT_ID/adminconsent?client_id=YOUR_CLIENT_ID
   - Open this URL in a browser and consent as admin

5. Check ROPC Policy (if authentication still fails):
   - Go to Azure AD > Security > Conditional Access
   - Ensure no policies block the Resource Owner Password Credentials flow
   - Or create an exclusion for your test user/app

6. Get Required Information:
   - TENANT_ID: Go to Azure AD > Overview > copy "Tenant ID"
   - CLIENT_ID: Go to your app registration > Overview > copy "Application (client) ID"
   - USERNAME: Your email address
   - PASSWORD: Your password

7. Install Required Package:
   pip install requests

8. Update the configuration variables in the main() function with your actual values.

FIXING "CONSENT REQUIRED" ERROR:
===============================
The script now provides specific guidance for consent errors:

1. Admin Consent Method:
   - Use the admin consent URL provided in the error message
   - OR go to your app > API permissions > "Grant admin consent"

2. If you're not an admin:
   - Ask your Azure AD administrator to grant consent
   - Or have them add you to the "Cloud Application Administrator" role temporarily

3. Alternative Authentication (if ROPC is blocked):
   - Consider using device code flow or interactive authentication
   - ROPC may be disabled by organizational policy

IMPORTANT SECURITY NOTES:
========================
- The Resource Owner Password Credentials (ROPC) flow is used here for simplicity
- This flow may be blocked by default in many organizations
- For production use, consider using more secure flows like Authorization Code flow
- Never hardcode credentials in production code - use environment variables or secure vaults
- This script stores credentials in memory only and doesn't persist them

TROUBLESHOOTING:
===============
- If authentication fails, check if ROPC is enabled in your Azure AD
- Ensure your account has the necessary permissions
- Check if multi-factor authentication is required (ROPC doesn't support MFA)
- Verify that the application permissions are properly granted and consented
- Some organizations disable ROPC flow entirely for security reasons
"""
