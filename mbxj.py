#!/usr/bin/env python3
"""
Microsoft Graph API Shared Mailbox Access Script
This script demonstrates how to access a specific shared mailbox using Graph API
with delegated permissions (no admin consent required for broad access).
"""

import requests
import json
from datetime import datetime, timedelta
import base64
import msal
import logging
from typing import List, Dict, Optional

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GraphSharedMailboxClient:
    def __init__(self, client_id: str, client_secret: str, tenant_id: str, shared_mailbox_email: str):
        """
        Initialize the Graph API client for shared mailbox access
        
        Args:
            client_id (str): Azure App Registration Client ID
            client_secret (str): Azure App Registration Client Secret
            tenant_id (str): Azure Tenant ID
            shared_mailbox_email (str): Email address of the shared mailbox
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.shared_mailbox_email = shared_mailbox_email
        self.access_token = None
        self.graph_url = "https://graph.microsoft.com/v1.0"
        
        # Required scopes - these are delegated permissions
        self.scopes = [
            "https://graph.microsoft.com/Mail.Read.Shared",
            "https://graph.microsoft.com/Mail.Send.Shared",
            "https://graph.microsoft.com/Mail.ReadWrite.Shared"
        ]
    
    def authenticate(self, username: str, password: str) -> bool:
        """
        Authenticate using Resource Owner Password Credentials (ROPC) flow
        Note: ROPC should only be used when other flows are not feasible
        
        Args:
            username (str): Your username
            password (str): Your password
            
        Returns:
            bool: True if authentication successful
        """
        try:
            # Create MSAL app
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}"
            )
            
            # Acquire token using username/password
            result = app.acquire_token_by_username_password(
                username=username,
                password=password,
                scopes=self.scopes
            )
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Successfully authenticated with Graph API")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def authenticate_interactive(self) -> bool:
        """
        Interactive authentication using device code flow
        This will handle consent automatically when you first authenticate
        
        Returns:
            bool: True if authentication successful
        """
        try:
            # Create MSAL public client app (better for interactive auth)
            app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}"
            )
            
            # Try to get token from cache first
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(self.scopes, account=accounts[0])
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    logger.info("Token acquired from cache")
                    return True
            
            # If no cached token, initiate device flow
            print("🔐 Starting authentication process...")
            print("⚠️  If this is your first time, you'll be asked to consent to permissions.")
            
            flow = app.initiate_device_flow(scopes=self.scopes)
            if "user_code" not in flow:
                raise Exception("Failed to create device flow")
            
            print("\n" + "="*60)
            print("AUTHENTICATION REQUIRED")
            print("="*60)
            print(flow["message"])
            print("="*60)
            print("ℹ️  After entering the code, you may see a consent screen.")
            print("ℹ️  Click 'Accept' to grant permissions for the shared mailbox access.")
            print()
            
            # Wait for user to complete authentication
            result = app.acquire_token_by_device_flow(flow)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Successfully authenticated with Graph API")
                print("✅ Authentication successful! Consent has been granted.")
                return True
            else:
                error_msg = result.get('error_description', 'Unknown error')
                logger.error(f"Authentication failed: {error_msg}")
                
                # Provide specific guidance for common consent issues
                if "consent" in error_msg.lower():
                    print("\n❌ Consent was not granted. Please try again and click 'Accept' on the consent screen.")
                elif "AADSTS65001" in error_msg:
                    print("\n❌ User declined consent. Please try again and accept the permissions.")
                
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def get_consent_url(self) -> str:
        """
        Generate a consent URL that can be opened in browser manually
        Useful if device flow doesn't work in your environment
        
        Returns:
            str: Consent URL
        """
        scope_param = " ".join(self.scopes)
        consent_url = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize?"
            f"client_id={self.client_id}&"
            f"response_type=code&"
            f"redirect_uri=http://localhost&"
            f"scope={scope_param}&"
            f"response_mode=query"
        )
        return consent_url
    
    def _make_request(self, endpoint: str, method: str = "GET", data: Dict = None) -> Optional[Dict]:
        """
        Make authenticated request to Graph API
        
        Args:
            endpoint (str): API endpoint
            method (str): HTTP method
            data (dict): Request payload
            
        Returns:
            dict: Response data or None if error
        """
        if not self.access_token:
            logger.error("No access token available. Please authenticate first.")
            return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"{self.graph_url}{endpoint}"
        
        try:
            if method.upper() == "GET":
                response = requests.get(url, headers=headers)
            elif method.upper() == "POST":
                response = requests.post(url, headers=headers, json=data)
            elif method.upper() == "PATCH":
                response = requests.patch(url, headers=headers, json=data)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")
            
            response.raise_for_status()
            return response.json() if response.content else {}
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Request failed: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response: {e.response.text}")
            return None
    
    def get_mailbox_folders(self) -> List[Dict]:
        """
        Get all folders in the shared mailbox
        
        Returns:
            list: List of folder objects
        """
        endpoint = f"/users/{self.shared_mailbox_email}/mailFolders"
        result = self._make_request(endpoint)
        
        if result and "value" in result:
            logger.info(f"Retrieved {len(result['value'])} folders from shared mailbox")
            return result["value"]
        return []
    
    def get_messages(self, folder_id: str = "inbox", limit: int = 10, days_back: int = 7) -> List[Dict]:
        """
        Get messages from specified folder
        
        Args:
            folder_id (str): Folder ID or well-known name (inbox, sentitems, etc.)
            limit (int): Maximum number of messages
            days_back (int): How many days back to search
            
        Returns:
            list: List of message objects
        """
        # Calculate date filter
        start_date = (datetime.now() - timedelta(days=days_back)).isoformat() + "Z"
        
        # Build endpoint with filter
        endpoint = f"/users/{self.shared_mailbox_email}/mailFolders/{folder_id}/messages"
        endpoint += f"?$filter=receivedDateTime ge {start_date}"
        endpoint += f"&$orderby=receivedDateTime desc"
        endpoint += f"&$top={limit}"
        endpoint += "&$select=id,subject,sender,receivedDateTime,isRead,hasAttachments,bodyPreview"
        
        result = self._make_request(endpoint)
        
        if result and "value" in result:
            logger.info(f"Retrieved {len(result['value'])} messages from {folder_id}")
            return result["value"]
        return []
    
    def get_message_details(self, message_id: str) -> Optional[Dict]:
        """
        Get full details of a specific message
        
        Args:
            message_id (str): Message ID
            
        Returns:
            dict: Full message object
        """
        endpoint = f"/users/{self.shared_mailbox_email}/messages/{message_id}"
        return self._make_request(endpoint)
    
    def send_email(self, to_recipients: List[str], subject: str, body: str, 
                   cc_recipients: List[str] = None, body_type: str = "Text") -> bool:
        """
        Send email from shared mailbox
        
        Args:
            to_recipients (list): List of recipient email addresses
            subject (str): Email subject
            body (str): Email body
            cc_recipients (list): List of CC recipients
            body_type (str): "Text" or "HTML"
            
        Returns:
            bool: True if successful
        """
        # Build recipient objects
        to_list = [{"emailAddress": {"address": addr}} for addr in to_recipients]
        cc_list = [{"emailAddress": {"address": addr}} for addr in (cc_recipients or [])]
        
        message_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body
                },
                "toRecipients": to_list,
                "ccRecipients": cc_list
            }
        }
        
        endpoint = f"/users/{self.shared_mailbox_email}/sendMail"
        result = self._make_request(endpoint, method="POST", data=message_data)
        
        if result is not None:
            logger.info(f"Email sent successfully to {', '.join(to_recipients)}")
            return True
        return False
    
    def search_messages(self, search_term: str, folder_id: str = "inbox", limit: int = 10) -> List[Dict]:
        """
        Search for messages containing specific text
        
        Args:
            search_term (str): Text to search for
            folder_id (str): Folder to search in
            limit (int): Maximum results
            
        Returns:
            list: List of matching messages
        """
        endpoint = f"/users/{self.shared_mailbox_email}/mailFolders/{folder_id}/messages"
        endpoint += f"?$search=\"{search_term}\""
        endpoint += f"&$top={limit}"
        endpoint += "&$select=id,subject,sender,receivedDateTime,isRead,hasAttachments,bodyPreview"
        
        result = self._make_request(endpoint)
        
        if result and "value" in result:
            logger.info(f"Found {len(result['value'])} messages matching '{search_term}'")
            return result["value"]
        return []
    
    def mark_as_read(self, message_id: str) -> bool:
        """
        Mark a message as read
        
        Args:
            message_id (str): Message ID
            
        Returns:
            bool: True if successful
        """
        endpoint = f"/users/{self.shared_mailbox_email}/messages/{message_id}"
        data = {"isRead": True}
        
        result = self._make_request(endpoint, method="PATCH", data=data)
        return result is not None
    
    def display_messages(self, messages: List[Dict]):
        """
        Display message information in a readable format
        
        Args:
            messages (list): List of message objects
        """
        print(f"\n{'='*60}")
        print(f"SHARED MAILBOX: {self.shared_mailbox_email}")
        print(f"{'='*60}")
        
        for i, msg in enumerate(messages, 1):
            sender = msg.get("sender", {}).get("emailAddress", {}).get("address", "Unknown")
            subject = msg.get("subject", "No Subject")
            received = msg.get("receivedDateTime", "")
            is_read = msg.get("isRead", False)
            has_attachments = msg.get("hasAttachments", False)
            preview = msg.get("bodyPreview", "")[:100] + "..." if len(msg.get("bodyPreview", "")) > 100 else msg.get("bodyPreview", "")
            
            # Format date
            if received:
                try:
                    dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
                    received = dt.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass
            
            print(f"\nMessage {i}:")
            print(f"  From: {sender}")
            print(f"  Subject: {subject}")
            print(f"  Received: {received}")
            print(f"  Read: {'Yes' if is_read else 'No'}")
            print(f"  Attachments: {'Yes' if has_attachments else 'No'}")
            print(f"  Preview: {preview}")
            print("-" * 50)
    
    def display_folders(self, folders: List[Dict]):
        """
        Display folder information
        
        Args:
            folders (list): List of folder objects
        """
        print(f"\n{'='*40}")
        print(f"MAILBOX FOLDERS")
        print(f"{'='*40}")
        
        for folder in folders:
            name = folder.get("displayName", "Unknown")
            total_count = folder.get("totalItemCount", 0)
            unread_count = folder.get("unreadItemCount", 0)
            folder_id = folder.get("id", "")
            
            print(f"📁 {name}")
            print(f"   Total: {total_count}, Unread: {unread_count}")
            print(f"   ID: {folder_id}")
            print()


def main():
    """
    Main function demonstrating script usage
    """
    # Azure App Registration Configuration
    CLIENT_ID = "your-client-id"  # Azure App Registration Client ID
    CLIENT_SECRET = "your-client-secret"  # Azure App Registration Client Secret  
    TENANT_ID = "your-tenant-id"  # Azure Tenant ID
    SHARED_MAILBOX = "shared-mailbox@domain.com"  # Shared mailbox email
    
    # User credentials (only needed for username/password auth)
    USERNAME = "bilbo@domain.com"  # Your user account
    PASSWORD = "your-password"  # Your password
    
    print("Microsoft Graph API Shared Mailbox Client")
    print("=" * 50)
    
    # Create client
    client = GraphSharedMailboxClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID, SHARED_MAILBOX)
    
    # Choose authentication method
    print("Choose authentication method:")
    print("1. Username/Password (ROPC)")
    print("2. Interactive Device Code (Recommended)")
    print("3. Generate consent URL manually")
    
    auth_method = input("\nEnter choice (1, 2, or 3): ").strip()
    
    if auth_method == "1":
        print(f"\n🔑 Authenticating as {USERNAME}...")
        if not client.authenticate(USERNAME, PASSWORD):
            print("❌ Authentication failed!")
            print("💡 Try interactive authentication (option 2) if you haven't consented yet.")
            return
    elif auth_method == "3":
        # Generate consent URL for manual consent
        consent_url = client.get_consent_url()
        print(f"\n🌐 Open this URL in your browser to provide consent:")
        print(f"\n{consent_url}\n")
        print("📋 Steps:")
        print("1. Copy the URL above")
        print("2. Open it in your browser")
        print("3. Login as bilbo@domain.com")
        print("4. Click 'Accept' on the consent screen")
        print("5. Come back and run the script again with option 1 or 2")
        return
    else:
        # Interactive authentication (handles consent automatically)
        print(f"\n🔑 Starting interactive authentication for {USERNAME}...")
        if not client.authenticate_interactive():
            print("❌ Authentication failed!")
            print("💡 Make sure to accept the consent screen when prompted.")
            return
    
    print(f"\n✅ Successfully connected to shared mailbox: {SHARED_MAILBOX}")
    print(f"👤 Authenticated as: bilbo")
    
    try:
        # Display available folders
        print("\n📂 Getting folder list...")
        folders = client.get_mailbox_folders()
        client.display_folders(folders)
        
        # Get recent messages from inbox
        print("\n📧 Getting recent inbox messages...")
        messages = client.get_messages(folder_id="inbox", limit=5, days_back=30)
        client.display_messages(messages)
        
        # Search for messages
        search_term = input("\n🔍 Enter search term (or press Enter to skip): ").strip()
        if search_term:
            search_results = client.search_messages(search_term, limit=3)
            if search_results:
                print(f"\n📋 Search results for '{search_term}':")
                client.display_messages(search_results)
            else:
                print(f"No messages found matching '{search_term}'")
        
        # Optional: Send email
        send_email = input("\n📤 Send test email? (y/n): ").strip().lower()
        if send_email == 'y':
            recipient = input("Enter recipient email: ").strip()
            if recipient:
                success = client.send_email(
                    to_recipients=[recipient],
                    subject="Test from Shared Mailbox via Graph API",
                    body="This is a test email sent from the shared mailbox using Microsoft Graph API and Python.",
                    body_type="Text"
                )
                print(f"📧 Email sent: {'✅ Success' if success else '❌ Failed'}")
        
    except Exception as e:
        logger.error(f"Error during operations: {str(e)}")
        
        # Provide helpful error messages
        if "Insufficient privileges" in str(e):
            print("\n❌ Error: Insufficient privileges")
            print("💡 Make sure:")
            print("   - You have 'Full Access' to the shared mailbox")
            print("   - You've provided consent for the required permissions")
        elif "Forbidden" in str(e):
            print("\n❌ Error: Access forbidden")
            print("💡 This usually means consent hasn't been granted yet.")
            print("   Try running the script again with option 3 to generate a consent URL.")


def setup_instructions():
    """
    Print setup instructions for first-time users
    """
    print("\n" + "="*70)
    print("FIRST TIME SETUP INSTRUCTIONS")
    print("="*70)
    print("1. 🏢 Azure App Registration Setup:")
    print("   - Go to Azure Portal → App Registrations → New Registration")
    print("   - Name: 'Shared Mailbox Access'")
    print("   - Redirect URI: http://localhost (Public client)")
    print("")
    print("2. 🔐 Add Delegated Permissions:")
    print("   - Go to 'API Permissions'")
    print("   - Add permission → Microsoft Graph → Delegated permissions")
    print("   - Add: Mail.Read.Shared, Mail.Send.Shared, Mail.ReadWrite.Shared")
    print("")
    print("3. ⚙️  Configure Authentication:")
    print("   - Go to 'Authentication'")
    print("   - Enable 'Allow public client flows' = Yes")
    print("")
    print("4. 🔑 Create Client Secret (if using confidential client):")
    print("   - Go to 'Certificates & secrets'")
    print("   - New client secret")
    print("")
    print("5. 🎯 Grant User Consent:")
    print("   - Run this script and choose option 3 to generate consent URL")
    print("   - Or run option 2 for interactive consent")
    print("")
    print("6. 📧 Ensure Mailbox Access:")
    print("   - Admin must grant 'Full Access' to shared mailbox for user 'bilbo'")
    print("="*70)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--setup":
        setup_instructions()
    else:
        main()
