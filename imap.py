#!/usr/bin/env python3
"""
Script to access a shared mailbox in Exchange Online via IMAP
and read the first 10 email subjects from the inbox.
"""

import imaplib
import email
import email.header
import sys
from getpass import getpass

def decode_header(header_value):
    """
    Decode email header that might be encoded (e.g., =?UTF-8?B?...?=)
    """
    if not header_value:
        return ""
    
    decoded_parts = email.header.decode_header(header_value)
    decoded_string = ""
    
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            if encoding:
                try:
                    decoded_string += part.decode(encoding)
                except (UnicodeDecodeError, LookupError):
                    decoded_string += part.decode('utf-8', errors='ignore')
            else:
                decoded_string += part.decode('utf-8', errors='ignore')
        else:
            decoded_string += str(part)
    
    return decoded_string

def connect_to_shared_mailbox():
    """
    Connect to Exchange Online shared mailbox via IMAP and read email subjects
    """
    # Exchange Online IMAP settings
    IMAP_SERVER = "outlook.office365.com"
    IMAP_PORT = 993
    
    # Get credentials
    print("Exchange Online Shared Mailbox Access")
    print("=" * 40)
    
    # User credentials (user with delegated access)
    username = input("Enter your email address: ")
    password = getpass("Enter your password: ")
    
    # Shared mailbox email
    shared_mailbox = input("Enter shared mailbox email address: ")
    
    try:
        # Connect to IMAP server
        print(f"\nConnecting to {IMAP_SERVER}...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        
        # Login with user credentials
        print("Authenticating...")
        mail.login(username, password)
        
        # Select the shared mailbox inbox
        # For shared mailboxes, use the format: shared_mailbox_email/INBOX
        shared_inbox = f"{shared_mailbox}/INBOX"
        print(f"Accessing shared mailbox: {shared_mailbox}")
        
        result = mail.select(shared_inbox)
        if result[0] != 'OK':
            print(f"Error: Could not access shared mailbox '{shared_mailbox}'")
            print("Make sure:")
            print("1. You have delegated access to this mailbox")
            print("2. The shared mailbox email address is correct")
            print("3. Your account has the necessary permissions")
            return
        
        # Search for all emails in the inbox
        print("\nSearching for emails...")
        result, message_ids = mail.search(None, 'ALL')
        
        if result != 'OK':
            print("Error searching for emails")
            return
        
        # Get list of email IDs
        email_ids = message_ids[0].split()
        
        if not email_ids:
            print("No emails found in the shared mailbox inbox.")
            return
        
        print(f"Found {len(email_ids)} emails. Reading first 10 subjects...\n")
        
        # Read first 10 emails (or all if less than 10)
        emails_to_read = min(10, len(email_ids))
        
        # Get the most recent emails (IMAP returns oldest first, so we reverse)
        recent_email_ids = email_ids[-emails_to_read:][::-1]
        
        print("Email Subjects:")
        print("-" * 50)
        
        for i, email_id in enumerate(recent_email_ids, 1):
            try:
                # Fetch email headers only for better performance
                result, msg_data = mail.fetch(email_id, '(RFC822.HEADER)')
                
                if result != 'OK':
                    print(f"{i:2d}. Error fetching email {email_id.decode()}")
                    continue
                
                # Parse email headers
                email_message = email.message_from_bytes(msg_data[0][1])
                
                # Get and decode subject
                subject = email_message.get('Subject', 'No Subject')
                decoded_subject = decode_header(subject)
                
                # Get sender
                from_header = email_message.get('From', 'Unknown Sender')
                decoded_from = decode_header(from_header)
                
                # Get date
                date_header = email_message.get('Date', 'Unknown Date')
                
                print(f"{i:2d}. {decoded_subject}")
                print(f"    From: {decoded_from}")
                print(f"    Date: {date_header}")
                print()
                
            except Exception as e:
                print(f"{i:2d}. Error processing email {email_id.decode()}: {str(e)}")
        
        # Close connection
        mail.close()
        mail.logout()
        print("Connection closed successfully.")
        
    except imaplib.IMAP4.error as e:
        print(f"IMAP Error: {str(e)}")
        print("\nTroubleshooting tips:")
        print("1. Verify your credentials are correct")
        print("2. Check if you have delegated access to the shared mailbox")
        print("3. Ensure the shared mailbox email address is correct")
        print("4. Try accessing the shared mailbox through Outlook first")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return

if __name__ == "__main__":
    try:
        connect_to_shared_mailbox()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        sys.exit(1)
