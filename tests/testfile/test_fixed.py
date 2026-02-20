#!/usr/bin/env python3
"""
Fixed version of test.py - with email limit and progress tracking
Prevents hanging by limiting scan to recent emails only
"""

import win32com.client
import pandas as pd
import json
import re
from typing import Dict, List, Any, Optional
from datetime import datetime, timedelta

# === CONFIGURATION ===
MAX_EMAILS_TO_SCAN = 200  # Limit to most recent 200 emails (adjust as needed)
SCAN_LAST_N_DAYS = 30     # Only scan emails from last 30 days
ENABLE_PROGRESS = True     # Show progress messages

class EmailFilterManager:
    """Manages email filters with Power Automate expression language support"""
    
    def __init__(self):
        self.filters = []
        self.filter_file = "email_filters.json"
        self.load_filters()
    
    def load_filters(self):
        """Load filters from JSON file"""
        try:
            with open(self.filter_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    self.filters = json.loads(content)
                else:
                    self.filters = []
        except (FileNotFoundError, json.JSONDecodeError):
            self.filters = []
        
    def save_filters(self):
        """Save filters to JSON file"""
        with open(self.filter_file, "w", encoding="utf-8") as f:
            json.dump(self.filters, f, indent=4, ensure_ascii=False)
    
    def get_next_id(self) -> int:
        """Get the next available ID for a new filter"""
        if not self.filters:
            return 1
        max_id = max(f['id'] for f in self.filters)
        return max_id + 1
    
    def create_filter(self, name: str, from_email: str = "", subject_filter: str = "", 
                      body_filter: str = "", to_email: str = "", action: str = "", description: str = "") -> bool:
        """Create a new email filter"""
        filter_data = {
            "id": self.get_next_id(),
            "name": name,
            "from_email": from_email,
            "subject_filter": subject_filter,
            "body_filter": body_filter,
            "to_email": to_email,
            "action": action,
            "description": description,
            "enabled": True,
            "created_at": str(pd.Timestamp.now())
        }
        self.filters.append(filter_data)
        self.save_filters()
        return True
    
    def evaluate_power_automate_expression(self, expression: str, email_data: Dict) -> bool:
        """Evaluate Power Automate expression language against email data"""
        
        # Helper functions that mimic Power Automate
        def equals(value1: Any, value2: Any) -> bool:
            return str(value1).lower() == str(value2).lower()
        
        def contains(text: str, search: str) -> bool:
            return search.lower() in text.lower()
        
        def startswith(text: str, search: str) -> bool:
            return str(text).lower().startswith(str(search).lower())
        
        def endswith(text: str, search: str) -> bool:
            return str(text).lower().endswith(str(search).lower())
        
        def length(text: str) -> int:
            return len(str(text))
        
        def empty(value: Any) -> bool:
            return not value or str(value).strip() == ""
        
        def not_empty(value: Any) -> bool:
            return not empty(value)
        
        def and_expr(*conditions) -> bool:
            return all(conditions)
        
        def or_expr(*conditions) -> bool:
            return any(conditions)
        
        def not_expr(condition: bool) -> bool:
            return not condition
        
        # Get email fields for the expression
        subject = email_data.get('subject', '')
        body = email_data.get('body', '')
        sender = email_data.get('sender', '')
        
        try:
            # Replace Power Automate function names with Python equivalents
            # Note: This is a simple implementation. For complex expressions,
            # you might need a proper expression parser
            
            # Allow direct function calls in the expression
            result = eval(expression, {
                "equals": equals,
                "contains": contains,
                "startswith": startswith,
                "endswith": endswith,
                "length": length,
                "empty": empty,
                "not_empty": not_empty,
                "and": and_expr,
                "or": or_expr,
                "not": not_expr,
                "subject": subject,
                "body": body,
                "sender": sender
            })
            
            return bool(result)
        except Exception as e:
            print(f"Error evaluating expression: {e}")
            return False
    
    def apply_filters(self, email_data: Dict) -> List[str]:
        """Apply all enabled filters to email data and return matching actions"""
        matching_actions = []
        
        for filter_data in self.filters:
            if not filter_data.get('enabled', True):
                continue
            
            # Check from email filter
            if filter_data.get('from_email'):
                if not contains(email_data.get('sender', ''), filter_data['from_email']):
                    continue
            
            # Check subject filter (Power Automate expression)
            if filter_data.get('subject_filter'):
                if not self.evaluate_power_automate_expression(
                    filter_data['subject_filter'], email_data):
                    continue
            
            # Check body filter (Power Automate expression)
            if filter_data.get('body_filter'):
                if not self.evaluate_power_automate_expression(
                    filter_data['body_filter'], email_data):
                    continue
            
            # Check TO email filter (only type=1 recipients)
            if filter_data.get('to_email'):
                to_recipients = [r['name'] for r in email_data.get('recipients', []) if r.get('type') == 1]
                if not any(filter_data['to_email'] in to_name for to_name in to_recipients):
                    continue
            
            # If all conditions match, add the action
            if filter_data.get('action'):
                matching_actions.append(filter_data['action'])
        
        return matching_actions
    
    def extract_mx_data(self, email_data: Dict) -> Dict:
        """Extract MX specific data from email body"""
        body_text = email_data.get('body', '')
        
        def trim(text: str) -> str:
            return text.strip()
        
        def first(items: List[str]) -> str:
            return items[0] if items else ''
        
        def split(text: str, separator: str) -> List[str]:
            return text.split(separator)
        
        def substring(text: str, start_index: int, length: Optional[int] = None) -> str:
            if length is None:
                return text[start_index:]
            return text[start_index:start_index + length]
        
        def index_of(text: str, search: str) -> int:
            return text.lower().find(search.lower())
        
        def equals(value1: Any, value2: Any) -> bool:
            return str(value1).lower() == str(value2).lower()
        
        def concat(*args) -> str:
            return "".join(str(arg) for arg in args)
        
        mx_data = {}
        
        try:
            # Extract ticket_number
            number_index = index_of(body_text, 'Number:')
            if number_index != -1:
                number_text = substring(body_text, number_index + 7, 200)
                number_split = split(number_text, 'User:')
                ticket_number = trim(first(number_split))
                mx_data['ticket_number'] = ticket_number
            
            # Extract shop
            location_index = index_of(body_text, 'Location:')
            if location_index != -1:
                location_text = substring(body_text, location_index + 9, 200)
                location_split = split(location_text, 'Category:')
                shop_raw = trim(first(location_split))
                
                shop_parts = split(shop_raw, '-')
                shop_base = trim(first(shop_parts))
                
                if equals(substring(shop_base, 0, 1), '0'):
                    shop_processed = concat('MX', substring(shop_base, 1))
                else:
                    shop_processed = concat('MX', shop_base)
                
                mx_data['shop'] = shop_processed
            
            # Extract description
            desc_index = index_of(body_text, 'Short Description:')
            if desc_index != -1:
                desc_text = substring(body_text, desc_index + 18, 500)
                desc_split = split(desc_text, '\r\n')
                description = trim(first(desc_split))
                mx_data['description'] = description
                
        except Exception as e:
            print(f"Error extracting MX data: {e}")
        
        return mx_data

    def extract_cdc_data(self, email_data: Dict) -> Dict:
        """Extract CDC specific data with Regex"""
        body = email_data.get('body', '')
        data = {}
        try:
            # Extract ticket number
            ticket_match = re.search(r'Inci\. ID:\s*([A-Z0-9]+)', body)
            if ticket_match:
                data['ticket_number'] = ticket_match.group(1)
            
            # Extract shop
            shop_match = re.search(r'Cust\. Name:\s*.*?\((.*?)\)', body, re.DOTALL)
            if shop_match:
                shop = shop_match.group(1).strip()
                # Add "cdc" prefix if shop doesn't start with "ss"
                if not shop.lower().startswith('ss'):
                    shop = f'cdc{shop}'
                data['shop'] = shop

            # Extract description
            desc_match = re.search(r'Description:\s*(.*?)\r\n', body, re.DOTALL)
            if desc_match:
                data['description'] = desc_match.group(1).strip()
                
        except Exception as e:
            print(f"Error extracting CDC data: {e}")
        return data

# Helper function for backward compatibility
def contains(text: str, search: str) -> bool:
    return search.lower() in text.lower()

if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    import sys
    if sys.platform == 'win32':
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
    
    print("=" * 60)
    print("EMAIL SCANNER - FIXED VERSION")
    print("=" * 60)
    print(f"Configuration:")
    print(f"  - Max emails to scan: {MAX_EMAILS_TO_SCAN}")
    print(f"  - Scan last N days: {SCAN_LAST_N_DAYS}")
    print(f"  - Progress tracking: {'Enabled' if ENABLE_PROGRESS else 'Disabled'}")
    print("=" * 60)
    
    # Calculate cutoff date
    cutoff_date = datetime.now() - timedelta(days=SCAN_LAST_N_DAYS)
    
    print(f"\n[1/4] Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        print(f"[OK] Connected to Outlook inbox")
    except Exception as e:
        print(f"[ERROR] Error connecting to Outlook: {e}")
        exit(1)
    
    print(f"\n[2/4] Fetching recent messages...")
    try:
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by date, newest first
        
        # Restrict to recent emails for performance
        total_in_inbox = messages.Count
        print(f"[OK] Found {total_in_inbox} total emails in inbox")
        print(f"  Limiting scan to {MAX_EMAILS_TO_SCAN} most recent emails")
    except Exception as e:
        print(f"[ERROR] Error fetching messages: {e}")
        exit(1)
    
    print(f"\n[3/4] Processing emails...")
    
    # Initialize filter manager
    filter_manager = EmailFilterManager()
    print(f"  Loaded {len(filter_manager.filters)} email filters")
    
    # Collect all emails data
    all_emails = []
    processed_count = 0
    skipped_old = 0
    skipped_error = 0
    
    for i, message in enumerate(messages):
        # Limit number of emails processed
        if processed_count >= MAX_EMAILS_TO_SCAN:
            print(f"\n  Reached limit of {MAX_EMAILS_TO_SCAN} emails")
            break
        
        try:
            # Check if email is too old
            try:
                email_date = message.ReceivedTime
                if email_date < cutoff_date:
                    skipped_old += 1
                    continue
            except:
                pass  # If we can't get date, process it anyway
            
            # Progress indicator
            if ENABLE_PROGRESS and processed_count % 10 == 0:
                print(f"  Processing email {processed_count + 1}...", end='\r')
            
            # Get sender info
            try:
                sender_name = message.SenderName
            except:
                sender_name = "Unknown"
            
            sender_email = ""
            try:
                if message.SenderEmailType == "EX":
                    sender_email = message.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    sender_email = message.SenderEmailAddress
            except:
                pass
            
            sender = f"{sender_name} <{sender_email}>"
            
            # Get basic email data
            try:
                date = str(message.ReceivedTime)
            except:
                date = "Unknown"
            
            try:
                subject = message.Subject
            except:
                subject = "Unknown"
            
            try:
                body = message.Body
            except:
                body = "Unknown"
            
            # Get recipients
            recipients = []
            try:
                for recipient in message.Recipients:
                    recipient_info = {
                        "name": recipient.Name,
                        "email": recipient.Address,
                        "type": recipient.Type  # 1=To, 2=Cc, 3=Bcc
                    }
                    recipients.append(recipient_info)
            except:
                recipients = []
            
            email_data = {
                "sender": sender,
                "date": date,
                "subject": subject,
                "body": body,
                "recipients": recipients
            }
            
            # Apply filters
            actions = filter_manager.apply_filters(email_data)
            
            # Check if email has iSupport recipient
            has_isupport = False
            for recipient in recipients:
                if recipient["name"] == "iSupport":
                    has_isupport = True
                    break
            
            # Keep email if it has iSupport recipient OR matches any filter
            if has_isupport or actions:
                email_data["filter_actions"] = actions
                
                # Extract MX data if needed
                mx_filter_matched = ("send_mx_alert" in actions or 
                                   "system.MX@hkt-emsconnect.com" in email_data.get('sender', ''))
                if mx_filter_matched:
                    mx_data = filter_manager.extract_mx_data(email_data)
                    if mx_data:
                        email_data.update(mx_data)
                
                # Extract CDC data if needed
                if 'process_cdc' in actions or 'extract_cdc' in actions:
                    cdc_data = filter_manager.extract_cdc_data(email_data)
                    if cdc_data:
                        email_data.update(cdc_data)
                
                all_emails.append(email_data)
            
            processed_count += 1
            
        except Exception as e:
            skipped_error += 1
            if ENABLE_PROGRESS:
                print(f"  Warning: Error processing email {i + 1}: {str(e)[:50]}")
            continue
    
    print(f"\n  Processed {processed_count} emails")
    print(f"  Skipped {skipped_old} old emails (beyond {SCAN_LAST_N_DAYS} days)")
    print(f"  Skipped {skipped_error} emails due to errors")
    print(f"  Kept {len(all_emails)} matching emails")
    
    # Save to JSON
    print(f"\n[4/4] Saving results...")
    try:
        with open("outlook_emails.json", "w", encoding="utf-8") as json_file:
            json.dump(all_emails, json_file, indent=4, ensure_ascii=False)
        print(f"[OK] Successfully saved {len(all_emails)} emails to outlook_emails.json")
    except Exception as e:
        print(f"[ERROR] Error saving to JSON: {e}")
        exit(1)
    
    print("\n" + "=" * 60)
    print("SCAN COMPLETE!")
    print("=" * 60)
    
    # Summary
    if all_emails:
        print(f"\nSummary:")
        print(f"  Total emails saved: {len(all_emails)}")
        
        # Count by type
        mx_count = sum(1 for e in all_emails if 'send_mx_alert' in e.get('filter_actions', []))
        cdc_count = sum(1 for e in all_emails if 'extract_cdc' in e.get('filter_actions', []))
        
        print(f"  MX tickets: {mx_count}")
        print(f"  CDC tickets: {cdc_count}")
        print(f"  Other: {len(all_emails) - mx_count - cdc_count}")
        
        print(f"\nNext step: Run 'python create_tickets.py' to process these emails")
    else:
        print("\n[WARNING] No matching emails found")
        print("  Check your filters in email_filters.json")
