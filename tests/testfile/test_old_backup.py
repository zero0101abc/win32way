# https://youtu.be/50o6RTvYIpY
"""
Automating manual tasks - DigitalSreeni
Reading Outlook inbox (or other mail folder) and extracting 
the required information. 
"""

import win32com.client
import pandas as pd
import json
import re
from typing import Dict, List, Any, Optional

class EmailFilterManager:
    """Manages email filters with Power Automate expression language support"""
    
    def __init__(self):
        self.filters = []
        self.filter_file = "email_filters.json"
        self.load_filters()
    
    def load_filters(self):
        """修正：解決檔案讀取指標問題"""
        try:
            with open(self.filter_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    # 改用 json.loads 處理已經讀取出來的字串 content
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
    
    def edit_filter(self, filter_id: int, **kwargs) -> bool:
        """Edit an existing filter"""
        for filter_data in self.filters:
            if filter_data["id"] == filter_id:
                for key, value in kwargs.items():
                    if key in filter_data:
                        filter_data[key] = value
                filter_data["updated_at"] = str(pd.Timestamp.now())
                self.save_filters()
                return True
        return False
    
    def delete_filter(self, filter_id: int) -> bool:
        """修正：回傳值應反映是否成功刪除"""
        initial_len = len(self.filters)
        self.filters = [f for f in self.filters if f["id"] != filter_id]
        if len(self.filters) < initial_len:
            self.save_filters()
            return True
        return False
    
    def list_filters(self) -> List[Dict]:
        """List all filters"""
        return self.filters
    
    def evaluate_power_automate_expression(self, expression: str, email_data: Dict) -> bool:
        """Evaluate Power Automate expression language against email data"""
        
        # Helper functions that mimic Power Automate
        def equals(value1: Any, value2: Any) -> bool:
            return str(value1).lower() == str(value2).lower()
        
        def contains(text: str, search: str) -> bool:
            return search.lower() in text.lower()
        
        def starts_with(text: str, prefix: str) -> bool:
            return text.lower().startswith(prefix.lower())
        
        def ends_with(text: str, suffix: str) -> bool:
            return text.lower().endswith(suffix.lower())
        
        def concat(*args) -> str:
            return "".join(str(arg) for arg in args)
        
        def substring(text: str, start_index: int, length: Optional[int] = None) -> str:
            if length is None:
                return text[start_index:]
            return text[start_index:start_index + length]
        
        def index_of(text: str, search: str) -> int:
            return text.lower().find(search.lower())
        
        def split(text: str, separator: str) -> List[str]:
            return text.split(separator)
        
        def first(items: List) -> Any:
            return items[0] if items else None
        
        def trim(text: str) -> str:
            return text.strip()
        
        def if_condition(condition: Any, true_value: Any, false_value: Any) -> Any:
            return true_value if condition else false_value
        
        def or_(*args) -> bool:
            return any(args)
        
        def and_(*args) -> bool:
            return all(args)
        
        # Create evaluation context
        context = {
            'equals': equals,
            'contains': contains,
            'startsWith': starts_with,
            'endsWith': ends_with,
            'concat': concat,
            'substring': substring,
            'indexOf': index_of,
            'split': split,
            'first': first,
            'trim': trim,
            'if': if_condition,
            'or': or_,
            'and': and_,
            'sender': email_data.get('sender', ''),
            'subject': email_data.get('subject', ''),
            'body': email_data.get('body', ''),
            'date': email_data.get('date', ''),
            'from_email': email_data.get('sender', ''),
            'recipients': email_data.get('recipients', []),
            'to': [r['email'] for r in email_data.get('recipients', []) if r.get('type') == 1]
        }
        
        try:
            # Replace Power Automate syntax with Python equivalents
            expression = expression.replace('triggerBody()', 'email_data')
            
            # Simple evaluation for common patterns
            if 'equals(' in expression:
                # Extract function calls and evaluate
                pattern = r'(\w+)\(([^)]+)\)'
                matches = re.findall(pattern, expression)
                
                for func_name, args_str in matches:
                    args = [arg.strip().strip('"\'') for arg in args_str.split(',')]
                    
                    # Replace variables with actual values
                    for i, arg in enumerate(args):
                        if arg in ['sender', 'subject', 'body', 'from_email']:
                            args[i] = context.get(arg, '')
                    
                    if func_name == 'equals':
                        return len(args) >= 2 and equals(args[0], args[1])
                    elif func_name == 'contains':
                        return len(args) >= 2 and contains(args[0], args[1])
                    elif func_name == 'startsWith':
                        return len(args) >= 2 and starts_with(args[0], args[1])
            
            # For simple string comparisons
            if '"' in expression or "'" in expression:
                # Extract quoted strings and compare
                quoted_pattern = r'["\']([^"\']+)["\']'
                quoted_values = re.findall(quoted_pattern, expression)
                
                for quoted_value in quoted_values:
                    if quoted_value.lower() in context['subject'].lower():
                        return True
                    if quoted_value.lower() in context['body'].lower():
                        return True
                    if quoted_value.lower() in context['sender'].lower():
                        return True
            
            return False
            
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
            
            # Check TO email filter (only type=1 recipients) - FIXED: check name instead of email
            if filter_data.get('to_email'):
                to_recipients = [r['name'] for r in email_data.get('recipients', []) if r.get('type') == 1]
                if not any(filter_data['to_email'] in to_name for to_name in to_recipients):
                    continue
            
            # If all conditions match, add the action
            if filter_data.get('action'):
                matching_actions.append(filter_data['action'])
        
        return matching_actions
    
    def extract_mx_data(self, email_data: Dict) -> Dict:
        """Extract MX specific data from email body using Power Automate expressions"""
        body_text = email_data.get('body', '')
        
        # Helper functions for Power Automate expressions
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
        
        def if_condition(condition: Any, true_value: Any, false_value: Any) -> Any:
            return true_value if condition else false_value
        
        mx_data = {}
        
        try:
            # Extract ticket_number: trim(first(split(substring(outputs('Text'), add(indexOf(outputs('Text'), 'Number:'),7),200), 'User:')))
            number_index = index_of(body_text, 'Number:')
            if number_index != -1:
                number_text = substring(body_text, number_index + 7, 200)
                number_split = split(number_text, 'User:')
                ticket_number = trim(first(number_split))
                mx_data['ticket_number'] = ticket_number
            
            # Extract shop: trim(first(split(substring(outputs('Text'), add(indexOf(outputs('Text'), 'Location:'),9),200), 'Category:')))
            location_index = index_of(body_text, 'Location:')
            if location_index != -1:
                location_text = substring(body_text, location_index + 9, 200)
                location_split = split(location_text, 'Category:')
                shop_raw = trim(first(location_split))
                
                # Process shop with: concat('MX', if(equals(substring(trim(first(split(outputs('Location'), '-'))), 0, 1), '0'), 
                # substring(trim(first(split(outputs('Location'), '-'))), 1), 
                # trim(first(split(outputs('Location'), '-')))))
                shop_parts = split(shop_raw, '-')
                shop_base = trim(first(shop_parts))
                
                if equals(substring(shop_base, 0, 1), '0'):
                    shop_processed = concat('MX', substring(shop_base, 1))
                else:
                    shop_processed = concat('MX', shop_base)
                
                mx_data['shop'] = shop_processed
            
            # Extract description: trim(first(split(substring(outputs('Text'), add(indexOf(outputs('Text'), 'description:'),13),200), 'You can')))
            description_index = index_of(body_text, 'description:')
            if description_index != -1:
                description_text = substring(body_text, description_index + 13, 200)
                description_split = split(description_text, 'You can')
                description = trim(first(description_split))
                mx_data['description'] = description
                
        except Exception as e:
            print(f"Error extracting MX data: {e}")
        
        return mx_data

    # === 請將此函式貼在 extract_mx_data 函式之後，但在 Class 結束之前 ===
    def extract_cdc_data(self, email_data: Dict) -> Dict:
        """New: Extract CDC specific data with Regex"""
        body = email_data.get('body', '')
        data = {}
        try:
            import re
            # 1. 抓取工單編號 (例如: HK541062)
            ticket_match = re.search(r'Inci\. ID:\s*([A-Z0-9]+)', body)
            if ticket_match:
                data['ticket_number'] = ticket_match.group(1)
            
            # 2. 抓取分店編號 (例如: 333) - 位於 Cust. Name 的括號內
            shop_match = re.search(r'Cust\. Name:\s*.*?\((.*?)\)', body, re.DOTALL)
            if shop_match:
                shop = shop_match.group(1).strip()
                # Add "cdc" prefix if shop doesn't start with "ss"
                if not shop.lower().startswith('ss'):
                    shop = f'cdc{shop}'
                data['shop'] = shop

            # 3. 抓取描述 (Description)
            desc_match = re.search(r'Description:\s*(.*?)\r\n', body, re.DOTALL)
            if desc_match:
                data['description'] = desc_match.group(1).strip()
                
        except Exception as e:
            print(f"Error extracting CDC data: {e}")
        return data
    
# Helper function for backward compatibility
def contains(text: str, search: str) -> bool:
    return search.lower() in text.lower()

# === NEW: Merge functions ===
def load_existing_emails():
    """Load existing outlook_emails.json if it exists"""
    try:
        with open('outlook_emails.json', 'r', encoding='utf-8') as f:
            existing = json.load(f)
            print(f"Loaded {len(existing)} existing emails from outlook_emails.json")
            return existing
    except FileNotFoundError:
        print("No existing outlook_emails.json found - will create new one")
        return []
    except Exception as e:
        print(f"Error loading existing emails: {e}")
        return []

def create_email_key(email_data):
    """Create unique key for email to detect duplicates"""
    # Use combination of date + subject + sender as unique identifier
    date = email_data.get('date', '')
    subject = email_data.get('subject', '')
    sender = email_data.get('sender', '')
    return f"{date}|{subject}|{sender}"

def merge_emails(existing_emails, new_emails):
    """Merge new emails with existing, avoiding duplicates"""
    
    # Create dictionary of existing emails by unique key
    existing_dict = {}
    for email in existing_emails:
        key = create_email_key(email)
        existing_dict[key] = email
    
    added_count = 0
    
    # Add new emails if not duplicate
    for new_email in new_emails:
        key = create_email_key(new_email)
        if key not in existing_dict:
            existing_dict[key] = new_email
            added_count += 1
    
    # Convert back to list
    merged_emails = list(existing_dict.values())
    
    # Sort by date (newest first)
    try:
        merged_emails.sort(key=lambda x: x.get('date', ''), reverse=True)
    except:
        pass
    
    print(f"\nMerge Summary:")
    print(f"  Existing emails: {len(existing_emails)}")
    print(f"  New emails scanned: {len(new_emails)}")
    print(f"  New emails added: {added_count}")
    print(f"  Duplicates skipped: {len(new_emails) - added_count}")
    print(f"  Total emails: {len(merged_emails)}")
    
    return merged_emails

# === NEW: Configuration ===
MAX_EMAILS_TO_SCAN = 200  # Limit to most recent 200 emails
SCAN_LAST_N_DAYS = 30     # Only scan emails from last 30 days

if __name__ == "__main__":
    from datetime import datetime, timedelta
    import time
    
    print("="*60)
    print("EMAIL SCANNER WITH MERGE & LIMIT")
    print("="*60)
    print(f"Configuration:")
    print(f"  - Max emails to scan: {MAX_EMAILS_TO_SCAN}")
    print(f"  - Scan last N days: {SCAN_LAST_N_DAYS}")
    print("="*60)
    
    # Calculate cutoff date
    cutoff_date = datetime.now() - timedelta(days=SCAN_LAST_N_DAYS)
    
    print(f"\n[1/5] Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        print(f"[OK] Connected to Outlook inbox")
    except Exception as e:
        print(f"[ERROR] {e}")
        exit(1)
    
    print(f"\n[2/5] Fetching messages...")
    try:
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by date, newest first
        
        total_in_inbox = messages.Count
        print(f"[OK] Found {total_in_inbox} total emails in inbox")
        print(f"  Limiting scan to {MAX_EMAILS_TO_SCAN} most recent emails")
    except Exception as e:
        print(f"[ERROR] {e}")
        exit(1)

"""
#My test emails look like this...

Subject: From Sreeni

Name: John Smith
Company: ABC international
Email: test1@test.com
Message: Hi there, getting in touch


I sent a few emails in this format so we can extract various details and 
capture them into an Excel document. 

"""

# Initialize filter manager for use in the main loop
filter_manager = EmailFilterManager()

print(f"\n[3/5] Scanning emails...")

#Collect all emails data
all_emails = []
processed_count = 0
skipped_old = 0
skipped_error = 0

for i, message in enumerate(messages):
    # Limit number of emails processed
    if processed_count >= MAX_EMAILS_TO_SCAN:
        print(f"\nReached limit of {MAX_EMAILS_TO_SCAN} emails")
        break
    
    try:
        # Check if email is too old
        try:
            email_date = message.ReceivedTime
            if email_date < cutoff_date:
                skipped_old += 1
                continue
        except:
            pass  # If can't get date, process anyway
        
        # Progress indicator
        if processed_count % 10 == 0 and processed_count > 0:
            print(f"  Processing email {processed_count}...", end='\r')
    except:
        pass
    try:
        sender_name = message.SenderName
    except:
        sender_name = "Unknown"
    
    sender_email = ""
    try:
        # 處理 Exchange 內部帳號 (EX) 與一般 SMTP 帳號
        if message.SenderEmailType == "EX":
            sender_email = message.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            sender_email = message.SenderEmailAddress
    except:
        pass
    
    # 組合完整資訊，讓篩選器可以比對 Email
    sender = f"{sender_name} <{sender_email}>"
    # === 修正結束 ===
    
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
    
    # Get recipients (who gets the mail)
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
    
    # Apply EmailFilterManager filters
    actions = filter_manager.apply_filters(email_data)
    
    # Filter: only keep emails sent to "iSupport" OR emails that match filters
    has_isupport = False
    for recipient in recipients:
        if recipient["name"] == "iSupport":
            has_isupport = True
            break
    
    # Keep email if it has iSupport recipient OR matches any filter
    if has_isupport or actions:
        # Add filter actions to email data
        email_data["filter_actions"] = actions
        
        # Extract MX specific data if MX filter matched or from MX sender
        mx_filter_matched = ("mx" in [action.lower() for action in actions] or 
                           "system.MX@hkt-emsconnect.com" in email_data.get('sender', ''))
        # === 補回這段被遺失的 MX 處理邏輯 ===
        if mx_filter_matched:
            mx_data = filter_manager.extract_mx_data(email_data)
            email_data.update(mx_data)
            if mx_data:
                try:
                    print(f"Extracted MX data: {mx_data}")
                except:
                    print("Extracted MX data")
        # ======================================
        # === 新增：檢查是否需要處理 CDC ===
        if 'process_cdc' in actions or 'extract_cdc' in actions:
            cdc_data = filter_manager.extract_cdc_data(email_data)
            email_data.update(cdc_data)
            if cdc_data:
                try:
                    print(f"Extracted CDC data: {cdc_data}")
                except:
                    print("Extracted CDC data")
        # === 結束 ===
        
        all_emails.append(email_data)
        processed_count += 1
        
        # Print which filters matched
        if actions:
            try:
                print(f"  Matched: {', '.join(actions)} - {subject[:40]}...")
            except:
                print(f"  Matched filters")
    
    except Exception as e:
        skipped_error += 1
        continue

print(f"\n  Scanned {processed_count} emails")
print(f"  Skipped {skipped_old} old emails (beyond {SCAN_LAST_N_DAYS} days)")
print(f"  Skipped {skipped_error} emails due to errors")
print(f"  Kept {len(all_emails)} matching emails")

# Load existing emails and merge
print(f"\n[4/5] Merging with existing data...")
existing_emails = load_existing_emails()
merged_emails = merge_emails(existing_emails, all_emails)

# Save all emails to JSON file
print(f"\n[5/5] Saving to outlook_emails.json...")
with open("outlook_emails.json", "w", encoding="utf-8") as json_file:
    json.dump(merged_emails, json_file, indent=4, ensure_ascii=False)

print(f"[OK] Successfully saved {len(merged_emails)} total emails to outlook_emails.json")

print("\n" + "="*60)
print("SCAN COMPLETE!")
print("="*60)
print(f"Next step: Run 'python create_tickets.py' to process emails into tickets")

# Test MX data extraction
def test_mx_extraction():
    """Test the MX data extraction with sample data"""
    sample_email = {
        "sender": "system.MX@hkt-emsconnect.com",
        "subject": "Incident has been assigned to your group",
        "body": """Incident Details:
Number: 12345
User: John Doe
Location: 0123-HK Main Office
Category: Hardware
description: Server is down and needs immediate attention You can contact support for more details.
""",
        "date": "2026-01-26",
        "recipients": []
    }
    
    fm = EmailFilterManager()
    mx_data = fm.extract_mx_data(sample_email)
    print("Test MX Extraction Results:")
    print(json.dumps(mx_data, indent=2))

# Uncomment to run test
# test_mx_extraction()

# Example usage of EmailFilterManager
def setup_example_filters():
    """Setup example email filters"""
    filter_manager = EmailFilterManager()
    
    # Create filter for specific sender
    filter_manager.create_filter(
        name="System MX Filter",
        from_email="system.MX@hkt-emsconnect.com",
        subject_filter='equals(subject, "Alert")',
        body_filter='contains(body, "error")',
        action="send_to_support_team",
        description="Filter for system alerts with errors"
    )
    
    # Create filter for iSupport emails
    filter_manager.create_filter(
        name="iSupport Filter",
        subject_filter='contains(subject, "Support Request")',
        action="log_to_database",
        description="Filter for support request emails"
    )
    
    # Create filter with complex Power Automate expression
    filter_manager.create_filter(
        name="Urgent Filter",
        subject_filter='or(contains(subject, "Urgent"), contains(subject, "Critical"))',
        body_filter='contains(body, "immediate attention")',
        action="high_priority_notification",
        description="Filter for urgent emails requiring immediate attention"
    )
    
    print("Example filters created successfully!")
    return filter_manager

# Uncomment to setup example filters
# filter_manager = setup_example_filters()

# Example of applying filters to emails
def apply_filters_to_emails():
    """Apply filters to collected emails"""
    filter_manager = EmailFilterManager()
    
    filtered_results = []
    for email in all_emails:
        actions = filter_manager.apply_filters(email)
        if actions:
            filtered_results.append({
                "email": {
                    "sender": email["sender"],
                    "subject": email["subject"],
                    "date": email["date"]
                },
                "actions": actions
            })
    
    # Save filtered results
    with open("filtered_emails.json", "w", encoding="utf-8") as f:
        json.dump(filtered_results, f, indent=4, ensure_ascii=False)
    
    print(f"Applied filters to {len(all_emails)} emails, {len(filtered_results)} matched filters")

# Uncomment to apply filters
# apply_filters_to_emails()