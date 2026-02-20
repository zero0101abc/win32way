import win32com.client
import json
import re
from datetime import datetime
import time

# === SETTINGS ===
MAX_EMAILS = 100       # Scan last 100 emails
TIMEOUT_SECONDS = 30   # Stop after 30 seconds

class SafeEmailScanner:
    def __init__(self):
        self.load_filters()

    def load_filters(self):
        try:
            with open("email_filters.json", "r", encoding="utf-8") as f:
                self.filters = json.load(f)
        except:
            self.filters = []

    def apply_filters(self, email_data):
        actions = []
        sender = email_data.get('sender', '').lower()
        subject = email_data.get('subject', '').lower()
        body = email_data.get('body', '').lower()

        for f in self.filters:
            if not f.get('enabled', True): continue
            
            match = True
            # Check Sender (Safe)
            if f.get('from_email') and f['from_email'].lower() not in sender:
                match = False
            
            # Check Subject (Safe)
            if match and f.get('subject_filter') and f['subject_filter'].lower() not in subject:
                match = False
                
            # Check Body (Safe)
            if match and f.get('body_filter') and f['body_filter'].lower() not in body:
                match = False

            if match:
                actions.append(f['action'])
                
        return actions

    def extract_mx_data(self, body):
        data = {}
        try:
            if 'Number:' in body:
                data['ticket_number'] = body.split('Number:')[1].split()[0].strip()
            if 'Shop:' in body:
                data['shop'] = body.split('Shop:')[1].split()[0].strip()
            if 'Short description:' in body:
                data['description'] = body.split('Short description:')[1].split('You can view')[0].strip()
        except: pass
        return data

    def extract_cdc_data(self, body):
        data = {}
        try:
            ticket_match = re.search(r'Inci\. ID:\s+([A-Z0-9]+)', body)
            if ticket_match: data['ticket_number'] = ticket_match.group(1)
            
            shop_match = re.search(r'Cust\. Name:\s*.*?\((.*?)\)', body, re.DOTALL)
            if shop_match: data['shop'] = shop_match.group(1)

            desc_match = re.search(r'Description:\s+([^\r\n]+)', body)
            if desc_match: data['description'] = desc_match.group(1).strip()
        except: pass
        return data

if __name__ == "__main__":
    print(f"--- Starting Safe Scan (Limit: {MAX_EMAILS} emails) ---")
    start_time = time.time()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        scanner = SafeEmailScanner()
        results = []
        
        # Iterate with index to stop early
        for i, message in enumerate(messages):
            # 1. Check Limits
            if i >= MAX_EMAILS:
                print(f"\nLimit reached ({MAX_EMAILS} emails). Stopping.")
                break
            if time.time() - start_time > TIMEOUT_SECONDS:
                print(f"\nTimeout reached ({TIMEOUT_SECONDS}s). Stopping.")
                break
            
            try:
                # 2. Access SAFE properties only
                # STRICT SAFETY MODE: Use SenderName ONLY by default
                try:
                    sender = message.SenderName
                except:
                    sender = "Unknown"
                    
                # Only attempt to get Email Address if NOT Exchange (EX) to avoid prompts
                # If you get warnings, the line below is usually the culprit.
                # We wrap it in a try-block that specifically fails silently.
                try:
                    print('fk')
                except:
                    pass

                subject = message.Subject
                body = message.Body
                received_time = str(message.ReceivedTime)

                # 3. Create Data Object
                email_data = {
                    "subject": subject,
                    "sender": sender,
                    "body": body,
                    "received_time": received_time
                }

                # 4. Apply Filters
                actions = scanner.apply_filters(email_data)
                
                # 5. Extract Data if Matched
                if actions:
                    email_data['actions'] = actions
                    print(f". Matched: {subject[:30]}...")
                    
                    if 'send_mx_alert' in actions or 'mx' in str(actions).lower():
                        mx_data = scanner.extract_mx_data(body)
                        email_data.update(mx_data)
                    
                    if 'extract_cdc' in actions or 'process_cdc' in actions:
                        cdc_data = scanner.extract_cdc_data(body)
                        email_data.update(cdc_data)
                    
                    results.append(email_data)
                
                # Progress indicator every 10 emails
                if i % 10 == 0:
                    print(f"Scanned {i}...", end="\r")

            except Exception as e:
                # Skip errors on individual emails (e.g. encryption)
                continue

        # 6. Save Results
        print(f"\nSaving {len(results)} matches to outlook_emails.json...")
        with open("outlook_emails.json", "w", encoding="utf-8") as f:
            json.dump(results, f, indent=4, ensure_ascii=False)
            
        print("Done.")

    except Exception as e:
        print(f"Critical Error: {e}")