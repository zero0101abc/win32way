import http.server
import socketserver
import json
import subprocess
import os
import sys
import shutil
from datetime import datetime

# Port number (must match the fetch URL in your dashboard.html)
PORT = 8000

# Settings storage
SETTINGS = {
    "auto_refresh": False,
    "refresh_interval": 60,
    "default_status": "in progress",
    "default_handler": "USE_MISSING",
    "last_scan": None,
    "auto_refresh_after_scan": False
}

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        print(f"GET request: {self.path}")
        # Handle special scan routes
        if self.path == '/run-scan':
            self.run_scan("backend/test.py")
            return
        elif self.path == '/run-scan-all':
            self.run_scan("backend/test_all.py")
            return
        elif self.path.startswith('/api/'):
            print(f"API request: {self.path}")
            if self.path == '/api/filters':
                self.handle_get_filters()
            elif self.path == '/api/stats':
                self.handle_get_stats()
            elif self.path == '/api/settings':
                self.handle_get_settings()
            else:
                self.send_error(404, "API endpoint not found")
            return
        
        # For all other requests, serve files (html, json, etc.) normally
        super().do_GET()

    def do_OPTIONS(self):
        # Handle CORS preflight for browser security
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header("Access-Control-Allow-Headers", "X-Requested-With, Content-type")
        self.end_headers()

    def do_POST(self):
        if self.path == '/update-ticket':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            updated_ticket = json.loads(post_data.decode('utf-8'))
            
            # 1. Load existing ticket.json
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            ticket_path = os.path.join(project_root, "database", "ticket.json")
            
            with open(ticket_path, 'r', encoding='utf-8') as f:
                tickets = json.load(f)
                
            # 2. Find and update the specific ticket (by id or ticket_number)
            ticket_id = updated_ticket.get('id') or updated_ticket.get('ticket_number')
            for i, ticket in enumerate(tickets):
                ticket_key = ticket.get('id') or ticket.get('ticket_number')
                if ticket_key == ticket_id:
                    # Update all editable fields
                    if 'solution' in updated_ticket:
                        tickets[i]['solution'] = updated_ticket['solution']
                    if 'resolve_time' in updated_ticket:
                        tickets[i]['resolve_time'] = updated_ticket['resolve_time']
                    if 'ph_rm_os' in updated_ticket:
                        tickets[i]['ph_rm_os'] = updated_ticket['ph_rm_os']
                    if 'fu_action' in updated_ticket:
                        tickets[i]['fu_action'] = updated_ticket['fu_action']
                    if 'problem' in updated_ticket:
                        tickets[i]['problem'] = updated_ticket['problem']
                    if 'handled_by' in updated_ticket:
                        tickets[i]['handled_by'] = updated_ticket['handled_by']
                    if 'status' in updated_ticket:
                        tickets[i]['status'] = updated_ticket['status']
                    # Update ticket_number if it was empty before
                    if 'ticket_number' in updated_ticket:
                        tickets[i]['ticket_number'] = updated_ticket['ticket_number']
                    break
                    
            # 3. Save back to file
            with open(ticket_path, 'w', encoding='utf-8') as f:
                json.dump(tickets, f, indent=2, ensure_ascii=False)
                
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))
        elif self.path == '/add-ticket':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            new_ticket = json.loads(post_data.decode('utf-8'))
            
            # 1. Load existing ticket.json
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            ticket_path = os.path.join(project_root, "database", "ticket.json")
            
            with open(ticket_path, 'r', encoding='utf-8') as f:
                tickets = json.load(f)
                
            # 2. Create new ticket with default values
            ticket_number = new_ticket.get('ticket_number', '').strip()
            # Ticket number is now optional (can be null/empty)
            # Generate a unique ID if not provided
            if not ticket_number:
                import datetime
                ticket_number = "TEMP-" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            
            # Check if ticket already exists
            for ticket in tickets:
                if ticket['ticket_number'] == ticket_number:
                    self.send_response(400)
                    self.send_header('Content-type', 'application/json')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": f"Ticket {ticket_number} already exists"}).encode('utf-8'))
                    return
            
            # Create new ticket object
            import datetime
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            
            created_ticket = {
                "ticket_number": ticket_number,
                "shop": new_ticket.get('shop', '').strip() or '',
                "description": new_ticket.get('description', '').strip() or '',
                "date": new_ticket.get('date', '').strip() or now,
                "problem": new_ticket.get('problem', '').strip() or '',
                "resolve_time": new_ticket.get('resolve_time', '').strip() or '',
                "ph_rm_os": new_ticket.get('ph_rm_os', '').strip() or '',
                "solution": new_ticket.get('solution', '').strip() or '',
                "fu_action": new_ticket.get('fu_action', '').strip() or '',
                "handled_by": new_ticket.get('handled_by', '').strip() or '',
                "status": new_ticket.get('status', 'in progress').strip() or 'in progress'
            }
            
            # 3. Add to tickets array
            tickets.append(created_ticket)
            
            # 4. Save back to file
            with open(ticket_path, 'w', encoding='utf-8') as f:
                json.dump(tickets, f, indent=2, ensure_ascii=False)
                
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "success", "ticket": created_ticket}).encode('utf-8'))
        
        # === API ENDPOINTS ===
        elif self.path == '/api/filters' and self.command == 'GET':
            self.handle_get_filters()
        elif self.path == '/api/filters' and self.command == 'POST':
            self.handle_add_filter()
        elif self.path.startswith('/api/filters/') and self.command == 'PUT':
            filter_id = int(self.path.split('/')[-1])
            self.handle_update_filter(filter_id)
        elif self.path.startswith('/api/filters/') and self.command == 'DELETE':
            filter_id = int(self.path.split('/')[-1])
            self.handle_delete_filter(filter_id)
        elif self.path == '/api/scan' and self.command == 'POST':
            self.handle_scan()
        elif self.path == '/api/stats' and self.command == 'GET':
            self.handle_get_stats()
        elif self.path == '/api/settings' and self.command == 'GET':
            self.handle_get_settings()
        elif self.path == '/api/settings' and self.command == 'POST':
            self.handle_update_settings()
        elif self.path == '/api/backup' and self.command == 'POST':
            self.handle_backup()
        elif self.path == '/api/clear-tickets' and self.command == 'POST':
            self.handle_clear_tickets()
        elif self.path == '/api/clear-emails' and self.command == 'POST':
            self.handle_clear_emails()
        else:
            self.send_response(404)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "error", "message": "Endpoint not found"}).encode('utf-8'))

    def handle_get_filters(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        filters_path = os.path.join(project_root, "database", "email_filters.json")
        try:
            with open(filters_path, 'r', encoding='utf-8') as f:
                filters = json.load(f)
            self.send_response(200)
        except:
            filters = []
            self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(filters).encode('utf-8'))

    def handle_add_filter(self):
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        data = json.loads(post_data.decode('utf-8'))
        
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        filters_path = os.path.join(project_root, "database", "email_filters.json")
        
        with open(filters_path, 'r', encoding='utf-8') as f:
            filters = json.load(f)
        
        new_id = max([f.get('id', 0) for f in filters], default=0) + 1
        new_filter = {
            "id": new_id,
            "name": data.get('name', ''),
            "from_email": data.get('from_email', ''),
            "subject_filter": data.get('subject_filter', ''),
            "body_filter": data.get('body_filter', ''),
            "to_email": data.get('to_email', ''),
            "action": data.get('action', ''),
            "description": data.get('description', ''),
            "enabled": data.get('enabled', True),
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
        }
        filters.append(new_filter)
        
        with open(filters_path, 'w', encoding='utf-8') as f:
            json.dump(filters, f, indent=2, ensure_ascii=False)
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success", "filter": new_filter}).encode('utf-8'))

    def handle_update_filter(self, filter_id):
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        data = json.loads(post_data.decode('utf-8'))
        
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        filters_path = os.path.join(project_root, "database", "email_filters.json")
        
        with open(filters_path, 'r', encoding='utf-8') as f:
            filters = json.load(f)
        
        updated = False
        for f in filters:
            if f.get('id') == filter_id:
                for key in ['name', 'from_email', 'subject_filter', 'body_filter', 'to_email', 'action', 'description', 'enabled']:
                    if key in data:
                        f[key] = data[key]
                updated = True
                break
        
        if updated:
            with open(filters_path, 'w', encoding='utf-8') as f:
                json.dump(filters, f, indent=2, ensure_ascii=False)
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success" if updated else "error"}).encode('utf-8'))

    def handle_delete_filter(self, filter_id):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        filters_path = os.path.join(project_root, "database", "email_filters.json")
        
        with open(filters_path, 'r', encoding='utf-8') as f:
            filters = json.load(f)
        
        filters = [f for f in filters if f.get('id') != filter_id]
        
        with open(filters_path, 'w', encoding='utf-8') as f:
            json.dump(filters, f, indent=2, ensure_ascii=False)
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))

    def handle_scan(self):
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        data = json.loads(post_data.decode('utf-8'))
        mode = data.get('mode', 'quick')
        
        script_name = "backend/test_all.py" if mode == 'full' else "backend/test.py"
        self.run_scan(script_name)

    def handle_get_stats(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        ticket_path = os.path.join(project_root, "database", "ticket.json")
        emails_path = os.path.join(project_root, "database", "outlook_emails.json")
        
        ticket_count = 0
        email_count = 0
        
        if os.path.exists(ticket_path):
            with open(ticket_path, 'r', encoding='utf-8') as f:
                tickets = json.load(f)
                ticket_count = len(tickets)
        
        if os.path.exists(emails_path):
            with open(emails_path, 'r', encoding='utf-8') as f:
                emails = json.load(f)
                email_count = len(emails)
        
        stats = {
            "ticket_count": ticket_count,
            "email_count": email_count,
            "last_scan": SETTINGS.get('last_scan'),
            "server_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(stats).encode('utf-8'))

    def handle_get_settings(self):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(SETTINGS).encode('utf-8'))

    def handle_update_settings(self):
        global SETTINGS
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        data = json.loads(post_data.decode('utf-8'))
        
        for key in data:
            SETTINGS[key] = data[key]
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))

    def handle_backup(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = os.path.join(project_root, "backup", timestamp)
        
        os.makedirs(backup_dir, exist_ok=True)
        
        files_to_backup = [
            ("database/ticket.json", "tickets.json"),
            ("database/email_filters.json", "filters.json"),
            ("database/outlook_emails.json", "emails.json")
        ]
        
        for src, dst in files_to_backup:
            src_path = os.path.join(project_root, src)
            if os.path.exists(src_path):
                shutil.copy2(src_path, os.path.join(backup_dir, dst))
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success", "backup_dir": backup_dir}).encode('utf-8'))

    def handle_clear_tickets(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        ticket_path = os.path.join(project_root, "database", "ticket.json")
        
        with open(ticket_path, 'w', encoding='utf-8') as f:
            json.dump([], f, indent=2, ensure_ascii=False)
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))

    def handle_clear_emails(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        emails_path = os.path.join(project_root, "database", "outlook_emails.json")
        
        with open(emails_path, 'w', encoding='utf-8') as f:
            json.dump([], f, indent=2, ensure_ascii=False)
        
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "success"}).encode('utf-8'))

    def run_scan(self, script_name="test_quick.py"):
        print(f"Received request to scan emails using {script_name}...")
        try:
            # Get project root directory
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            # 1. Execute script to fetch emails from Outlook -> outlook_emails.json
            print(f"Running {script_name}...")
            timeout = 300 if "test_all" in script_name else 180
            script_path = os.path.join(project_root, script_name)
            scan_result = subprocess.run(
                [sys.executable, script_path], 
                capture_output=True, 
                timeout=timeout,
                cwd=project_root  # Run from project root
            )
            
            # Handle output with proper encoding
            stdout_text = scan_result.stdout.decode('utf-8', errors='replace') if scan_result.stdout else ""
            stderr_text = scan_result.stderr.decode('utf-8', errors='replace') if scan_result.stderr else ""
            
            if stdout_text: print(f"{script_name} Output:", stdout_text)
            if stderr_text: print(f"{script_name} Errors:", stderr_text)

            if scan_result.returncode != 0:
                raise Exception(f"{script_name} failed with code {scan_result.returncode}")

            # 2. Execute create_tickets.py to process emails -> ticket.json
            print("Running create_tickets.py...")
            create_tickets_path = os.path.join(project_root, "backend", "create_tickets.py")
            ticket_result = subprocess.run(
                [sys.executable, create_tickets_path],
                capture_output=True,
                cwd=project_root  # Run from project root
            )
            
            # Handle output with proper encoding
            ticket_stdout = ticket_result.stdout.decode('utf-8', errors='replace') if ticket_result.stdout else ""
            ticket_stderr = ticket_result.stderr.decode('utf-8', errors='replace') if ticket_result.stderr else ""

            if ticket_stdout: print("create_tickets.py Output:", ticket_stdout)
            if ticket_stderr: print("create_tickets.py Errors:", ticket_stderr)

            if ticket_result.returncode != 0:
                raise Exception(f"create_tickets.py failed with code {ticket_result.returncode}")

            SETTINGS['last_scan'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 3. Read and return ticket.json (NOT outlook_emails.json)
            ticket_path = os.path.join(project_root, "database", "ticket.json")
            if os.path.exists(ticket_path):
                with open(ticket_path, "r", encoding='utf-8') as f:
                    data = json.load(f)
                
                response = {
                    "status": "success", 
                    "message": f"Scan and ticket creation complete using {script_name}",
                    "data": data,
                    "source": "ticket.json"
                }
            else:
                response = {"status": "error", "message": "ticket.json not found after processing."}

        except Exception as e:
            print(f"Server Error: {str(e)}")
            response = {"status": "error", "message": str(e)}

        # Send JSON response back to the frontend
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(response).encode('utf-8'))

if __name__ == "__main__":
    # Prevent "Address already in use" errors on restart
    socketserver.TCPServer.allow_reuse_address = True
    
    print(f"Starting server at http://localhost:{PORT}")
    print(f"1. Run 'python -m backend.server' from the project root (win32way/)")
    print(f"2. Open http://localhost:{PORT}/frontend/dashboard.html in your browser.")
    print(f"3. Quick scanner: scans 50 most recent emails in <30 seconds")
    print(f"4. Full scanner: scans ALL emails (may take several minutes)")
    
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")