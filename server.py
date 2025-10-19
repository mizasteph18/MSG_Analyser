import http.server
import socketserver
import webbrowser
import os
import json
import time
import glob
from urllib.parse import urlparse, parse_qs
import traceback
from concurrent.futures import ThreadPoolExecutor

class OptimizedMSGAnalyzer:
    def __init__(self, base_folder="msg_files"):
        self.base_folder = base_folder
        self._ensure_base_folder()
        self._cache = {}
        self._cache_time = {}
        self.cache_timeout = 300
        self.message_status = {}
        self.message_comments = {}
        self.thread_pool = ThreadPoolExecutor(max_workers=4)
        self._process_metadata = {}
    
    def _ensure_base_folder(self):
        if not os.path.exists(self.base_folder):
            os.makedirs(self.base_folder)
            print(f"Created base folder: {self.base_folder}")
    
    def get_processes(self):
        """Fast process listing"""
        if self._process_metadata and time.time() - self._process_metadata.get('_timestamp', 0) < 60:
            return self._process_metadata.get('processes', [])
        
        processes = []
        try:
            for item in os.listdir(self.base_folder):
                item_path = os.path.join(self.base_folder, item)
                if os.path.isdir(item_path):
                    msg_files = [f for f in os.listdir(item_path) if f.lower().endswith('.msg')]
                    processes.append({
                        "id": item,
                        "name": item.replace("_", " ").title(),
                        "count": len(msg_files)
                    })
        except Exception as e:
            print(f"Error reading processes: {e}")
        
        if not processes:
            processes = self._create_sample_structure()
        
        self._process_metadata = {
            'processes': processes,
            '_timestamp': time.time()
        }
        
        return processes
    
    def _create_sample_structure(self):
        sample_folders = ["Marketing_Process", "Sales_Process", "HR_Process", "IT_Process"]
        for folder in sample_folders:
            folder_path = os.path.join(self.base_folder, folder)
            os.makedirs(folder_path, exist_ok=True)
        print("Created sample folder structure. Please add .msg files to the subfolders.")
        return [{"id": folder, "name": folder.replace("_", " ").title(), "count": 0} for folder in sample_folders]
    
    def get_messages_for_process_optimized(self, process_id, limit=50, offset=0):
        """Optimized message loading with pagination"""
        cache_key = f"messages_{process_id}_{limit}_{offset}"
        current_time = time.time()
        
        if (cache_key in self._cache and 
            current_time - self._cache_time.get(cache_key, 0) < self.cache_timeout):
            return self._cache[cache_key]
        
        print(f"Loading messages for {process_id} (limit: {limit}, offset: {offset})")
        
        messages = self._load_messages_batch(process_id, limit, offset)
        
        self._cache[cache_key] = messages
        self._cache_time[cache_key] = current_time
        
        return messages
    
    def _load_messages_batch(self, process_id, limit, offset):
        """Load only a batch of messages"""
        messages = []
        process_path = os.path.join(self.base_folder, process_id)
        
        if not os.path.exists(process_path):
            return {
                "messages": messages,
                "total_count": 0,
                "has_more": False,
                "offset": offset,
                "limit": limit
            }
        
        try:
            msg_files = glob.glob(os.path.join(process_path, "*.msg"))
            total_count = len(msg_files)
            msg_files.sort(key=os.path.getmtime, reverse=True)
            
            batch_files = msg_files[offset:offset + limit]
            
            print(f"Processing {len(batch_files)} files out of {total_count} total")
            
            for i, msg_file in enumerate(batch_files):
                try:
                    filename = os.path.basename(msg_file)
                    message_id = f"{process_id}_{os.path.splitext(filename)[0]}"
                    
                    # Create realistic sample data
                    message_data = self._create_sample_message_data(msg_file, process_id, message_id, i)
                    messages.append(message_data)
                    
                except Exception as e:
                    print(f"Error with {msg_file}: {e}")
                    continue
            
            messages.sort(key=lambda x: x.get('date', ''), reverse=True)
            
        except Exception as e:
            print(f"Error reading messages for process {process_id}: {e}")
            traceback.print_exc()
        
        return {
            "messages": messages,
            "total_count": total_count,
            "has_more": (offset + limit) < total_count,
            "offset": offset,
            "limit": limit
        }
    
    def _create_sample_message_data(self, msg_file, process_id, message_id, index):
        """Create realistic sample message data"""
        subjects = [
            "Quarterly Business Review Meeting",
            "Project Alpha Status Update", 
            "Client Presentation Materials",
            "Team Building Event Planning",
            "Budget Approval Request Q3",
            "New Feature Development Discussion",
            "Security Audit Findings",
            "Weekly Sales Report",
            "Product Launch Timeline",
            "Customer Feedback Analysis"
        ]
        
        senders = [
            "john.doe@company.com",
            "sarah.smith@company.com", 
            "mike.johnson@company.com",
            "lisa.wang@company.com",
            "david.brown@company.com"
        ]
        
        bodies = [
            "Hi team, I wanted to follow up on our discussion from yesterday regarding the upcoming project deadlines.",
            "Please find attached the documents for review. Let me know if you have any questions.",
            "We need to schedule a meeting to discuss the implementation plan for the new features.",
            "The client has provided feedback on the initial proposal. Let's discuss how to address their concerns.",
            "Here is the updated timeline for the project. Please review and provide your input.",
            "I've completed the initial analysis and have some findings to share with the team.",
            "Following up on our conversation, here are the action items we agreed upon.",
            "The meeting minutes from today's discussion are attached for your reference.",
            "We need to coordinate with other teams to ensure a smooth rollout process.",
            "Please review the attached document and provide your feedback by end of day."
        ]
        
        import random
        from datetime import datetime, timedelta
        
        # Generate random date within last 30 days
        random_days = random.randint(0, 30)
        random_hours = random.randint(0, 23)
        random_minutes = random.randint(0, 59)
        message_date = datetime.now() - timedelta(days=random_days, hours=random_hours, minutes=random_minutes)
        
        subject = random.choice(subjects)
        if index % 3 == 0:
            subject = f"RE: {subject}"
        elif index % 5 == 0:
            subject = f"FW: {subject}"
        
        return {
            "id": message_id,
            "subject": subject,
            "from": random.choice(senders),
            "to": "team@company.com, managers@company.com",
            "date": message_date.isoformat(),
            "body": random.choice(bodies),
            "body_preview": random.choice(bodies)[:100] + "...",
            "body_type": "text",
            "status": self.message_status.get(message_id, "untagged"),
            "threadId": f"thread_{hash(subject) % 1000}",
            "filename": os.path.basename(msg_file),
            "attachments": [{"name": "document.pdf", "url": f"/api/attachment/{message_id}/0", "type": "PDF"}] if index % 4 == 0 else [],
            "containsThread": index % 3 == 0,
            "comments": self.message_comments.get(message_id, []),
            "has_attachments": index % 4 == 0,
            "needs_full_parse": True
        }
    
    def get_message_full_content(self, process_id, message_id):
        """Load full content only when needed"""
        cache_key = f"full_{message_id}"
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        # Create enhanced full content
        full_message = self._create_sample_message_data(
            f"{message_id.replace(process_id + '_', '')}.msg", 
            process_id, 
            message_id, 
            hash(message_id) % 10
        )
        
        # Enhance with more details for full view
        full_message["to"] = "john.doe@company.com, jane.smith@company.com, mike.johnson@company.com, sarah.wilson@company.com"
        full_message["body"] = f"""Dear Team,

This is the complete message content for {full_message['subject']}.

I wanted to provide you with a comprehensive update on the current project status. We have made significant progress in the following areas:

1. Development milestones achieved
2. Client feedback incorporated
3. Next steps identified

Please review the attached documents and let me know if you have any questions or concerns. We should schedule a follow-up meeting to discuss the implementation plan.

Best regards,
Team Lead

Original Message:
This would be the actual content extracted from the .msg file in a production environment."""
        
        full_message["body_type"] = "text"
        full_message["attachments"] = [
            {"name": "project_document.pdf", "url": f"/api/attachment/{message_id}/0", "type": "PDF"},
            {"name": "meeting_notes.docx", "url": f"/api/attachment/{message_id}/1", "type": "Word"},
            {"name": "data_analysis.xlsx", "url": f"/api/attachment/{message_id}/2", "type": "Excel"}
        ]
        
        # Cache full content
        self._cache[cache_key] = full_message
        self._cache_time[cache_key] = time.time()
        
        return full_message
    
    def get_attachment(self, process_id, message_id, attachment_index):
        """Handle attachment download - returns sample file"""
        try:
            # In a real implementation, this would extract the actual attachment
            # For now, return a sample response
            return {"message": f"Attachment {attachment_index} would be downloaded here"}
        except Exception as e:
            print(f"Error getting attachment: {e}")
            raise
    
    def update_message_status(self, process_id, message_id, status):
        try:
            self.message_status[message_id] = status
            # Invalidate relevant caches
            for key in list(self._cache.keys()):
                if f"messages_{process_id}" in key or f"full_{message_id}" in key:
                    del self._cache[key]
            return True
        except Exception as e:
            print(f"Error updating status: {e}")
            return False
    
    def add_comment_to_message(self, process_id, message_id, comment_data):
        try:
            if message_id not in self.message_comments:
                self.message_comments[message_id] = []
            
            comment_data["date"] = time.strftime("%Y-%m-%dT%H:%M:%S")
            self.message_comments[message_id].append(comment_data)
            
            # Invalidate relevant caches
            for key in list(self._cache.keys()):
                if f"messages_{process_id}" in key or f"full_{message_id}" in key:
                    del self._cache[key]
                    
            return True
        except Exception as e:
            print(f"Error adding comment: {e}")
            return False
    
    def clear_cache(self):
        self._cache.clear()
        self._cache_time.clear()

# Create analyzer
analyzer = OptimizedMSGAnalyzer()

class MSGHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith('/api/'):
            self.handle_api_request()
        else:
            self.serve_html()
    
    def do_POST(self):
        if self.path.startswith('/api/'):
            self.handle_api_post()
        else:
            self.send_error(404, "Endpoint not found")
    
    def serve_html(self):
        try:
            if self.path == '/' or self.path.endswith('.html'):
                with open('index.html', 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                self.send_response(200)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(html_content.encode('utf-8'))
            else:
                super().do_GET()
                
        except Exception as e:
            print(f"Error serving HTML: {e}")
            self.send_error(500, f"Error serving page: {str(e)}")
    
    def handle_api_request(self):
        try:
            if self.path == '/api/processes':
                self.handle_processes()
            elif self.path.startswith('/api/messages/'):
                self.handle_messages()
            elif self.path.startswith('/api/message/'):
                self.handle_single_message()
            elif self.path.startswith('/api/attachment/'):
                self.handle_attachment()
            elif self.path == '/api/health':
                self.handle_health()
            else:
                self.send_error(404, f"API endpoint not found: {self.path}")
        except Exception as e:
            print(f"API error in handle_api_request: {e}")
            traceback.print_exc()
            self.send_error(500, f"Server error: {str(e)}")
    
    def handle_api_post(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            if self.path.startswith('/api/message/'):
                path_parts = self.path.split('/')
                if len(path_parts) >= 5:
                    process_id = path_parts[3]
                    message_id = path_parts[4]
                    action = path_parts[5] if len(path_parts) > 5 else None
                    
                    if action == 'status':
                        self.handle_update_status(process_id, message_id, post_data)
                    elif action == 'comment':
                        self.handle_add_comment(process_id, message_id, post_data)
                    else:
                        self.send_error(404, "Action not found")
                else:
                    self.send_error(400, "Invalid URL format")
            elif self.path == '/api/refresh-cache':
                self.handle_refresh_cache()
            else:
                self.send_error(404, "API endpoint not found")
        except Exception as e:
            print(f"API POST error: {e}")
            traceback.print_exc()
            self.send_error(500, f"Server error: {str(e)}")
    
    def handle_processes(self):
        try:
            processes = analyzer.get_processes()
            self.send_json_response(processes)
        except Exception as e:
            print(f"Error in handle_processes: {e}")
            self.send_error(500, f"Error retrieving processes: {str(e)}")
    
    def handle_messages(self):
        path_parts = self.path.split('/')
        process_id = path_parts[3] if len(path_parts) > 3 else None
        
        if not process_id:
            self.send_error(400, "Process ID required")
            return
        
        query = urlparse(self.path).query
        params = parse_qs(query)
        limit = int(params.get('limit', [50])[0])
        offset = int(params.get('offset', [0])[0])
        
        try:
            messages_data = analyzer.get_messages_for_process_optimized(process_id, limit, offset)
            self.send_json_response(messages_data)
        except Exception as e:
            print(f"Error in handle_messages: {e}")
            self.send_error(500, f"Error retrieving messages: {str(e)}")
    
    def handle_single_message(self):
        path_parts = self.path.split('/')
        if len(path_parts) < 5:
            self.send_error(400, "Invalid message URL")
            return
        
        process_id = path_parts[3]
        message_id = path_parts[4]
        
        try:
            full_message = analyzer.get_message_full_content(process_id, message_id)
            self.send_json_response(full_message)
        except Exception as e:
            print(f"Error loading full message: {e}")
            self.send_error(500, f"Error loading message: {str(e)}")
    
    def handle_attachment(self):
        path_parts = self.path.split('/')
        if len(path_parts) < 6:
            self.send_error(400, "Invalid attachment URL")
            return
        
        process_id = path_parts[3]
        message_id = path_parts[4]
        try:
            attachment_index = int(path_parts[5])
        except ValueError:
            self.send_error(400, "Invalid attachment index")
            return
        
        try:
            attachment_data = analyzer.get_attachment(process_id, message_id, attachment_index)
            self.send_json_response(attachment_data)
        except Exception as e:
            print(f"Error handling attachment: {e}")
            self.send_error(500, f"Error with attachment: {str(e)}")
    
    def handle_update_status(self, process_id, message_id, post_data):
        try:
            data = json.loads(post_data.decode('utf-8'))
            status = data.get('status')
            
            if status not in ['keep', 'review', 'untagged']:
                self.send_error(400, "Invalid status")
                return
            
            success = analyzer.update_message_status(process_id, message_id, status)
            
            if success:
                self.send_json_response({"message": "Status updated successfully"})
            else:
                self.send_error(500, "Failed to update status")
        except Exception as e:
            print(f"Error updating status: {e}")
            self.send_error(500, f"Error updating status: {str(e)}")
    
    def handle_add_comment(self, process_id, message_id, post_data):
        try:
            data = json.loads(post_data.decode('utf-8'))
            success = analyzer.add_comment_to_message(process_id, message_id, data)
            
            if success:
                self.send_json_response({"message": "Comment added successfully"})
            else:
                self.send_error(500, "Failed to add comment")
        except Exception as e:
            print(f"Error adding comment: {e}")
            self.send_error(500, f"Error adding comment: {str(e)}")
    
    def handle_refresh_cache(self):
        analyzer.clear_cache()
        self.send_json_response({"message": "Cache refreshed successfully"})
    
    def handle_health(self):
        self.send_json_response({
            "status": "healthy", 
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "version": "2.0-optimized"
        })
    
    def send_json_response(self, data):
        try:
            response_data = json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.send_header('Content-length', str(len(response_data)))
            self.end_headers()
            self.wfile.write(response_data)
        except Exception as e:
            print(f"Error sending JSON response: {e}")
            self.send_error(500, "Error creating response")
    
    def log_message(self, format, *args):
        print(f"{self.client_address[0]} - {format % args}")

def start_server(port=8000):
    print("🚀 Starting OPTIMIZED MSG Analyzer Server...")
    print(f"📁 Working directory: {os.getcwd()}")
    
    if not os.path.exists('index.html'):
        print("❌ ERROR: index.html not found!")
        print("Please make sure index.html is in the same directory as server.py")
        return
    
    try:
        with socketserver.TCPServer(("", port), MSGHandler) as httpd:
            print(f"✅ Server started successfully on http://localhost:{port}")
            print("⚡ Features: Lazy Loading + Pagination + Real-time Search")
            print("📂 Message folder: msg_files/")
            print("⏹️  Press Ctrl+C to stop")
            
            try:
                webbrowser.open(f'http://localhost:{port}')
            except:
                print("⚠️  Could not open browser automatically")
            
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\n🛑 Server stopped")
                
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"❌ Port {port} is busy! Try: python server.py {port+1}")
        else:
            print(f"❌ Server error: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == '__main__':
    import sys
    port = 8000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print("Invalid port. Using 8000")
    
    start_server(port)