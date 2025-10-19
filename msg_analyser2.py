import http.server
import socketserver
import threading
import webbrowser
import os
import json
import extract_msg
from datetime import datetime
import tempfile
import shutil
import glob
from urllib.parse import urlparse, parse_qs
import html
import time
from functools import lru_cache
import traceback

class MSGAnalyzer:
    def __init__(self, base_folder="msg_files"):
        self.base_folder = base_folder
        self._ensure_base_folder()
        self._cache = {}
        self._cache_time = {}
        self.cache_timeout = 300  # 5 minutes cache
        self.message_status = {}
        self.message_comments = {}
    
    def _ensure_base_folder(self):
        if not os.path.exists(self.base_folder):
            os.makedirs(self.base_folder)
            print(f"Created base folder: {self.base_folder}")
    
    def get_processes(self):
        processes = []
        try:
            if not os.path.exists(self.base_folder):
                self._ensure_base_folder()
                return self._create_sample_structure()
                
            for item in os.listdir(self.base_folder):
                item_path = os.path.join(self.base_folder, item)
                if os.path.isdir(item_path):
                    msg_files = glob.glob(os.path.join(item_path, "*.msg"))
                    processes.append({
                        "id": item,
                        "name": item.replace("_", " ").title(),
                        "count": len(msg_files)
                    })
        except Exception as e:
            print(f"Error reading processes: {e}")
            traceback.print_exc()
        
        if not processes:
            return self._create_sample_structure()
        
        return processes
    
    def _create_sample_structure(self):
        sample_folders = ["Marketing_Process", "Sales_Process", "HR_Process", "IT_Process"]
        for folder in sample_folders:
            folder_path = os.path.join(self.base_folder, folder)
            os.makedirs(folder_path, exist_ok=True)
            # Create a README file in each folder
            with open(os.path.join(folder_path, "README.txt"), "w") as f:
                f.write(f"Add .msg files to this folder ({folder}) for them to appear in the analyzer.")
        
        print("Created sample folder structure. Please add .msg files to the subfolders.")
        return [{"id": folder, "name": folder.replace("_", " ").title(), "count": 0} for folder in sample_folders]
    
    def get_messages_for_process_cached(self, process_id):
        cache_key = f"messages_{process_id}"
        current_time = time.time()
        
        if (cache_key in self._cache and 
            cache_key in self._cache_time and
            current_time - self._cache_time[cache_key] < self.cache_timeout):
            return self._cache[cache_key]
        
        messages = self.get_messages_for_process(process_id)
        self._cache[cache_key] = messages
        self._cache_time[cache_key] = current_time
        
        return messages
    
    def get_messages_for_process(self, process_id):
        messages = []
        process_path = os.path.join(self.base_folder, process_id)
        
        if not os.path.exists(process_path):
            print(f"Process path does not exist: {process_path}")
            return messages
        
        try:
            msg_files = glob.glob(os.path.join(process_path, "*.msg"))
            print(f"Found {len(msg_files)} .msg files in {process_path}")
            
            for msg_file in msg_files:
                try:
                    message_data = self._parse_msg_file(msg_file, process_id)
                    if message_data:
                        # Apply saved status and comments
                        message_id = message_data["id"]
                        if message_id in self.message_status:
                            message_data["status"] = self.message_status[message_id]
                        if message_id in self.message_comments:
                            message_data["comments"] = self.message_comments[message_id]
                        
                        messages.append(message_data)
                except Exception as e:
                    print(f"Error parsing {msg_file}: {e}")
                    traceback.print_exc()
                    continue
            
            messages.sort(key=lambda x: x.get('date', ''), reverse=True)
            
        except Exception as e:
            print(f"Error reading messages for process {process_id}: {e}")
            traceback.print_exc()
        
        return messages
    
    def _parse_msg_file(self, file_path, process_id):
        try:
            msg = extract_msg.Message(file_path)
        except Exception as e:
            print(f"Error opening msg file {file_path}: {e}")
            return None
            
        try:
            subject = msg.subject or "No Subject"
            sender = msg.sender or "Unknown Sender"
            recipients = self._parse_recipients(msg)
            date = self._parse_date(msg.date)
            
            # Try to get HTML body first, then plain text
            body = ""
            body_type = "text"
            if hasattr(msg, 'htmlBody') and msg.htmlBody:
                body = msg.htmlBody
                body_type = "html"
            else:
                body = msg.body or ""
                body_type = "text"
            
            filename = os.path.basename(file_path)
            message_id = f"{process_id}_{os.path.splitext(filename)[0]}"
            
            attachments = self._extract_attachments(msg, message_id)
            thread_info = self._parse_thread_info(msg, subject)
            
            message_data = {
                "id": message_id,
                "subject": subject,
                "from": sender,
                "to": recipients,
                "date": date,
                "body": body,
                "body_type": body_type,
                "status": "untagged",
                "threadId": thread_info["thread_id"],
                "filename": filename,
                "attachments": attachments,
                "containsThread": thread_info["contains_thread"],
                "comments": []
            }
            
            return message_data
            
        except Exception as e:
            print(f"Error parsing message content from {file_path}: {e}")
            traceback.print_exc()
            return None
        finally:
            try:
                msg.close()
            except:
                pass
    
    def _parse_recipients(self, msg):
        recipients = []
        
        try:
            if hasattr(msg, 'to') and msg.to:
                recipients.extend([r.strip() for r in msg.to.split(';')])
            
            if hasattr(msg, 'cc') and msg.cc:
                recipients.extend([r.strip() for r in msg.cc.split(';')])
            
            if hasattr(msg, 'bcc') and msg.bcc:
                recipients.extend([r.strip() for r in msg.bcc.split(';')])
        except:
            pass
        
        return ', '.join(recipients) if recipients else "No Recipients"
    
    def _parse_date(self, date_str):
        if not date_str:
            return datetime.now().isoformat()
        
        try:
            # Remove timezone name for parsing
            date_str_clean = date_str.split(' (')[0]
            
            for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%a, %d %b %Y %H:%M:%S %Z', 
                       '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
                try:
                    dt = datetime.strptime(date_str_clean, fmt)
                    return dt.isoformat()
                except ValueError:
                    continue
            return datetime.now().isoformat()
        except:
            return datetime.now().isoformat()
    
    def _extract_attachments(self, msg, message_id):
        attachments = []
        
        try:
            if hasattr(msg, 'attachments') and msg.attachments:
                for i, attachment in enumerate(msg.attachments):
                    if hasattr(attachment, 'longFilename') and attachment.longFilename:
                        attachments.append({
                            "name": attachment.longFilename,
                            "url": f"/api/attachment/{message_id}/{i}",
                            "type": self._get_attachment_type(attachment.longFilename)
                        })
        except Exception as e:
            print(f"Error extracting attachments: {e}")
        
        return attachments
    
    def _get_attachment_type(self, filename):
        ext = os.path.splitext(filename)[1].lower()
        type_map = {
            '.pdf': 'PDF', '.doc': 'Word', '.docx': 'Word',
            '.xls': 'Excel', '.xlsx': 'Excel', '.ppt': 'PowerPoint',
            '.pptx': 'PowerPoint', '.txt': 'Text', '.jpg': 'Image',
            '.jpeg': 'Image', '.png': 'Image', '.gif': 'Image',
            '.zip': 'Archive', '.rar': 'Archive'
        }
        return type_map.get(ext, 'File')
    
    def _parse_thread_info(self, msg, subject):
        thread_id = f"thread_{hash(subject) % 10000}"
        
        contains_thread = False
        
        try:
            if hasattr(msg, 'body') and msg.body:
                body_lower = msg.body.lower()
                thread_indicators = ['original message', 'forwarded message', 'from:', 'sent:', 'to:', 'subject:']
                contains_thread = any(indicator in body_lower for indicator in thread_indicators)
        except:
            pass
        
        return {
            "thread_id": thread_id,
            "contains_thread": contains_thread
        }
    
    def get_attachment(self, process_id, message_id, attachment_index):
        try:
            original_msg_id = message_id.replace(f"{process_id}_", "")
            msg_file_path = os.path.join(self.base_folder, process_id, f"{original_msg_id}.msg")
            
            if not os.path.exists(msg_file_path):
                raise FileNotFoundError(f"Message file not found: {msg_file_path}")
            
            msg = extract_msg.Message(msg_file_path)
            try:
                if hasattr(msg, 'attachments') and msg.attachments:
                    if 0 <= attachment_index < len(msg.attachments):
                        attachment = msg.attachments[attachment_index]
                        
                        temp_dir = tempfile.mkdtemp()
                        temp_path = os.path.join(temp_dir, attachment.longFilename)
                        
                        with open(temp_path, 'wb') as f:
                            f.write(attachment.data)
                        
                        return temp_path
                    else:
                        raise IndexError(f"Attachment index {attachment_index} out of range")
                else:
                    raise ValueError("No attachments found in message")
            finally:
                msg.close()
        except Exception as e:
            print(f"Error getting attachment: {e}")
            traceback.print_exc()
            raise
    
    def update_message_status(self, process_id, message_id, status):
        try:
            self.message_status[message_id] = status
            # Invalidate cache for this process
            cache_key = f"messages_{process_id}"
            if cache_key in self._cache:
                del self._cache[cache_key]
            return True
        except Exception as e:
            print(f"Error updating status: {e}")
            return False
    
    def add_comment_to_message(self, process_id, message_id, comment_data):
        try:
            if message_id not in self.message_comments:
                self.message_comments[message_id] = []
            
            comment_data["date"] = datetime.now().isoformat()
            self.message_comments[message_id].append(comment_data)
            
            # Invalidate cache for this process
            cache_key = f"messages_{process_id}"
            if cache_key in self._cache:
                del self._cache[cache_key]
                
            return True
        except Exception as e:
            print(f"Error adding comment: {e}")
            return False
    
    def clear_cache(self):
        self._cache.clear()
        self._cache_time.clear()

# Create analyzer
analyzer = MSGAnalyzer()

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
            # Serve index.html for root path
            if self.path == '/' or self.path == '/index.html':
                with open('index.html', 'r', encoding='utf-8') as f:
                    html_content = f.read()
            else:
                # For other files, use the default implementation
                return super().do_GET()
        except FileNotFoundError:
            html_content = """
            <!DOCTYPE html>
            <html>
            <head><title>MSG Analyzer</title></head>
            <body>
                <h1>MSG Analyzer</h1>
                <p>Error: index.html not found. Please make sure it's in the same directory.</p>
            </body>
            </html>
            """
        
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.send_header('Content-length', str(len(html_content.encode('utf-8'))))
        self.end_headers()
        self.wfile.write(html_content.encode('utf-8'))
    
    def handle_api_request(self):
        try:
            if self.path == '/api/processes':
                self.handle_processes()
            elif self.path.startswith('/api/messages/'):
                self.handle_messages()
            elif self.path.startswith('/api/attachment/'):
                self.handle_attachment()
            elif self.path == '/api/health':
                self.handle_health()
            else:
                self.send_error(404, "API endpoint not found")
        except Exception as e:
            print(f"API error: {e}")
            traceback.print_exc()
            self.send_error(500, f"Server error: {str(e)}")
    
    def handle_api_post(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length > 0:
                post_data = self.rfile.read(content_length)
            else:
                post_data = b'{}'
                
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
        
        try:
            messages = analyzer.get_messages_for_process_cached(process_id)
            self.send_json_response(messages)
        except Exception as e:
            print(f"Error in handle_messages: {e}")
            self.send_error(500, f"Error retrieving messages: {str(e)}")
    
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
            attachment_path = analyzer.get_attachment(process_id, message_id, attachment_index)
            
            if attachment_path and os.path.exists(attachment_path):
                filename = os.path.basename(attachment_path)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/octet-stream')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(os.path.getsize(attachment_path)))
                self.end_headers()
                
                with open(attachment_path, 'rb') as f:
                    shutil.copyfileobj(f, self.wfile)
                
                # Clean up temporary file
                try:
                    temp_dir = os.path.dirname(attachment_path)
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                except Exception as e:
                    print(f"Error cleaning up temp files: {e}")
            else:
                self.send_error(404, "Attachment not found")
                
        except Exception as e:
            print(f"Error handling attachment: {e}")
            self.send_error(500, f"Error retrieving attachment: {str(e)}")
    
    def handle_update_status(self, process_id, message_id, post_data):
        try:
            data = json.loads(post_data.decode('utf-8'))
        except json.JSONDecodeError:
            self.send_error(400, "Invalid JSON")
            return
        
        status = data.get('status')
        if status not in ['keep', 'review', 'untagged']:
            self.send_error(400, "Invalid status")
            return
        
        success = analyzer.update_message_status(process_id, message_id, status)
        
        if success:
            self.send_json_response({"message": "Status updated successfully"})
        else:
            self.send_error(500, "Failed to update status")
    
    def handle_add_comment(self, process_id, message_id, post_data):
        try:
            data = json.loads(post_data.decode('utf-8'))
        except json.JSONDecodeError:
            self.send_error(400, "Invalid JSON")
            return
        
        # Validate required fields
        if 'key' not in data:
            self.send_error(400, "Missing required field: key")
            return
        
        success = analyzer.add_comment_to_message(process_id, message_id, data)
        
        if success:
            self.send_json_response({"message": "Comment added successfully"})
        else:
            self.send_error(500, "Failed to add comment")
    
    def handle_refresh_cache(self):
        analyzer.clear_cache()
        self.send_json_response({"message": "Cache refreshed successfully"})
    
    def handle_health(self):
        self.send_json_response({"status": "healthy", "timestamp": datetime.now().isoformat()})
    
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
    
    def log_message(self, format, *args):
        # Print basic logs for debugging
        print(f"{self.client_address[0]} - {format % args}")

def start_server(port=8000):
    # Check if index.html exists
    if not os.path.exists('index.html'):
        print("❌ ERROR: index.html not found in current directory!")
        print("Please make sure your HTML file is named 'index.html' and is in the same folder as this script.")
        
        # Create a basic index.html file
        try:
            with open('index.html', 'w', encoding='utf-8') as f:
                f.write("""<!DOCTYPE html>
<html>
<head>
    <title>MSG Analyzer</title>
    <style>body { font-family: Arial, sans-serif; padding: 20px; }</style>
</head>
<body>
    <h1>MSG Analyzer Backend is Running</h1>
    <p>The backend server is running, but the main interface file is missing.</p>
    <p>Please make sure you have the complete HTML file named 'index.html' in the same directory.</p>
</body>
</html>""")
            print("✅ Created a basic index.html file. Please replace it with the full interface.")
        except Exception as e:
            print(f"Could not create index.html: {e}")
    
    try:
        with socketserver.TCPServer(("", port), MSGHandler) as httpd:
            print(f"🚀 MSG Analyzer started on http://localhost:{port}")
            print("📁 Message folder:", os.path.abspath("msg_files"))
            print("⚡ Performance: Cache enabled")
            print("⏹️  To stop: Ctrl+C")
            
            print("🌐 Opening browser...")
            try:
                webbrowser.open(f'http://localhost:{port}')
            except:
                print("⚠️  Could not open browser automatically. Please navigate to the URL above.")
            
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\n🛑 Server stopped")
    except OSError as e:
        if e.errno == 48 or e.errno == 10048:  # Address already in use
            print(f"❌ Port {port} is already in use!")
            print(f"💡 Try using a different port: python script.py {port+1}")
        else:
            print(f"❌ Error starting server: {e}")

if __name__ == '__main__':
    import sys
    port = 8000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print("Invalid port number. Using default port 8000.")
    
    start_server(port)