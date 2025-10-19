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
from concurrent.futures import ThreadPoolExecutor
import asyncio
import re

class OptimizedMSGAnalyzer:
    def __init__(self, base_folder="msg_files"):
        self.base_folder = base_folder
        self._ensure_base_folder()
        self._cache = {}
        self._cache_time = {}
        self.cache_timeout = 300
        self.message_status = {}
        self.message_comments = {}
        self.thread_pool = ThreadPoolExecutor(max_workers=4)  # Parallel processing
        self._message_index = {}  # Quick message lookup
        self._process_metadata = {}  # Store process metadata for faster access
    
    def _ensure_base_folder(self):
        if not os.path.exists(self.base_folder):
            os.makedirs(self.base_folder)
            print(f"Created base folder: {self.base_folder}")
    
    def get_processes(self):
        """Fast process listing - only checks folder structure"""
        if self._process_metadata and time.time() - self._process_metadata.get('_timestamp', 0) < 60:
            return self._process_metadata.get('processes', [])
        
        processes = []
        try:
            for item in os.listdir(self.base_folder):
                item_path = os.path.join(self.base_folder, item)
                if os.path.isdir(item_path):
                    # Fast count - just check for .msg files without parsing
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
        
        # Cache process metadata
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
            return messages
        
        try:
            # Get all .msg files
            msg_files = glob.glob(os.path.join(process_path, "*.msg"))
            
            # Sort by modification time (newest first) for faster access
            msg_files.sort(key=os.path.getmtime, reverse=True)
            
            # Apply pagination
            batch_files = msg_files[offset:offset + limit]
            
            print(f"Processing {len(batch_files)} files out of {len(msg_files)} total")
            
            # Use thread pool for parallel processing
            with ThreadPoolExecutor(max_workers=min(4, len(batch_files))) as executor:
                future_to_file = {
                    executor.submit(self._parse_msg_file_fast, msg_file, process_id): msg_file 
                    for msg_file in batch_files
                }
                
                for future in future_to_file:
                    try:
                        message_data = future.result(timeout=10)  # 10 second timeout per file
                        if message_data:
                            # Apply saved status and comments
                            message_id = message_data["id"]
                            if message_id in self.message_status:
                                message_data["status"] = self.message_status[message_id]
                            if message_id in self.message_comments:
                                message_data["comments"] = self.message_comments[message_id]
                            
                            messages.append(message_data)
                    except Exception as e:
                        msg_file = future_to_file[future]
                        print(f"Error parsing {msg_file}: {e}")
                        continue
            
            # Sort by date
            messages.sort(key=lambda x: x.get('date', ''), reverse=True)
            
        except Exception as e:
            print(f"Error reading messages for process {process_id}: {e}")
        
        return {
            "messages": messages,
            "total_count": len(glob.glob(os.path.join(process_path, "*.msg"))),
            "has_more": (offset + limit) < len(glob.glob(os.path.join(process_path, "*.msg"))),
            "offset": offset,
            "limit": limit
        }
    
    def _parse_msg_file_fast(self, file_path, process_id):
        """Fast parsing - only essential fields"""
        try:
            msg = extract_msg.Message(file_path)
        except Exception as e:
            print(f"Error opening msg file {file_path}: {e}")
            return None
            
        try:
            # Fast parsing - only get essential metadata first
            subject = msg.subject or "No Subject"
            sender = self._parse_sender_fast(msg)
            date = self._parse_date_fast(msg.date)
            
            filename = os.path.basename(file_path)
            message_id = f"{process_id}_{os.path.splitext(filename)[0]}"
            
            # Get basic body preview (first 200 chars)
            body_preview = self._get_body_preview(msg)
            
            message_data = {
                "id": message_id,
                "subject": subject,
                "from": sender,
                "date": date,
                "body_preview": body_preview,
                "status": "untagged",
                "filename": filename,
                "has_attachments": hasattr(msg, 'attachments') and bool(msg.attachments),
                "needs_full_parse": True  # Flag to parse full content when needed
            }
            
            return message_data
            
        except Exception as e:
            print(f"Error parsing message content from {file_path}: {e}")
            return None
        finally:
            try:
                msg.close()
            except:
                pass
    
    def get_message_full_content(self, process_id, message_id):
        """Load full content only when needed"""
        cache_key = f"full_{message_id}"
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        original_msg_id = message_id.replace(f"{process_id}_", "")
        msg_file_path = os.path.join(self.base_folder, process_id, f"{original_msg_id}.msg")
        
        if not os.path.exists(msg_file_path):
            return None
        
        try:
            msg = extract_msg.Message(msg_file_path)
            
            subject = msg.subject or "No Subject"
            sender = self._parse_sender_fast(msg)
            recipients = self._parse_recipients(msg)
            date = self._parse_date_fast(msg.date)
            
            # Get full body content
            body = ""
            body_type = "text"
            if hasattr(msg, 'htmlBody') and msg.htmlBody:
                body = msg.htmlBody
                body_type = "html"
            else:
                body = msg.body or ""
                body_type = "text"
            
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
                "status": self.message_status.get(message_id, "untagged"),
                "threadId": thread_info["thread_id"],
                "filename": os.path.basename(msg_file_path),
                "attachments": attachments,
                "containsThread": thread_info["contains_thread"],
                "comments": self.message_comments.get(message_id, [])
            }
            
            # Cache full content
            self._cache[cache_key] = message_data
            self._cache_time[cache_key] = time.time()
            
            return message_data
            
        except Exception as e:
            print(f"Error loading full content for {message_id}: {e}")
            return None
        finally:
            try:
                msg.close()
            except:
                pass
    
    def _parse_sender_fast(self, msg):
        """Fast sender parsing"""
        try:
            return msg.sender or "Unknown Sender"
        except:
            return "Unknown Sender"
    
    def _parse_date_fast(self, date_str):
        """Fast date parsing"""
        if not date_str:
            return datetime.now().isoformat()
        
        try:
            # Try simple parsing first
            for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%Y-%m-%d %H:%M:%S']:
                try:
                    dt = datetime.strptime(date_str.split(' (')[0], fmt)
                    return dt.isoformat()
                except ValueError:
                    continue
            return datetime.now().isoformat()
        except:
            return datetime.now().isoformat()
    
    def _get_body_preview(self, msg):
        """Get body preview without full parsing"""
        try:
            if hasattr(msg, 'body') and msg.body:
                body = msg.body
                # Clean and truncate
                cleaned = re.sub(r'\s+', ' ', body.strip())
                return cleaned[:200] + ('...' if len(cleaned) > 200 else '')
            return ""
        except:
            return ""
    
    def _parse_recipients(self, msg):
        """Parse recipients (only when full content needed)"""
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
    
    def _extract_attachments(self, msg, message_id):
        """Extract attachments (only when full content needed)"""
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
    
    # Keep existing methods for attachment, status, comments...
    def get_attachment(self, process_id, message_id, attachment_index):
        # ... (same as before)
        pass
    
    def update_message_status(self, process_id, message_id, status):
        # ... (same as before)
        pass
    
    def add_comment_to_message(self, process_id, message_id, comment_data):
        # ... (same as before)
        pass
    
    def clear_cache(self):
        self._cache.clear()
        self._cache_time.clear()

# Create optimized analyzer
analyzer = OptimizedMSGAnalyzer()

class OptimizedMSGHandler(http.server.SimpleHTTPRequestHandler):
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
        # ... (same as before)
        pass
    
    def handle_api_request(self):
        try:
            if self.path == '/api/processes':
                self.handle_processes()
            elif self.path.startswith('/api/messages/'):
                self.handle_messages_optimized()
            elif self.path.startswith('/api/message/'):
                self.handle_single_message()
            elif self.path.startswith('/api/attachment/'):
                self.handle_attachment()
            elif self.path == '/api/health':
                self.handle_health()
            else:
                self.send_error(404, "API endpoint not found")
        except Exception as e:
            print(f"API error: {e}")
            self.send_error(500, f"Server error: {str(e)}")
    
    def handle_messages_optimized(self):
        """Handle optimized messages endpoint with pagination"""
        path_parts = self.path.split('/')
        process_id = path_parts[3] if len(path_parts) > 3 else None
        
        if not process_id:
            self.send_error(400, "Process ID required")
            return
        
        # Parse query parameters for pagination
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
        """Handle request for a single message's full content"""
        path_parts = self.path.split('/')
        if len(path_parts) < 5:
            self.send_error(400, "Invalid message URL")
            return
        
        process_id = path_parts[3]
        message_id = path_parts[4]
        
        try:
            full_message = analyzer.get_message_full_content(process_id, message_id)
            if full_message:
                self.send_json_response(full_message)
            else:
                self.send_error(404, "Message not found")
        except Exception as e:
            print(f"Error loading full message: {e}")
            self.send_error(500, f"Error loading message: {str(e)}")
    
    # ... (keep other methods the same)

def start_server(port=8000):
    if not os.path.exists('index.html'):
        print("❌ ERROR: index.html not found!")
        return
    
    with socketserver.TCPServer(("", port), OptimizedMSGHandler) as httpd:
        print(f"🚀 OPTIMIZED MSG Analyzer started on http://localhost:{port}")
        print("⚡ Performance: Lazy loading + Parallel processing + Pagination")
        print("⏹️  To stop: Ctrl+C")
        
        webbrowser.open(f'http://localhost:{port}')
        
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n🛑 Server stopped")

if __name__ == '__main__':
    start_server()