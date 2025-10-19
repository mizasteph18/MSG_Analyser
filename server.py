from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import os
import json
import extract_msg
from datetime import datetime
import tempfile
import shutil
import glob
from typing import List, Dict, Any
import mimetypes

app = Flask(__name__)
CORS(app)

# Configuration
MSG_FOLDER = "msg_files"
BASE_DIR = os.path.abspath(MSG_FOLDER)

class MSGAnalyzer:
    def __init__(self, base_folder: str):
        self.base_folder = base_folder
        self._ensure_base_folder()
    
    def _ensure_base_folder(self):
        if not os.path.exists(self.base_folder):
            os.makedirs(self.base_folder)
            print(f"Created base folder: {self.base_folder}")
    
    def get_processes(self) -> List[Dict[str, str]]:
        processes = []
        try:
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
        
        if not processes:
            self._create_sample_structure()
            return self.get_processes()
        
        return processes
    
    def _create_sample_structure(self):
        sample_folders = ["Marketing_Process", "Sales_Process", "HR_Process", "IT_Process"]
        for folder in sample_folders:
            folder_path = os.path.join(self.base_folder, folder)
            os.makedirs(folder_path, exist_ok=True)
        
        print("Created sample folder structure. Please add .msg files to the subfolders.")
    
    def get_messages_for_process(self, process_id: str) -> List[Dict[str, Any]]:
        messages = []
        process_path = os.path.join(self.base_folder, process_id)
        
        if not os.path.exists(process_path):
            return messages
        
        try:
            msg_files = glob.glob(os.path.join(process_path, "*.msg"))
            
            for msg_file in msg_files:
                try:
                    message_data = self._parse_msg_file(msg_file, process_id)
                    if message_data:
                        messages.append(message_data)
                except Exception as e:
                    print(f"Error parsing {msg_file}: {e}")
                    continue
            
            messages.sort(key=lambda x: x.get('date', ''), reverse=True)
            
        except Exception as e:
            print(f"Error reading messages for process {process_id}: {e}")
        
        return messages
    
    def _parse_msg_file(self, file_path: str, process_id: str) -> Dict[str, Any]:
        msg = extract_msg.openMsg(file_path)
        
        try:
            subject = msg.subject or "No Subject"
            sender = msg.sender or "Unknown Sender"
            recipients = self._parse_recipients(msg)
            date = self._parse_date(msg.date)
            body = msg.body or ""
            
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
                "status": "untagged",
                "threadId": thread_info["thread_id"],
                "filename": filename,
                "attachments": attachments,
                "containsThread": thread_info["contains_thread"],
                "comments": []
            }
            
            return message_data
            
        finally:
            msg.close()
    
    def _parse_recipients(self, msg) -> str:
        recipients = []
        
        if hasattr(msg, 'to') and msg.to:
            recipients.extend([r.strip() for r in msg.to.split(';')])
        
        if hasattr(msg, 'cc') and msg.cc:
            recipients.extend([r.strip() for r in msg.cc.split(';')])
        
        if hasattr(msg, 'bcc') and msg.bcc:
            recipients.extend([r.strip() for r in msg.bcc.split(';')])
        
        return ', '.join(recipients) if recipients else "No Recipients"
    
    def _parse_date(self, date_str: str) -> str:
        if not date_str:
            return datetime.now().isoformat()
        
        try:
            for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%a, %d %b %Y %H:%M:%S %Z', '%Y-%m-%d %H:%M:%S']:
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt.isoformat()
                except ValueError:
                    continue
            
            return datetime.now().isoformat()
        except:
            return datetime.now().isoformat()
    
    def _extract_attachments(self, msg, message_id: str) -> List[Dict[str, str]]:
        attachments = []
        
        if hasattr(msg, 'attachments') and msg.attachments:
            for i, attachment in enumerate(msg.attachments):
                if hasattr(attachment, 'longFilename') and attachment.longFilename:
                    attachments.append({
                        "name": attachment.longFilename,
                        "url": f"/api/attachment/{message_id}/{i}",
                        "type": self._get_attachment_type(attachment.longFilename)
                    })
        
        return attachments
    
    def _get_attachment_type(self, filename: str) -> str:
        ext = os.path.splitext(filename)[1].lower()
        type_map = {
            '.pdf': 'PDF', '.doc': 'Word', '.docx': 'Word',
            '.xls': 'Excel', '.xlsx': 'Excel', '.ppt': 'PowerPoint',
            '.pptx': 'PowerPoint', '.txt': 'Text', '.jpg': 'Image',
            '.jpeg': 'Image', '.png': 'Image', '.gif': 'Image',
            '.zip': 'Archive', '.rar': 'Archive'
        }
        return type_map.get(ext, 'File')
    
    def _parse_thread_info(self, msg, subject: str) -> Dict[str, Any]:
        thread_id = f"thread_{hash(subject) % 10000}"
        
        contains_thread = False
        
        if hasattr(msg, 'body') and msg.body:
            body_lower = msg.body.lower()
            thread_indicators = ['original message', 'forwarded message', 'from:', 'sent:', 'to:', 'subject:']
            contains_thread = any(indicator in body_lower for indicator in thread_indicators)
        
        return {
            "thread_id": thread_id,
            "contains_thread": contains_thread
        }
    
    def get_attachment(self, process_id: str, message_id: str, attachment_index: int) -> str:
        original_msg_id = message_id.replace(f"{process_id}_", "")
        msg_file_path = os.path.join(self.base_folder, process_id, f"{original_msg_id}.msg")
        
        if not os.path.exists(msg_file_path):
            raise FileNotFoundError(f"Message file not found: {msg_file_path}")
        
        msg = extract_msg.openMsg(msg_file_path)
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
    
    def update_message_status(self, process_id: str, message_id: str, status: str) -> bool:
        print(f"Updated message {message_id} in process {process_id} to status: {status}")
        return True
    
    def add_comment_to_message(self, process_id: str, message_id: str, comment_data: Dict) -> bool:
        print(f"Added comment to message {message_id} in process {process_id}: {comment_data}")
        return True

msg_analyzer = MSGAnalyzer(BASE_DIR)

# Servir le fichier HTML principal
@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

# Servir les fichiers statics (CSS, JS, etc.)
@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory('.', path)

# API Routes
@app.route('/api/processes', methods=['GET'])
def get_processes():
    processes = msg_analyzer.get_processes()
    return jsonify(processes)

@app.route('/api/messages/<process_id>', methods=['GET'])
def get_messages(process_id):
    try:
        messages = msg_analyzer.get_messages_for_process(process_id)
        return jsonify(messages)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/attachment/<process_id>/<message_id>/<int:attachment_index>', methods=['GET'])
def get_attachment(process_id, message_id, attachment_index):
    try:
        attachment_path = msg_analyzer.get_attachment(process_id, message_id, attachment_index)
        
        if attachment_path and os.path.exists(attachment_path):
            filename = os.path.basename(attachment_path)
            
            @app.after_request
            def remove_temp_file(response):
                try:
                    temp_dir = os.path.dirname(attachment_path)
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                except:
                    pass
                return response
            
            return send_file(attachment_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({"error": "Attachment not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/message/<process_id>/<message_id>/status', methods=['POST'])
def update_message_status(process_id, message_id):
    try:
        data = request.get_json()
        status = data.get('status')
        
        if status not in ['keep', 'review', 'untagged']:
            return jsonify({"error": "Invalid status"}), 400
        
        success = msg_analyzer.update_message_status(process_id, message_id, status)
        
        if success:
            return jsonify({"message": "Status updated successfully"})
        else:
            return jsonify({"error": "Failed to update status"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/message/<process_id>/<message_id>/comment', methods=['POST'])
def add_comment(process_id, message_id):
    try:
        data = request.get_json()
        
        required_fields = ['key', 'labels', 'text']
        for field in required_fields:
            if field not in data:
                return jsonify({"error": f"Missing field: {field}"}), 400
        
        success = msg_analyzer.add_comment_to_message(process_id, message_id, data)
        
        if success:
            return jsonify({"message": "Comment added successfully"})
        else:
            return jsonify({"error": "Failed to add comment"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})

@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Resource not found"}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "Internal server error"}), 500

if __name__ == '__main__':
    print(f"🚀 MSG Analyzer Backend starting...")
    print(f"📁 Base folder: {BASE_DIR}")
    print("📂 Available processes:", msg_analyzer.get_processes())
    print("🌐 Server running on: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000)