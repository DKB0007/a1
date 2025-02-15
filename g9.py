import streamlit as st
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import imaplib
import email
from email.header import decode_header
import time
from typing import List, Dict, Tuple
import re
import json
import pandas as pd
import base64

# Predefined sender credentials
SENDER_EMAIL = "trickydkb@gmail.com"
APP_PASSWORD = "fooyxtamhsrqtmso"

# Configure attachment storage
ATTACHMENT_DIR = "attachments"
if not os.path.exists(ATTACHMENT_DIR):
    os.makedirs(ATTACHMENT_DIR)

class EmailHandler:
    def __init__(self):
        self.imap_server = "imap.gmail.com"
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587

    def send_bulk_email(self, recipients: List[Dict], subject: str, body_template: str, attachments: List[str] = None) -> List[Dict]:
        """Send emails to multiple recipients with personalized content and attachments"""
        results = []

        for recipient in recipients:
            try:
                # Personalize the body for each recipient
                personalized_body = body_template.format(
                    name=recipient.get('name', 'Team Member'),
                    role=recipient.get('role', 'Assignee')
                )

                msg = MIMEMultipart()
                msg["From"] = SENDER_EMAIL
                msg["To"] = recipient['email']
                msg["Subject"] = subject
                msg.attach(MIMEText(personalized_body, "plain"))

                # Add attachments if any
                if attachments:
                    for attachment_path in attachments:
                        with open(attachment_path, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            filename = os.path.basename(attachment_path)
                            part.add_header(
                                "Content-Disposition",
                                f"attachment; filename= {filename}",
                            )
                            msg.attach(part)

                with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                    server.starttls()
                    server.login(SENDER_EMAIL, APP_PASSWORD)
                    server.send_message(msg)

                results.append({
                    "email": recipient['email'],
                    "status": "success",
                    "message": "Email sent successfully"
                })

            except Exception as e:
                results.append({
                    "email": recipient['email'],
                    "status": "failed",
                    "message": str(e)
                })

        return results

    def save_attachment(self, task_id: str, sender_email: str, part) -> str:
        """Save email attachment to disk"""
        filename = part.get_filename()
        if filename:
            # Create task-specific directory
            task_dir = os.path.join(ATTACHMENT_DIR, task_id, sender_email)
            os.makedirs(task_dir, exist_ok=True)

            # Clean filename and save
            filename = re.sub(r'[^\w\-_\. ]', '_', filename)
            filepath = os.path.join(task_dir, filename)

            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))

            return filepath
        return None

    def fetch_replies(self) -> List[Dict]:
        """Fetch email replies with attachment handling"""
        replies = []
        try:
            mail = self.connect_imap()
            mail.select('inbox')

            # Search for emails from the last 24 hours
            date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
            search_criteria = f'(SINCE "{date}")'

            _, message_numbers = mail.search(None, search_criteria)

            for num in message_numbers[0].split():
                _, msg_data = mail.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                subject = decode_header(email_message["subject"])[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode()

                # Get sender details
                from_header = decode_header(email_message["from"])[0][0]
                if isinstance(from_header, bytes):
                    from_header = from_header.decode()

                # Extract email address
                email_pattern = r'<(.+?)>'
                match = re.search(email_pattern, from_header)
                sender_email = match.group(1) if match else from_header

                if "Re: Task Assignment" in subject:
                    task_id = self.extract_task_id(subject)

                    # Handle attachments and body
                    attachments = []
                    body = ""

                    if email_message.is_multipart():
                        for part in email_message.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode('utf-8', 'ignore')
                            elif part.get_content_maintype() != 'multipart':
                                attachment_path = self.save_attachment(task_id, sender_email, part)
                                if attachment_path:
                                    attachments.append(attachment_path)
                    else:
                        body = email_message.get_payload(decode=True).decode('utf-8', 'ignore')

                    replies.append({
                        "task_id": task_id,
                        "from_email": sender_email,
                        "subject": subject,
                        "body": body,
                        "attachments": attachments,
                        "timestamp": email_message["date"]
                    })

            mail.close()
            mail.logout()

        except Exception as e:
            st.error(f"Error fetching replies: {str(e)}")

        return replies

    def extract_task_id(self, subject: str) -> str:
        """Extract task ID from email subject"""
        match = re.search(r'\[Task ID: (.+?)\]', subject)
        return match.group(1) if match else ""

    def connect_imap(self):
        """Establish IMAP connection"""
        mail = imaplib.IMAP4_SSL(self.imap_server)
        mail.login(SENDER_EMAIL, APP_PASSWORD)
        return mail

def process_reply(reply: Dict) -> str:
    """Process reply content and determine status"""
    body_lower = reply["body"].lower()

    if "completed" in body_lower or "done" in body_lower:
        return "completed"
    elif "in progress" in body_lower or "working" in body_lower:
        return "in progress"
    elif "acknowledge" in body_lower or "received" in body_lower:
        return "acknowledged"
    else:
        return "replied"

def initialize_session_state():
    """Initialize all session state variables"""
    if 'assignments' not in st.session_state:
        st.session_state.assignments = []
    if 'email_handler' not in st.session_state:
        st.session_state.email_handler = EmailHandler()
    if 'task_statuses' not in st.session_state:
        st.session_state.task_statuses = {}
    if 'recipients' not in st.session_state:
        st.session_state.recipients = [{"email": "", "name": "", "role": ""}]
    if 'replies_history' not in st.session_state:
        st.session_state.replies_history = {}
    if 'attachments' not in st.session_state:
        st.session_state.attachments = []

def add_recipient():
    """Add a new recipient to the form"""
    st.session_state.recipients.append({"email": "", "name": "", "role": ""})

def remove_recipient(index):
    """Remove a recipient from the form"""
    if len(st.session_state.recipients) > 1:
        st.session_state.recipients.pop(index)
        st.rerun()

def main():
    st.title("📋 Multi-User Task Assignment System")

    initialize_session_state()

    # Task Assignment Form
    st.markdown("### Assign New Task")

    # Recipients management
    st.subheader("Recipients")
    cols = st.columns([3, 2, 2, 1])
    with cols[0]:
        st.markdown("**Email**")
    with cols[1]:
        st.markdown("**Name**")
    with cols[2]:
        st.markdown("**Role**")

    recipients_valid = True
    for idx, recipient in enumerate(st.session_state.recipients):
        cols = st.columns([3, 2, 2, 1])
        with cols[0]:
            recipient['email'] = st.text_input("Email", recipient['email'], key=f"email_{idx}")
        with cols[1]:
            recipient['name'] = st.text_input("Name", recipient['name'], key=f"name_{idx}")
        with cols[2]:
            recipient['role'] = st.text_input("Role", recipient['role'], key=f"role_{idx}")
        with cols[3]:
            if st.button("❌", key=f"remove_{idx}"):
                remove_recipient(idx)

        if not recipient['email']:
            recipients_valid = False

    if st.button("➕ Add Recipient"):
        add_recipient()
        st.rerun()

    # Task details form
    with st.form("task_form"):
        task_title = st.text_input("Task Title")
        priority = st.select_slider("Priority Level",
                                  options=["Low", "Medium", "High", "Urgent"],
                                  value="Medium")
        due_date = st.date_input("Due Date",
                                min_value=datetime.now().date(),
                                value=datetime.now().date() + timedelta(days=1))

        task_description = st.text_area("Task Description", height=100)

        # File upload
        uploaded_files = st.file_uploader("Attach Files", accept_multiple_files=True)

        submit_button = st.form_submit_button("Assign Task", type="primary")

        if submit_button and recipients_valid:
            if task_title and task_description:
                # Generate unique task ID
                task_id = f"TASK-{datetime.now().strftime('%Y%m%d%H%M%S')}"

                # Save uploaded files
                attachments = []
                if uploaded_files:
                    task_dir = os.path.join(ATTACHMENT_DIR, task_id, "original")
                    os.makedirs(task_dir, exist_ok=True)

                    for uploaded_file in uploaded_files:
                        file_path = os.path.join(task_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        attachments.append(file_path)

                # Create email content
                subject = f"Task Assignment: {task_title} [Task ID: {task_id}] (Priority: {priority})"
                body_template = f"""
Dear {{name}},

You have been assigned a new task as {{role}}.

Task Details:
-------------
Title: {task_title}
Priority: {priority}
Due Date: {due_date}

Description:
{task_description}

Please reply to this email with one of the following:
- "Acknowledged" to confirm receipt
- "In Progress" when you start working
- "Completed" when the task is done

Task ID: {task_id}

Best regards,
Task Management System
"""

                # Send emails to all recipients
                results = st.session_state.email_handler.send_bulk_email(
                    st.session_state.recipients,
                    subject,
                    body_template,
                    attachments
                )

                # Process results
                success_count = sum(1 for r in results if r['status'] == 'success')
                if success_count == len(results):
                    st.success(f"Task assigned successfully to {success_count} recipients!")
                    st.balloons()

                    # Save assignment with attachment info
                    st.session_state.assignments.append({
                        'id': task_id,
                        'title': task_title,
                        'priority': priority,
                        'due_date': due_date,
                        'description': task_description,
                        'recipients': st.session_state.recipients.copy(),
                        'attachments': attachments,
                        'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

                    # Initialize status tracking for each recipient
                    for recipient in st.session_state.recipients:
                        status_key = f"{task_id}_{recipient['email']}"
                        st.session_state.task_statuses[status_key] = "pending"
                        st.session_state.replies_history[status_key] = []

                else:
                    st.warning(f"Task assigned to {success_count}/{len(results)} recipients")
                    for result in results:
                        if result['status'] == 'failed':
                            st.error(f"Failed to send to {result['email']}: {result['message']}")
            else:
                st.error("Please fill in all required fields.")

    # Check for replies
    if st.button("Check for Replies"):
        with st.spinner("Fetching replies..."):
            replies = st.session_state.email_handler.fetch_replies()

            if replies:
                st.success(f"Found {len(replies)} new replies!")

                # Process and organize replies by task and recipient
                for reply in replies:
                    task_id = reply['task_id']
                    sender_email = reply['from_email']
                    status = process_reply(reply)

                    # Update status and save reply with attachments
                    status_key = f"{task_id}_{sender_email}"
                    if status_key in st.session_state.task_statuses:
                        st.session_state.task_statuses[status_key] = status

                        # Correctly store attachments within replies history
                        st.session_state.replies_history[status_key].append({
                            'timestamp': reply['timestamp'],
                            'message': reply['body'],
                            'status': status,
                            'attachments': reply['attachments']  # Store attachments here
                        })
    # Display Task Dashboard
    if st.session_state.assignments:
        st.markdown("### Task Dashboard")

        for task in st.session_state.assignments:
            with st.expander(f"Task: {task['title']} (ID: {task['id']})"):
                st.markdown(f"""
                **Priority:** {task['priority']}
                **Due Date:** {task['due_date']}
                **Created:** {task['time']}

                **Description:**
                {task['description']}
                """)

                # Display original task attachments
                if task.get('attachments'):
                    st.markdown("#### Original Task Attachments")
                    for attachment in task['attachments']:
                        filename = os.path.basename(attachment)
                        if os.path.exists(attachment):
                            with open(attachment, "rb") as f:
                                btn = st.download_button(
                                    label=f"📎 Download {filename}",
                                    data=f.read(),
                                    file_name=filename,
                                    mime="application/octet-stream"
                                )

                # Create status table for recipients
                status_data = []
                for recipient in task['recipients']:
                    status_key = f"{task['id']}_{recipient['email']}"
                    status = st.session_state.task_statuses.get(status_key, "pending")
                    replies = st.session_state.replies_history.get(status_key, [])

                    last_reply_timestamp = replies[-1]['timestamp'] if replies else 'No replies'
                    total_attachments = sum(len(reply.get('attachments', [])) for reply in replies)

                    status_data.append({
                        'Name': recipient['name'],
                        'Email': recipient['email'],
                        'Role': recipient['role'],
                        'Status': status.title(),
                        'Last Reply': last_reply_timestamp,
                        'Attachments': total_attachments
                    })

                if status_data:
                    st.markdown("#### Recipient Status")
                    df = pd.DataFrame(status_data)
                    st.dataframe(df, hide_index=True)

                    # Show reply history for each recipient
                    st.markdown("#### Reply History")
                    for recipient in task['recipients']:
                        status_key = f"{task['id']}_{recipient['email']}"
                        replies = st.session_state.replies_history.get(status_key, [])

                        if replies:
                            st.markdown(f"**{recipient['name']} ({recipient['email']})**")
                            for reply in replies:
                                st.markdown(f"""
                                > {reply['timestamp']} - Status: {reply['status'].title()}
                                > {reply['message'][:200]}{'...' if len(reply['message']) > 200 else ''}
                                """)

                                # Display reply attachments
                                if reply.get('attachments'):
                                    st.markdown("📎 **Attachments:**")
                                    for attachment in reply['attachments']:
                                        filename = os.path.basename(attachment)
                                        if os.path.exists(attachment):
                                            with open(attachment, "rb") as f:
                                                btn = st.download_button(
                                                    label=f"Download {filename}",
                                                    data=f.read(),
                                                    file_name=filename,
                                                    mime="application/octet-stream",
                                                    key=f"{status_key}_{filename}"
                                                )

                # Add task statistics
                st.markdown("#### Task Statistics")
                total_replies = sum(len(replies) for replies in st.session_state.replies_history.values())
                total_attachments = sum(
                    len(reply.get('attachments', []))
                    for replies in st.session_state.replies_history.values()
                    for reply in replies
                )

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Recipients", len(task['recipients']))
                with col2:
                    st.metric("Total Replies", total_replies)
                with col3:
                    st.metric("Total Attachments", total_attachments)
                with col4:
                    completed_count = sum(
                        1 for status in st.session_state.task_statuses.values()
                        if status == "completed"
                    )
                    st.metric("Completed", f"{completed_count}/{len(task['recipients'])}")

    # Add data persistence using pickle
    if st.button("Save System State"):
        try:
            import pickle  # Import here to avoid unnecessary dependency if not used

            state_data = {
                'assignments': st.session_state.assignments,
                'task_statuses': st.session_state.task_statuses,
                'replies_history': st.session_state.replies_history,
                'recipients': st.session_state.recipients
            }

            with open('task_system_state.pkl', 'wb') as f:  # Use .pkl extension
                pickle.dump(state_data, f)

            st.success("System state saved successfully!")
        except Exception as e:
            st.error(f"Error saving system state: {str(e)}")

    if st.button("Load System State"):
        try:
            import pickle
            if os.path.exists('task_system_state.pkl'):
                with open('task_system_state.pkl', 'rb') as f:
                    state_data = pickle.load(f)

                st.session_state.assignments = state_data['assignments']
                st.session_state.task_statuses = state_data['task_statuses']
                st.session_state.replies_history = state_data['replies_history']
                st.session_state.recipients = state_data.get('recipients', [{"email": "", "name": "", "role": ""}])  #Handle missing key

                st.success("System state loaded successfully!")
                st.rerun()
            else:
                st.warning("No saved state found.")
        except Exception as e:
            st.error(f"Error loading system state: {str(e)}")

if __name__ == "__main__":
    main()