# ================================
# 📦 PACKAGE INSTALL & IMPORTS
# ================================
import streamlit as st
import os
import fitz  # PyMuPDF
import pandas as pd
import zipfile
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ================================
# ⚙️ CONFIGURATION
# ================================
st.set_page_config(page_title="MR.Strict", layout="centered")
SENDER_EMAIL = "godwinscollabproject1@gmail.com"
SENDER_PASS = "mvyvmhawgrgbqivb"

# ================================
# 📧 EMAIL SENDER CLASS
# ================================
class EmailSender:
    def __init__(self, sender_email, sender_password):
        self.sender_email = sender_email
        self.sender_password = sender_password

    def send_email(self, to_email, subject, body, attachment_path=None):
        msg = MIMEMultipart()
        msg["From"] = self.sender_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        if attachment_path:
            with open(attachment_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                msg.attach(part)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.sender_email, self.sender_password)
        server.send_message(msg)
        server.quit()
        return f"📧 Email sent successfully to: {to_email}"

# ================================
# 📄 PDF TEXT EXTRACTION FUNCTION
# ================================
def extract_text(pdf_path_or_file):
    try:
        if isinstance(pdf_path_or_file, str):
            doc = fitz.open(pdf_path_or_file)
        else:
            doc = fitz.open(stream=pdf_path_or_file.read(), filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        return text if text.strip() else "⚠️ No text found in the PDF."
    except Exception as e:
        return f"❌ Error: {str(e)}"

# ================================
# 🎯 EVALUATION FUNCTION
# ================================
def evaluate_and_score(original, student):
    orig_lines = set(line.strip().lower() for line in original.strip().splitlines() if line.strip())
    stud_lines = set(line.strip().lower() for line in student.strip().splitlines() if line.strip())

    matching_lines = orig_lines & stud_lines
    total_lines = len(stud_lines)
    matched_lines = len(matching_lines)

    orig_words = set(original.lower().split())
    stud_words = set(student.lower().split())
    matching_words = orig_words & stud_words

    line_score = (matched_lines / total_lines * 100) if total_lines else 0
    word_score = (len(matching_words) / len(orig_words) * 100) if orig_words else 0
    final_score = (0.6 * word_score + 0.4 * line_score)

    if final_score >= 90:
        marks = "10/10"
    elif final_score >= 80:
        marks = "9/10"
    elif final_score >= 70:
        marks = "8/10"
    elif final_score >= 60:
        marks = "7/10"
    elif final_score >= 50:
        marks = "6/10"
    elif final_score >= 40:
        marks = "5/10"
    elif final_score >= 30:
        marks = "4/10"
    elif final_score >= 20:
        marks = "3/10"
    elif final_score >= 10:
        marks = "2/10"
    else:
        marks = "1/10"

    return marks, final_score

# ================================
# 🖥️ STREAMLIT UI
# ================================
st.title("🧠 MR.Strict - Auto PDF Evaluator + Email Marks")
st.markdown("This tool compares student answers with the original answer and emails the marks as a CSV.")

# Step 1 - Upload original answer PDF
st.subheader("📌 Step 1: Upload Original Answer PDF")
original_pdf = st.file_uploader("Upload Original Answer", type=["pdf"])

# Text display for original
if original_pdf:
    extracted_text = extract_text(original_pdf)
    st.text_area("📖 Extracted Text from Original Answer", value=extracted_text, height=200)

# Step 2 - Upload ZIP of student PDFs
st.subheader("📂 Step 2: Upload Student PDFs ZIP")
student_zip = st.file_uploader("Upload ZIP File", type=["zip"])

# Step 3 - Teacher's email
st.subheader("📧 Step 3: Enter Teacher's Email")
receiver_email = st.text_input("Teacher Email")

# Evaluate button
st.subheader("📤 Step 4: Click Below to Evaluate and Email")
if st.button("🧮 Evaluate & Send Email"):
    if not original_pdf or not student_zip or not receiver_email:
        st.error("❗ Please complete all steps before submitting.")
    else:
        st.info("⏳ Processing...")
        try:
            original_text = extract_text(original_pdf)
            if original_text.startswith("❌") or original_text.startswith("⚠️"):
                st.error("❌ Failed to extract text from original PDF.")
            else:
                with tempfile.TemporaryDirectory() as tmpdir:
                    zip_path = os.path.join(tmpdir, "students.zip")
                    with open(zip_path, "wb") as f:
                        f.write(student_zip.read())
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(tmpdir)

                    results = []
                    for file in os.listdir(tmpdir):
                        if file.endswith(".pdf"):
                            filepath = os.path.join(tmpdir, file)
                            student_text = extract_text(filepath)
                            if student_text.startswith("❌") or student_text.startswith("⚠️"):
                                continue
                            marks, final_score = evaluate_and_score(original_text, student_text)
                            results.append({
                                "Student File": file,
                                "Marks": marks,
                                "Score %": f"{final_score:.2f}%"
                            })

                    if not results:
                        st.warning("⚠️ No valid student PDFs found.")
                    else:
                        df = pd.DataFrame(results)
                        csv_path = os.path.join(tmpdir, "assignment_marks.csv")
                        df.to_csv(csv_path, index=False)

                        mailer = EmailSender(SENDER_EMAIL, SENDER_PASS)
                        mail_status = mailer.send_email(
                            to_email=receiver_email,
                            subject="📊 Assignment Marks - MR.Strict",
                            body="Hi Teacher,\n\nPlease find attached the evaluated assignment marks.\n\nRegards,\nMR.Strict 🤖",
                            attachment_path=csv_path
                        )
                        st.success(mail_status)
                        st.success("✅ Evaluation done and email sent!")
        except Exception as e:
            st.error(f"❌ Error occurred: {e}")
