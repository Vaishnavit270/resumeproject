import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os

# SMTP Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "vaishnavit22hcompe@student.mes.ac.in"  # Change this
EMAIL_PASSWORD = "elfc tfvx abuf vdzp"  # Change this

def create_offer_letter(candidate_name, job_title, company_name):
    """ Generates a professional PDF job offer letter. """
    pdf_filename = f"Offer_Letter_{candidate_name.replace(' ', '_')}.pdf"
    
    c = canvas.Canvas(pdf_filename, pagesize=letter)
    
    # Set margins and font styles
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 750, f"{company_name}")
    
    c.setFont("Helvetica", 12)
    c.setFillColor(colors.black)
    
    # Draw a separator line
    c.line(50, 740, 550, 740)
    
    # Add date
    c.drawString(50, 720, "Date: _______________")
    
    # Salutation
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 690, f"Dear {candidate_name},")
    
    # Offer letter content
    c.setFont("Helvetica", 12)
    text = f"""
    We are delighted to offer you the position of {job_title} at {company_name}.
    
    Your skills and experience have impressed us, and we believe you will be 
    a valuable addition to our team. Please find below the key details 
    regarding your employment:

    - Position: {job_title}
    - Company: {company_name}
    - Start Date: _______________
    - Salary: _______________

    Please sign and return this letter as confirmation of your acceptance.
    
    Welcome to the team!

    Best regards,
    HR Team,
    {company_name}
    """
    
    # Write the offer letter content line by line
    text_lines = text.split("\n")
    y_position = 670  # Starting position for the text
    for line in text_lines:
        c.drawString(50, y_position, line.strip())
        y_position -= 20  # Move down for the next line
    
    # Signature area
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y_position - 40, "___________________________")
    c.drawString(50, y_position - 55, "Authorized Signatory")
    c.drawString(50, y_position - 70, f"{company_name}")

    c.save()
    return pdf_filename

def send_email(to_email, candidate_name, job_title, company_name):
    """ Sends an email with the job offer letter attached. """
    if pd.isna(to_email) or not to_email.strip():
        st.warning(f"Skipping {candidate_name} due to missing email.")
        return False

    to_email = str(to_email).strip()
    subject = f"Job Selection Notification from {company_name}"
    
    body = f"""Dear {candidate_name},

Congratulations! You have been selected for the {job_title} position at {company_name}.

Attached is your official job offer letter. Please review the details and get back to us.

Best regards,  
HR Team  
{company_name}
"""

    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Generate offer letter PDF
    pdf_filename = create_offer_letter(candidate_name, job_title, company_name)

    # Attach PDF to email
    with open(pdf_filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={pdf_filename}")
        msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
        server.quit()
        os.remove(pdf_filename)  # Delete the PDF after sending
        return True
    except Exception as e:
        st.error(f"Error sending email to {candidate_name} ({to_email}): {e}")
        return False

def filter_candidates(df, job_requirements):
    """ Filters candidates based on job field matching. """
    if "Predicted Field" not in df.columns:
        st.error("CSV file is missing the required 'Predicted Field' column.")
        return pd.DataFrame()

    filtered_df = df[df['Predicted Field'].astype(str).str.contains(job_requirements, case=False, na=False)]

    if "Mail" in filtered_df.columns:
        filtered_df = filtered_df.dropna(subset=['Mail'])  
        filtered_df = filtered_df[filtered_df['Mail'].astype(str).str.strip() != ""]  

    return filtered_df

def main():
    st.title("Candidate Selection System")
    
    company_name = st.text_input("Enter Your Company Name")
    if not company_name.strip():
        st.warning("Please enter your company name before proceeding.")
    
    uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])
    job_requirements = st.text_input("Enter Job Requirements (e.g., Data Science, Marketing)")
    
    if uploaded_file and job_requirements and company_name.strip():
        df = pd.read_csv(uploaded_file)

        if "Name" not in df.columns or "Mail" not in df.columns:
            st.error("CSV file must contain 'Name' and 'Mail' columns.")
            return

        filtered_candidates = filter_candidates(df, job_requirements)
        
        st.write("### Filtered Candidates")
        st.dataframe(filtered_candidates)
        
        selected_candidates = st.multiselect("Select Candidates to Notify", filtered_candidates['Name'].tolist())
        
        if st.button("Notify Selected Candidates"):
            for candidate in selected_candidates:
                candidate_row = filtered_candidates[filtered_candidates['Name'] == candidate].iloc[0]
                email = candidate_row['Mail']

                if send_email(email, candidate, job_requirements, company_name):
                    st.success(f"Email sent to {candidate} ({email}) with Offer Letter")
                else:
                    st.error(f"Failed to send email to {candidate}")
    
if __name__ == "__main__":
    main()
