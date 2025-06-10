import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

def send_email(to_email: str, student_name: str, results_df: pd.DataFrame) -> dict:
    """Send email to parent with student's results in an HTML table."""
    try:
        msg = MIMEMultipart('alternative')
        msg["From"] = SENDER_EMAIL
        msg["To"] = to_email
        msg["Subject"] = f"Result Update for {student_name} - 2 Year 3 Semester"

        # Create HTML table with embedded CSS
        html_content = f"""
        <html>
        <head>
        <style>
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
        </style>
        </head>
        <body>
            <p>Dear Parent,</p>
            <p>Your child, <strong>{student_name}</strong>'s results for 2 Year - 3 Semester are out:</p>
            <table>
                <tr>
                    <th>Course Code</th>
                    <th>Course Name</th>
                    <th>Month Year</th>
                    <th>Grade</th>
                    <th>Grade Points</th>
                    <th>Credits</th>
                    <th>Result</th>
                </tr>
        """

        # Add rows to HTML table
        for _, row in results_df.iterrows():
            html_content += f"""
                <tr>
                    <td>{row['Course Code']}</td>
                    <td>{row['Course Name']}</td>
                    <td>{row['Month Year']}</td>
                    <td>{row['Grade']}</td>
                    <td>{row['Grade Points']}</td>
                    <td>{row['Credits']}</td>
                    <td>{row['Result']}</td>
                </tr>"""

        html_content += """
            </table>
            <p>Thank you,<br>School Administration</p>
        </body>
        </html>
        """

        # Attach both plain text and HTML versions
        text_content = f"Dear Parent,\n\nYour child, {student_name}'s results for 2 Year - 3 Semester are available. Please view this email in HTML format for better formatting.\n\nThank you,\nSchool Administration"
        
        msg.attach(MIMEText(text_content, 'plain'))
        msg.attach(MIMEText(html_content, 'html'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        return {"status": "success", "to_email": to_email}
    except Exception as e:
        return {"status": "error", "to_email": to_email, "error": str(e)}

st.title("Result Email Sender")
st.write("Upload a CSV file with student results to send emails to parents.")

# File upload
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

# Read CSV with error handling and data validation
if uploaded_file is not None:
    try:
        # Read CSV with specific data types
        df = pd.read_csv(uploaded_file, dtype={
            'student_name': str,
            'parent_email': str,
            'Course Code': str,
            'Course Name': str,
            'Month Year': str,
            'Grade': str,
            'Grade Points': float,
            'Credits': float,
            'Result': str
        })
        
        # Display preview with limited rows
        st.write("Preview of uploaded data (first 5 rows):")
        st.dataframe(df.head())
        
        # Show total number of records
        st.info(f"Total records: {len(df)}, Unique students: {df['student_name'].nunique()}")

        # Validate CSV columns
        required_columns = {"student_name", "parent_email", "Course Code", "Course Name", 
                          "Month Year", "Grade", "Grade Points", "Credits", "Result"}
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            st.stop()

        # Validate email addresses
        invalid_emails = df[~df['parent_email'].str.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        if not invalid_emails.empty:
            st.warning("Invalid email addresses found:")
            st.dataframe(invalid_emails[['student_name', 'parent_email']])
            if st.button("Continue anyway"):
                pass
            else:
                st.stop()

        # Group results by student
        grouped = df.groupby("student_name")

        # Send emails with progress bar
        if st.button("Send Emails"):
            results = []
            progress_bar = st.progress(0)
            total_students = len(grouped)
            
            for idx, (student_name, student_results) in enumerate(grouped):
                parent_email = student_results["parent_email"].iloc[0]
                if student_results["parent_email"].nunique() > 1:
                    st.warning(f"Multiple parent emails found for {student_name}. Using: {parent_email}")
                
                result = send_email(parent_email, student_name, student_results)
                results.append(result)
                progress_bar.progress((idx + 1) / total_students)

            # Display results
            success_count = sum(1 for r in results if r["status"] == "success")
            error_results = [r for r in results if r["status"] == "error"]
            
            st.success(f"Sent {success_count} emails successfully!")
            if error_results:
                st.error("Some emails failed to send:")
                for result in error_results:
                    st.write(f"Failed to send to {result['to_email']}: {result['error']}")

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")