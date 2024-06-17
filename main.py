import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import ssl

def send_email(to_email, company_name):
    subject = f"Transform {company_name}’s Financial Management with Advanced AI Solutions"
    body = f"""\
    <html>
    <body style="font-family: Arial, sans-serif; color: #000;">
        <p style="color: #000;">Dear {company_name} Team,</p>

        <p style="color: #000;">I hope this email finds you well. My name is <b>Insert Name</b>, and I represent <b>Company Name</b>, a leading provider of <b>innovative bookkeeping services</b>.</p>

        <p style="color: #000;">At <b>Company Name</b>, we understand the unique financial challenges that businesses like <b>{company_name}</b> face daily. That’s why we’ve developed a comprehensive bookkeeping solution that not only ensures accuracy but also leverages cutting-edge AI technology to provide insightful financial analysis, helping you make smarter business decisions.</p>

        <p style="color: #000;">Here’s how our services can benefit <b>{company_name}</b>:</p>
        <ul style="color: #000;">
            <li><b>Affordable Pricing</b>: We offer competitive rates without compromising on the quality of our services. Our goal is to provide you with top-notch bookkeeping solutions that fit within your budget.</li>
            <li><b>Advanced AI Financial Analysis</b>: Our state-of-the-art AI tools analyze your financial data to provide deeper insights into your business's performance. This includes identifying trends, forecasting future sales, and uncovering opportunities to maximize profitability.</li>
            <li><b>Customized Reporting</b>: Receive detailed, easy-to-understand reports that are tailored to your specific needs. Our AI-driven insights will help you stay ahead of the competition and make informed decisions.</li>
            <li><b>Time-Saving Automation</b>: With our automated processes, you can focus on what you do best – running your business. Leave the tedious bookkeeping tasks to us and enjoy more time to enhance your  offerings and customer experience.</li>
        </ul>

        <p style="color: #000;">We are confident that our modern approach to bookkeeping, combined with our commitment to affordability, makes us the ideal partner for <b>{company_name}</b>. Let us help you achieve financial clarity and drive your business's success.</p>

        <p style="color: #000;">Would you be available for a <b>brief call next week</b> to discuss how our services can be tailored to meet the specific needs of <b>{company_name}</b>? I’m happy to provide a <b>free consultation</b> and answer any questions you may have.</p>

        <p style="color: #000;">Thank you for considering <b>Fake Company</b>. We look forward to the opportunity to support <b>{company_name}</b> in achieving its financial goals.</p>

        <p style="color: #000;">Best regards,<br>
        <b>John Smith</b><br>
        <b>General Manager/Managing Partner</b><br>
        <b>Fake Company</b><br>
        <img src="cid:signature_image" alt="Signature Image" width="186" height="100"><br>
        Email:  <a href="mailto:johnsmith@gmail.com">johnsmith@gmail.com</a><br>
        Website: <a href="http://www.fakecompany.com">fakecompany.com</a><br>
        </p>
    </body>
    </html>
    """

    msg = MIMEMultipart('related')
    msg['From'] = "jonsmith@gmail.com"  # Replace with your email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    # Add image attachment
    image_path = 'C:/Users/USER/Downloads/star.png'  # Replace with your image path
    with open(image_path, 'rb') as img:
        mime_img = MIMEImage(img.read())
        mime_img.add_header('Content-ID', '<signature_image>')
        mime_img.add_header('Content-Disposition', 'inline', filename='signature_image')
        msg.attach(mime_img)

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            server.set_debuglevel(1)  # Enable debug output
            print(f"Connecting to SMTP server to send email to {company_name} at {to_email}...")
            server.login("johnsmith@gmail.com", "ysdo sdnv ctgl kres")  # Replace with your email and app password
            print("Logged in successfully, sending email...")
            server.send_message(msg)
            print(f"Email sent to {company_name} at {to_email}")
    except Exception as e:
        print(f"Failed to send email to {company_name} at {to_email}: {e}")

def process_excel(file_path, start_row, end_row):
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        print("Excel file read successfully.")
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return

    if df.empty:
        print("The Excel file is empty or doesn't contain the expected columns.")
        return

    missing_emails = []

    for index, row in df.iloc[start_row:end_row].iterrows():
        company_name = row.get('Company Name')
        email = row.get('Email')

        if pd.notnull(email) and email.strip():
            print(f"Attempting to send email to {company_name} at {email}.")
            send_email(email, company_name)
        else:
            phone_number = row.get('Phone Number')
            print(f"No email for {company_name}, adding to missing emails list.")
            missing_emails.append({'Company Name': company_name, 'Phone Number': phone_number})

    if missing_emails:
        try:
            missing_df = pd.DataFrame(missing_emails)
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
                missing_df.to_excel(writer, sheet_name='Sheet2', index=False)
            print("Missing emails logged to Sheet2.")
        except Exception as e:
            print(f"Failed to write to Excel file: {e}")

if __name__ == "__main__":
    file_path = 'C:/Users/USER/Documents/Mississauga_Business_Directory.xlsx'
    print(f"Processing Excel file: {file_path}")
    process_excel(file_path, start_row=3000, end_row=3500)  # Adjust start_row and end_row as needed
    print("Processing complete.")
