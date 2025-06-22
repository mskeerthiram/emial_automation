#read data from sheets using python with gspread lib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("OnboardingForm").sheet1
expected_headers = ['Full Name', 'Email ID', 'Enrollment ID(register no)', 'Domain', 'Join Date', 'End Date']
data = sheet.get_all_records(expected_headers=expected_headers)
enrollment_dict = {}
for row in data:
    enrollment_id = row['Enrollment ID(register no)']
    enrollment_dict[enrollment_id] = row
print(enrollment_dict)
#generate unique pdf for each student automatically using fpdf lib
from fpdf import FPDF
internship_base_number=100
#loop
i=2
for enrollment_id, student in enrollment_dict.items():
    if student['is_sent']=="TRUE":
        continue
    internship_id=f"A0{internship_base_number}"
    pdf=FPDF()
    pdf.add_page()
    pdf.set_font("Arial",size=12)
    pdf.cell(200,10,txt="Offer Letter",align="C")
    pdf.ln(10)
    #add details
    pdf.cell(200,10,txt=f"Dear {student['Full Name']},",ln=True)
    pdf.cell(200,10,txt= f"Enrollment ID:{student['Enrollment ID(register no)']}",ln=True)
    pdf.ln(5)
#add body text
    pdf.multi_cell(0,10,txt="""Congratulations!We are pleased to offer you an opportunity to undergo On-The-Job Training (OJT) at VDart Academy.
This training program is designed to provide you with practical exposure and hands-on experience, enhancing your skills and preparing you for future career opportunities.
Please find below the details of the same""")
#dynamic details in table(vdart pdf)
    pdf.set_font("Arial","B",12)
    pdf.cell(0,10,txt=f'Enrollment:Academic Internship Enrollment ID: {internship_id}',ln=True)
    pdf.set_font("Arial","",12)
    pdf.cell(0,10,txt=f"Technology: {student['Domain']} Domain: Information Systems",ln=True)
    pdf.cell(0, 10, txt=f"Start Date: {student['Join Date']} End Date: {student['End Date']}", ln=True)
 # Step 6: Static details
    pdf.cell(0, 10, txt="Organization: VDart Academy Location: GCE - Mannapuram", ln=True)
    pdf.cell(0, 10, txt="Stipend: Not Applicable Reporting SME: Anubharathi P", ln=True)
    pdf.cell(0, 10, txt="Shift Time: 2:00 PM to 6:00 PM IST Shift Days: Monday to Friday", ln=True)
    pdf.cell(0, 10, txt="SME Email: anubharathi.p@vdartinc.com SME Mobile: +91 99445 48333", ln=True)
    pdf.ln(5)
   # Step 7: Closing
    pdf.multi_cell(0, 10, txt="""We believe this opportunity will contribute to your professional development, and we look forward to your active participation. 
Kindly confirm your acceptance by signing a copy of this letter by 18-Jun-2025.
Should you have any questions, feel free to contact us.""")
    filename=f"Offer_Letter_{enrollment_id}.pdf"
    pdf.output(filename)
    print("f {filename}generated!")
    internship_base_number+=1
#automate email using yagmail lib
    import yagmail
    from datetime import datetime,timedelta
    join_date_str = student.get('Join Date')  # safer access
    if join_date_str  and join_date_str.strip():
        join_date = datetime.strptime(join_date_str, "%d/%m/%Y")
    else:
        join_date = None
    scheduled_time=join_date-timedelta(days=2)
    if datetime.now()>=scheduled_time:
        yag=yagmail.SMTP("VdartAcademyOfficial123@gmail.com")
        subject="internship offer letter"
        body=f"""Dear {student['Full Name']},
We are pleased to offer you an opportunity to undergo On-The-Job Training (OJT) at VDart Academy.
Please find your offer letter attached for the internship starting on {join_date_str}.
Regards,  
VDart Academy"""
        yag.send(to=student['Email ID'],subject=subject ,contents=body,attachments=filename)
        print(f"email sent to {student['Email ID']}")
    else:
        print(f"email not sent to {student['Email ID']}")
    sheet.update_cell(i,7,"TRUE")
    i+=1





