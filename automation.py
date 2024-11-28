import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


df = pd.read_excel('students_attendance.xlsx')


def send_email(to_address, student_name, attendance_percentage):

    from_address = "vijaysuryavanshi6261@gmail.com"
    password = "ajcshkpoggenfnxc"
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = f"Attendance Alert for {student_name}"


    body = f"Dear Parent,\n\nYour child, {student_name}, has an attendance of {attendance_percentage}%, which is below the required 75%.\nPlease ensure that they attend classes regularly.\n\nBest Regards,\nSchool Administration"
    msg.attach(MIMEText(body, 'plain'))


    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_address, password)

    server.send_message(msg)
    server.quit()


low_attendance_students = df[df['Attendance (%)'] < 75]


for index, student in low_attendance_students.iterrows():
    student_name = student['Name']
    attendance_percentage = student['Attendance (%)']
    parent_email = student['Parent\'s Email']

    send_email(parent_email, student_name, attendance_percentage)

print("Emails sent successfully!")
