# Python code to illustrate Sending mail with attachments 
# from your Gmail account 

# libraries to be imported 
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import openpyxl as xw


wbOne = xw.load_workbook('responses.xlsx')
sht1 = wbOne['Sheet1']
totalStudents = sht1.max_row

emailIDs=[]
names=[]


for registrants in range(1, totalStudents):
    emailIDs.append(str(sht1['B' + str(registrants)].value))
    names.append(str(sht1['A' + str(registrants)].value))


fromaddr = "<YOUR EMAIL ID>"

# sending the mail 

# s.sendmail(fromaddr, toaddr, text) 

for toAdd in range(totalStudents):
    # storing the receivers email address 
    
    ## this condition is to send email to only first N number of ids in spreadsheet 
    if (toAdd < 17):
        continue

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 

    # storing the senders email address 
    msg['From'] = fromaddr 

    # storing the subject 
    msg['Subject'] = "Your certificate from BVP-HEC!"
    
    # string to store the body of the mail 
    body = "Thanks for participating in IDEAHACK organised by BVP HACKEREARTH CLUB. Here's your e-Certificate of Participation."

    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 

    msg['To'] = emailIDs[toAdd] 
    print (emailIDs[toAdd])

    attachment = open("/home/jaspreet/Desktop/eCertiGenerator/generatedeCertis/"+ names[toAdd] + ".png", "rb") 
    print ("eCerti: "+str(attachment))

    # open the file to be sent 
    filename = "eCertificate.png"

    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 

    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 

    # encode into base64 
    encoders.encode_base64(p) 

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 

    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 

    # start TLS for security 
    s.starttls() 

    # Authentication 
    s.login(fromaddr, '<YOUR PASSWORD>') 

    # Converts the Multipart msg into a string 
    text = msg.as_string() 

    #sending email 
    s.sendmail(fromaddr, emailIDs[toAdd], text)

    attachment.close()
    # terminating the session 
    s.quit() 
