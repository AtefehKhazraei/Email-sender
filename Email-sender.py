import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage


# Sending Email to a excel list of email
#-------------------------------------
# Reading an excel file using Python
import xlrd
i=0
# Give the location of the file. 
loc = ('email.xlsx')

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
sheet.cell_value(0, 0)


for i in range(234,sheet.nrows):
    x=sheet.cell_value(i, 0)
    print(i)
    print(sheet.cell_value(i, 0))
    msg = MIMEMultipart('related')
    msg['Subject'] = "Enter Subject of email here"
    msg['From'] = 'Enter your Email here'
    msg['To'] = sheet.cell_value(i, 0)

    html = """\
    <html>
      <head></head>
        <body>
          <img src="cid:image1" alt="p" ><br>
          <p>Enter your HTML massage here . </p>
        </div>          
        </body>
    </html>
    """
    # Record the MIME types of text/html.
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    msg.attach(part2)

    fp = open('pic.jpg', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)
    # if your email is not gmail thhen change the bleow
    server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server_ssl.ehlo()
    server_ssl.login('Enter your Email here', 'enter your email password here')

    server_ssl.sendmail('Enter your Email here', x, msg.as_string())

    server_ssl.close()
    print('successfully sent the mail to'+x)
#--------------------------------------------





# this part is for geting a confrim email in your inbox . just to see how the email looks like in your target list inbox
# you can fill it with your own email .
#-----------------------------------------------------------------------------------------------------------------
msg = MIMEMultipart('related')
msg['Subject'] = " Enter Subject of email here"
msg['From'] = 'Enter your Email here'
msg['To'] = 'Enter your Email here'

html = """\
<html>
  <head></head>
    <body>
      <img src="cid:image1" alt="p" ><br>
      <p>Coppy the same html massage that you used for the code above . this will send the same massage to your own inbox to see how it looks on their inbox</p>
    </div>          
    </body>
</html>
"""
# Record the MIME types of text/html.
part2 = MIMEText(html, 'html')

# Attach parts into message container.
msg.attach(part2)

fp = open('pic.jpeg', 'rb')
msgImage = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage.add_header('Content-ID', '<image1>')
msg.attach(msgImage)

server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server_ssl.ehlo()
server_ssl.login('Enter your Email here', 'Enter your Email password here')

server_ssl.sendmail('Enter your Email here', 'Enter your Email here', msg.as_string())

server_ssl.close()
print('successfully sent the mail')
#-----------------------------------------------------------------------------------------------------------------