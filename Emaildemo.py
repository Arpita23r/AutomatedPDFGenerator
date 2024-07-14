import smtplib
import var
sender_email = "sb7160@srmist.edu.in"

receiver_email = "saurajyotibhattacharjee64@gmail.com"


#create connection
connection = smtplib.SMTP("smtp.gmail.com", "587")

#tls
connection.starttls()
print("Connections made succesfull")
#login
connection.login(user=sender_email,password="wqlj hdrm dodw pdfz")
print("Logged in successfully")
connection.sendmail(from_addr=sender_email,to_addrs= receiver_email,msg="Welcome to SST of ONGC")
print("Email sent Successfully")
connection.close()
print("closed connections succesfully")