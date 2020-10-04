
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import yagmail
from pathlib import Path
import os

#Path of the excel file
data = pd.read_excel (r'/mnt/c/Users/DELL/Desktop/CertificateGenerator/yes.xlsx') 

#Name and Email are the coulumn names in the Excel file
name_list = data["Name"].tolist() 
email_list = data["Email"].tolist() 

count=2

for (i,j) in zip(name_list,email_list):

    # path of the image file on which name has to be written
    im = Image.open(r'/mnt/c/Users/DELL/Desktop/CertificateGenerator/new.jpg')
    d = ImageDraw.Draw(im)

    # Location of the place where the name has to be written
    location = (740, 674)

    # Text Color
    text_color = (0, 137, 209)

    # Font file in a .ttf format with size

    font = ImageFont.truetype("arial.ttf", 80)
    d.text(location, i, fill = text_color, font = font)
    
    #Path of the folder where certificates have to be generated
    im.save('/mnt/c/Users/DELL/Desktop/CertificateGenerator/Certificates/'+str(count)+"_" +j+"_" +i + ".pdf")
    count=count+1


# SENDING THE CERTIFICATES AS ATTACHEMENTS

count=2
 
sender_email = ""  #type your email here
subject = "Type Subject of Email here"
#sender_password = input(f'Please, enter the password for {sender_email}:n')
sender_password = "" #type your passsword here


# Base folder path name
dir_name='/mnt/c/Users/DELL/Desktop/CertificateGenerator/Certificates/'

# File format of the desired certificates
filename_suffix="pdf"

for (i,j) in zip(name_list,email_list):
    receiver_emails = j 
    
    try: 
        yag = yagmail.SMTP(user=sender_email, password=sender_password)


        # Certificate Names
        base_filename = str(count)+"_" +j+"_" +i

        contents = [
        "", #type the content of the mailo here 

        os.path.join(dir_name, base_filename + "." + filename_suffix)
        #Joins the path to obtain the single path of the attachement
        ]

        yag.send(receiver_emails, subject, contents)

    except Exception as e:
        print(f'Something went wrong!e{e}')
    count=count+1



