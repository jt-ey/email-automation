import pandas as pd
import win32com.client as win32
import json
import re
import os 

PATH = os.getcwd()
workbook_name = '\sample.xlsx' # Name of your excel file, in current working dir
with open('email_text.json') as f: # Load email content data as dictionary
    text_dict = json.load(f)
    

def create_mail(text, recipient_name, data, subject, recipient_email, attachment = False,send=False,):
    """
    @params:
    """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.CC = 'email@email.com'
    mail.Subject = subject
    recipient_name = recipient_name.replace('-',' ').title() 
    text = text.replace('{name}',recipient_name) # Add name to Email body
    
    if attachment:
        attachment_path = PATH+r'\attachments\{}'.format(data)
        mail.HtmlBody = text.replace('{name}',recipient_name)
        mail.Attachments.Add(attachment_path)
    else:
        data = data.style.set_properties(**{'background-color': '#ffffbf'}, subset=['Service Account Still Required? (Y/N)',
                                             'Updated Owner Name',
                                             'Remarks (if any)']).hide_index().render(index=False)
        
        ## Remove Null values
        data = data.replace('nan','')
        
        ## Styling
        data = data.replace('<table','<table border = 1 class="dataframe"')
        
        mail.HtmlBody = text.replace('{table}',data)
    if send:
        mail.send()
    else:
        mail.save()

def get_email_content(attachment,email_type):
    if attachment:
        d = text_dict[email_type]['attachment']
    else:
        d = text_dict[email_type]['no-attachment']
    body = d['body']
    subject = d['subject']
    return body, subject



df = pd.read_excel(PATH+workbook_name)
email_list = df.Email.unique().tolist() # All people that need to be email (no duplicates)

for email in email_list:
    data = df[df.Email == email] # Select all rows in data to be sent to this person
    name = email.split("@")[0].replace(".","-") # Get name from email address
    
    
    email_type_list = data.Type.unique()
    for email_type in email_type_list: # Split the types of emails to send (types of Orphan Accounts)
        table = data[data.Type == email_type][['Domain',
                                             'SamAccountName',
                                             'DisplayName',
                                             'Description',
                                             'Previous Owner employee ID',
                                             'Previous Owner Name',
                                             'Service Account Still Required? (Y/N)',
                                             'Updated Owner Name',
                                             'Remarks (if any)']]
        n_rows = table.shape[0]
        if n_rows > 5:
            """
            If more than 5 rows, create an attachment file
            """
            
            body, subject = get_email_content(True,email_type)
            
            ## name of attachment
            attachment_name = "{}_{}.xlsx".format(email_type,name)
            table.to_excel('attachments/{}'.format(attachment_name))
            create_mail(body,name, attachment_name, subject, email, attachment = True)
            
        elif n_rows > 0: # Just precautionary measure incase there are empty DFs
            """
            Otherwise, just keep it as a table
            """
            body, subject = get_email_content(False,email_type)
            create_mail(body, name,table, subject, email)
