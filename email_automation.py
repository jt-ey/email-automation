import pandas as pd
import win32com.client as win32
import json
import re
import os 

PATH = os.getcwd()
workbook_name = '\sample.xlsx' # Name of your excel file, in current working dir
with open('email_text.json') as f: # Load email content data as dictionary
    text_dict = json.load(f)
    
CCs = ['janelle.tang@au.ey.com', # CCed Recipients
      'janelle.tang@au.ey.com'] 


def convert_list_to_recipients(lst):
    result = ""
    for email in lst:
        result = result + email + ';'
    return result

def create_mail(text, recipient_name, data, subject, recipient_email, location = None, app = None, attachment = False,send=False,):
    """
    @params:
    
    """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.CC = convert_list_to_recipients(CCs)
    mail.Subject = subject
    
    # fill in name
    recipient_name = recipient_name.replace('-',' ').title() 
    text = text.replace('{name}',recipient_name) 
    
    # fill in date
    future_date = (dt.datetime.now().date()+dt.timedelta(days=3)).strftime("%d-%b-%Y")
    text = text.replace('{future date}', future_date)
    
#     # fill in location
#     if app:
#         text = text.replace('{application}')
    
#     # fill in app
#     if location:
#         text = text.replace('{location}')
    
    if attachment:
        attachment_path = PATH+r'\attachments\{}'.format(data)
        mail.HtmlBody = text.replace('{name}',recipient_name)
        mail.Attachments.Add(attachment_path)
    else:
        
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
        
        ## Highlighted columns for recipient to fill - in the email_text.JSON
        required_cols = text_dict[email_type]['attachment']['columns']
        required_cols = required_cols.split(", ")
        
        select_cols = ['Domain',
                     'SamAccountName',
                     'DisplayName',
                     'Description',
                     'Previous Owner employee ID',
                     'Previous Owner Name'] + required_cols
        
        
        table = data[data.Type == email_type][select_cols]
        n_rows = table.shape[0]
        
        # Styling 
        table.reset_index(drop=True)
        table = table.style.set_properties(**{'background-color': '#ffffbf','border':' 1pt solid #808080'}, subset=required_cols)
        
        if n_rows > 5:
            """
            If more than 5 rows, create an attachment file
            """
            
            body, subject = get_email_content(True,email_type)
            
            ## name of attachment
            attachment_name = "{}_{}.xlsx".format(email_type,name)
            table.to_excel('attachments/{}'.format(attachment_name),index=False)
            
            create_mail(body,name, attachment_name, subject, email, attachment = True)
            
        elif n_rows > 0:
            """
            Otherwise, just keep it as a table
            """
            html = table.render(index=False)
            body, subject = get_email_content(False,email_type)
            create_mail(body, name,html, subject, email)
