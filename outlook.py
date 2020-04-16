""" Recommended use: 
from robee import outlook as out"""

# Work In Progress 15.04.20

from win32com.client import Dispatch


class Outlook:

    def __init__(self):
        is_init = True

    
    def __enter__(self):
        #self.outlook = self.dispatch()
        self.dispatch()
        is_active = True
        return self

    def __exit__(self, exception_type, exception_value, traceback):
        is_active = False
        self.outlook.quit()
        print("in __exit__")

    def dispatch(self):
        #self.outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.outlook = Dispatch("Outlook.Application")
        return self.outlook

    def send_mail(self, 
                recipient, 
                subject, 
                body='', 
                cc='', 
                bcc='', 
                html_body='', 
                attachments=None,
                draft_path=None):
        #outlook = win32.Dispatch('outlook.application')
        outlook = self.outlook.GetNamespace("MAPI")
        if type(recipient) == list:
            recipient = '; '.join(recipient)
        if type(cc) == list:
            cc = '; '.join(cc)
        if type(bcc) == list:
            bcc = '; '.join(bcc)
    
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Cc = cc
        mail.Bcc = bcc
        mail.Subject = subject
        mail.Body = body
        mail.HTMLBody = html_body

        if attachments is not None:
            if type(attachments) == str:
                attachments = [attachments]
            for attachment in attachments:
                mail.Attachments.Add(attachment)
        if draft_path is not None:
            mail.SaveAs(draft_path)
        else:
            mail.Send()


    def get_mailboxes(get_objects=False): #if not objects then gets strings
        #dispatch
        # perhaps try: outlook = self.outlook
        outlook = self.outlook.GetNamespace("MAPI")
        list_mailboxes = []
        for mailbox in range(outlook.Folders.count):
            if get_objects:
                list_mailboxes.append(outlook.Folders[mailbox])
            else:
                list_mailboxes.append(str(outlook.Folders[mailbox]))
        return list_mailboxes


    def get_email(self, email_subject, folder_object):
        # or incorporate get_folder() and use folder_name instead ?
        
        target_mail = None
        for i, email in enumerate(folder_object.Items):
            if email_subject in email.Subject:
                target_mail = email
        return target_mail

    def get_emails(self, folder_object, string_list=True):
        """Uses folder object from get_folder() to list all emails from it"""
        # or incorporate get_folder() and use folder_name instead ?
        if string_list == True:
            list_emails = [str(msg) for msg in folder_object.Items]
        else:
            list_emails = [msg for msg in folder_object.Items]
        return list_emails

    def get_folder(self, folder_name):
        outlook = self.outlook.GetNamespace("MAPI")
        if type(folder_name) == str:
            folder_name = [folder_name]
        
        folder_dir = outlook.Folders(folder_name[0])
        for i, subfolder in enumerate(folder_name[1:]):
            folder_dir = folder_dir.Folders[subfolder]
        return folder_dir

    def move_email(self, subject_string, source_folder, target_folder):
        source_folder = self.get_folder(source_folder)
        target_folder = self.get_folder(target_folder)
        for i, email in enumerate(source_folder.Items):
            if subject_string in email.Subject:
                email.Move(target_folder)

    def read_email(self, email_object):
        email_content = {
            'sender': email_object.SenderEmailAddress,
            'subject': email_object.subject,
            'body': email_object.body
        }
        return email_content

    def save_attachments(self, email_object, target_path, prefix=None):
        from pathlib import Path
        list_saved = []
        for attachment in email_object.Attachments:
                save_path = Path(target_path, prefix, attachment.FileName)
                attachment.SaveAsFile(save_path)
                list_saved.append(path)
        return list_saved



