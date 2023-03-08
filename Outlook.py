import time
import win32com.client as win32
import Config as c
from pretty_html_table import build_table
# from main import event
from datetime import date, timedelta

def event():

    Date = date.today()
    CurrentDate = date.strftime(Date, "%d-%b-%Y")
    t = time.localtime()
    x = int(time.strftime("%H", t))

    if (x > 0) and (x <= 12):
        Event = 'Morning'
        return Event
    elif (x > 12) and (x < 18):
        Event = 'Afternoon'
        return Event
    else:
        Event = 'Evening'
        return Event

def mail_details(body : list):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = c.Receiver # ";".join(c.Receiver)
    mail.CC = ";".join(c.others)
    mail.Subject = "Daily Commit Tracker"
    body_content_IPC = body[0]
    body_content_RTOS = body[1]

    while (len(body_content_IPC) != 0) or (len(body_content_RTOS) != 0):
        if len(body_content_IPC) == 0:
            title_IPC = ""
        else:
            title_IPC = "VHMIPC"

        if len(body_content_RTOS) == 0:
            title_RTOS = ""
        else:
            title_RTOS = "MYRTOS"

        # mail.Body = 
        mail.HTMLBody = f"""
                        <html><body><p>Hi Ravi,</p>
                        <p>Please find the updated commit and merge status.</p>
                        <p><b>{title_IPC}</b></p>
                        {body_content_IPC}
                        <p><b>{title_RTOS}</b></p>
                        {body_content_RTOS}
                        <p> </p>
                        <p>Confluence URL : https://confluence.rampgroup.com/pages/viewpage.action?spaceKey=DEV&title=Gerrit-Bitbucket+Daily+commit+tracker</p>
                        <p>Regards,</p>
                        <p>DevOps Team</p>
                        </body></html>
                        """
            # <p>Regards,</p>
            # <p>Sourabh Gupta</p>

        # To attach a file to the email (optional):
        # attachment  = "Path to the attachment"
        # mail.Attachments.Add(attachment)
        mail.Send()
        print(f"\n7. Mail with tables has Sent  {date.today()}-{event()}!!")
        break
    else:
        print(f'<<<<No new entries found at {date.today()}-{event()}>>>>')

        mail.HTMLBody = f"""
            <html><body><p>Hi Ravi,</p>
            <p>No new commit and merge status in Gerrit from RAMP and GM side.</p>
            <p> </p>
            <p>Confluence URL : https://confluence.rampgroup.com/pages/viewpage.action?spaceKey=DEV&title=Gerrit-Bitbucket+Daily+commit+tracker</p>
            <p>Regards,</p>
            <p>DevOps Team</p>
            </body></html>
            """
        mail.Send()
        print(f"\n7. Empty Mail has Sent at {date.today()}-{event()}!!")

def send_table(data : list):
    table_data_IPC = data[0]
    table_data_RTOS = data[1]
    print("\n 6. This dataframe func has come to send_table func")
    table_IPC = build_table(table_data_IPC, 'grey_light', font_size= '10px', font_family='Open Sans, sans-serif', text_align='left', width='auto', index=False)
    table_RTOS = build_table(table_data_RTOS, 'grey_light', font_size=' 10px', font_family='Open Sans, sans-serif', text_align='left', width='auto', index=False)
    table = [table_IPC, table_RTOS]
    mail_details(table)
    return "Mail as Sent!!"


# """==================================================================This is for Daily PR mail Automation=================================================================="""


def PR_Mail():

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = c.Receiver # ";".join(c.Receiver)
    mail.CC = ";".join(c.CC_tags)
    mail.Subject = "Daily PR Report"

    mail.HTMLBody = f"""
                    <html><body><p>Hi Ravi,</p>
                    <p>Here is the PR Data till Date : {date.today()} 08:30 AM</p>
                    <p> </p>
                    <p>Daily PR Report Link : https://rampgroups-my.sharepoint.com/:f:/g/personal/atlassian_rampgroup_com/EjFQhp2qao9ImmZmfVmbnk8BQ9LlIFUi4njU_52-e9F32A?e=cg1CnF</p>
                    <p>Regards,</p>
                    <p>DevOps Team</p>
                    </body></html>
                    """
    mail.Send()
    print(f"\n\n\n8. Daily PR Report mail has been Send {date.today()}!!\n\n\n")
    print("      ================================================================>>>>>>>>-----------Daily PR------------<<<<<<<<<=============================================          ")