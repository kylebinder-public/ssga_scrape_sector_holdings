from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
#from smtplib import SMTP
import smtplib
#import sys
import os
import pandas as pd
import datetime


def send_mail_gmail(username, password, toaddrs_list,
                    msg_text, fromaddr=None, subject="Test mail",
                    attachment_path_list=None):
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login(username, password)
    msg = MIMEMultipart()
    sender = fromaddr
    recipients = toaddrs_list
    msg['Subject'] = subject
    if fromaddr is not None:
        msg['From'] = sender
    msg['To'] = ", ".join(recipients)
    if attachment_path_list is not None:
        os.chdir(attachment_path_list)
        files = os.listdir()
        for f in files:  # add files to the message
            try:
                file_path = os.path.join(attachment_path_list, f)
                attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
                attachment.add_header('Content-Disposition', 'attachment', filename=f)
                msg.attach(attachment)
            except:
                print("could not attach file")
    msg.attach(MIMEText(msg_text, 'html'))
    s.sendmail(sender, recipients, msg.as_string())

##########################################################
# STEP 1: GET ALL HOLDINGS FROM SSGA FUNDS:
##########################################################

# Full URL example: https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlf.xlsx

XLB_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlb.xlsx"
xlb_df = pd.read_excel(XLB_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLC_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlc.xlsx"
xlc_df = pd.read_excel(XLC_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLE_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xle.xlsx"
xle_df = pd.read_excel(XLE_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLF_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlf.xlsx"
xlf_df = pd.read_excel(XLF_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLI_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xli.xlsx"
xli_df = pd.read_excel(XLI_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)
    
XLK_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlk.xlsx"
xlk_df = pd.read_excel(XLK_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLP_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlp.xlsx"
xlp_df = pd.read_excel(XLP_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLRE_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlre.xlsx"
xlre_df = pd.read_excel(XLRE_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLU_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlu.xlsx"
xlu_df = pd.read_excel(XLU_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLV_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xlv.xlsx"
xlv_df = pd.read_excel(XLV_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)

XLY_URL="https://www.ssga.com/us/en/intermediary/etfs/library-content/products/fund-data/etfs/us/holdings-daily-us-en-xly.xlsx"
xly_df = pd.read_excel(XLY_URL, header=None,
                       sheet_name='holdings',
                       nrows=17)


##########################################################
# STEP 2: CONCATENATE ALL DATAFRAMES FROM STEP 1 AS SINGLE HTML:
##########################################################

msg_text_01 = \
    str('<br>') + \
    str('<br>') + \
    str('XLB Holdings: ') + str(xlb_df.iloc[2,1]) + \
    str('<br>') + str(xlb_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlb_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLC Holdings: ') + str(xlc_df.iloc[2,1]) + \
    str('<br>') + str(xlc_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlc_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLE Holdings: ') + str(xle_df.iloc[2,1]) + \
    str('<br>') + str(xle_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xle_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLF Holdings: ') + str(xlf_df.iloc[2,1]) + \
    str('<br>') + str(xlf_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlf_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLK Holdings: ') + str(xlk_df.iloc[2,1]) + \
    str('<br>') + str(xlk_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlk_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLP Holdings: ') + str(xlp_df.iloc[2,1]) + \
    str('<br>') + str(xlp_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlp_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLRE Holdings: ') + str(xlre_df.iloc[2,1]) + \
    str('<br>') + str(xlre_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlre_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLU Holdings: ') + str(xlu_df.iloc[2,1]) + \
    str('<br>') + str(xlu_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlu_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLV Holdings: ') + str(xlv_df.iloc[2,1]) + \
    str('<br>') + str(xlv_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xlv_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) + \
    str('<br>') + \
    str('<br>') + \
    str('XLY Holdings: ') + str(xly_df.iloc[2,1]) + \
    str('<br>') + str(xly_df.iloc[0,1]) + \
    str('<br>') + \
    str('<br>') + \
    xly_df.iloc[4:,[0,1,4,5]].to_html(index=False, header=False) 

##########################################################
# STEP 3: LOAD CREDENTIALS, SEND MESSAGE:
##########################################################
    
now = datetime.datetime.now()
dtime_string = now.strftime("%Y-%m-%d---%H-%M-%S")
subj_str = str('SSGA SECTOR HOLDINGS: ') + str(dtime_string)

send_mail_gmail(username='fake_gmail_address@gmail.com', password='pw_goes_here', \
                    toaddrs_list=['recipient_to_spam_01@gmail.com', 'recipient_to_spam_02@gmail.com'], \
                    msg_text = msg_text_01, fromaddr='PAUL ADOLF VOLCKER', \
                    subject=subj_str)




