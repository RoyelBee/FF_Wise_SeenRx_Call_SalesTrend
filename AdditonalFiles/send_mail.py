import xlrd
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from email import encoders
from email.mime.base import MIMEBase
import pandas as pd

# import Path as dir

import Files.sales_table_data as sales_table
import Files.yesterday_seen_rx as seenrx
import Files.emr_table_data as emrdata


def send_report(rsm, email):
    # import Files.banner as b
    # b.banner(rsm)
    import AdditonalFiles.generate_raw_data as row_data

    row_data.seen_rx_data()
    row_data.doctor_call_data()
    row_data.sales_trend_data()

    # ------------ Group email ----------------------------------------
    msgRoot = MIMEMultipart('related')
    me = 'erp-bi.service@transcombd.com'

    # to = [email, '']
    # cc = ['tafsir.bashar@skf.transcombd.com', '']
    # bcc = ['yakub@transcombd.com', 'rejaul.islam@transcombd.com']

    # #  --------  For testing purpose mail --------------
    email = 'rejaul.islam@transcombd.com'
    to = [email, '']
    cc = ['', '']
    bcc = ['', '']

    recipient = to + cc + bcc
    print('mail Sending to = ', email)

    subject = "SK+F Turkish Performance"

    email_server_host = 'mail.transcombd.com'
    port = 25

    msgRoot['From'] = me

    msgRoot['To'] = ', '.join(to)
    msgRoot['Cc'] = ', '.join(cc)
    msgRoot['Bcc'] = ', '.join(bcc)
    msgRoot['Subject'] = subject

    msg = MIMEMultipart()
    msgRoot.attach(msg)

    # img = open("./Images/Banner/banner_ai_" + str(rsm) + ".png", 'rb')
    # img1 = MIMEImage(img.read())
    # img.close()
    # img1.add_header('Content-ID', '<img1>')
    # msgRoot.attach(img1)
    #
    # img = open("./Images/image_1_" + str(rsm) + ".png", 'rb')
    # img2 = MIMEImage(img.read())
    # img.close()
    # img2.add_header('Content-ID', '<img2>')
    # msgRoot.attach(img2)

    # < img src = "cid:img1" >
    msgText = MIMEText("""
                         
                    
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <style>
        table {
            border-collapse: collapse;
            border: 1px solid black;
            padding: 1px;
            table-layout: fixed;
            width: 655px;
        }

        td {
            padding-top: 3px;
            border-bottom: 1px solid black;
            text-align: center;
            white-space: nowrap;
            border: 1px solid black;
            overflow: hidden;

        }

        th.unit {
            padding: 3px;
            border: 1px solid black;
            background-color: #87CEFA;
            width: 1px;
            font-size: 14px;
            white-space: nowrap;
        }

        td.idcol {
            text-align: center;
            border: 1px solid black;
            width: 1px;
            white-space: nowrap;
            overflow: hidden;
        }
        </style>
    </head>

    <br>

    <body
            <table>
            <tr><th colspan='5' style=" background-color: #b2ff66; font-size: 20px; "> Sales Trend % </th></tr>
            </table>
                        <table>""" +  + """</table>
    </body>

    <br>

    <body
            <table>
            <tr><th colspan='5' style=" background-color: #b2ff66; font-size: 20px; "> Yesterday Seen Rx </th></tr>
            </table>
                        <table>""" +  + """</table>
    </body>

    <br>

    <body
            <table>
            <tr><th colspan='5' style=" background-color: #b2ff66; font-size: 20px; ">Yesterday EMR SHARE </th></tr>
            </table>
                        <table>""" +  + """</table>
    </body>

    </br>

    </br>
    </html>
    <br><p style="text-align:left;">If there is any inconvenience, you are requested to communicate with the ERP BI Service:<br>
                         <b>(Mobile: 01713-389972, 01713-380502)</b><br><br>
                         <b>Regards</b><br>
                         <b>ERP BI Service</b><br>
                         <b>Information System Automation (ISA)</b><br><br>
                         <p style="color:blue">****This is a system generated report ****</p>
    """, 'html')

    msg.attach(msgText)


    # file_location = r"./Data/EMR_" + str(rsm) + ".xlsx"
    # filename = os.path.basename(file_location)
    # attachment = open(file_location, "rb")
    # part_2 = MIMEBase('application', 'octet-stream')
    # part_2.set_payload(attachment.read())
    # encoders.encode_base64(part_2)
    # part_2.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # msgRoot.attach(part_2)

    # # ----------- Finally send mail and close server connection ---
    server = smtplib.SMTP(email_server_host, port)
    server.ehlo()
    server.sendmail(me, recipient, msgRoot.as_string())
    server.close()
    print('Mail Send -------- \n')
