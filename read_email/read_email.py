""" 
    Name:           read_email.py
    Author:         Ali Alhamedi
    Date:           April 21, 2021
    Description:    This module contains the read_email and save_att functions that read from Outlook email and save the attachments
                    in the path given.
"""
import win32com.client
import os
import pandas as pd
import dateutil.parser
from datetime import datetime, timedelta

def read_email(received_dt, folder=None, subject=None, email=None):
    """Return COM email object

    Args:
        received_dt (datetime): the date of the email.
        folder (String, optional): the index sub-folder. Defaults to None.
        subject (String, optional): The subject of the email. Defaults to None.
        email (String, optional): The email of the sender. Defaults to None.

    Returns:
        COM Object: COM Object that contains the emails details. (should be used with save_att func)
    """
    outlook = win32com.client.Dispatch("outlook.application")
    mapi = outlook.GetNamespace("MAPI")
    if folder != None:
        inbox = mapi.GetDefaultFolder(6).Folders(folder)
    else:
        inbox = mapi.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    # --Take the emails from specific email and subject--
    messages = messages.Restrict("[ReceivedTime] >= '" + adjust_date(received_dt) + "'")
    if email != None:
        messages = messages.Restrict("[SenderEmailAddress] = '"+ email + "'")
    if subject != None:
        messages = messages.Restrict("[Subject] = '" + subject + "'")

    return messages


def save_att(messages, path, att_name = None):
    """Takes COM object and the path to save the attachments """
    msg = 0
    files = []
    output_directory = path
    while msg < len(messages):
        message=messages[msg]
        for att in message.Attachments:
            if att_name != None:
                if att_name in att.FileName:
                    att.SaveASFile(os.path.join(output_directory, att.FileName))
                    files.append(att.FileName)
            else:
                att.SaveASFile(os.path.join(output_directory, att.FileName))
        msg+=1
    return files


def adjust_date(date):
    date_time = datetime.combine(date, datetime.min.time())
    date_time = date_time.strftime("%a %d-%b-%y %I:%M %p")
    return date_time

