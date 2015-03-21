import re
import time
import datetime

from xlwings import Workbook, Range, Sheet
import pandas as pd
import numpy as np
import xlsxwriter
import win32com.client
import os

def email_passback_templates():

    wb = Workbook.caller()

    sheet = Range('passback', 'AA1').value
    passback_folder = sheet[:sheet.rindex('\\')] + '\\Passback Templates ' + Range('passback', 'L2').value
    passback_templates = os.listdir(passback_folder)

    contacts = pd.DataFrame(Range('passback_contacts', 'A1').table.value, columns = Range('passback_contacts', 'A1').horizontal.value)
    contacts.drop(0, inplace = True)

    passback_publishers = []
    for i in passback_templates:
        passback_publishers.append(i.split('_', 1)[0])

    templates_to_send = []
    for pub in list(contacts['Publisher']):
        if pub in passback_publishers:
            templates_to_send.append(passback_folder + '\\' + pub + '_' + Range('passback', 'L1').value + ' - ' +
                                    Range('passback', 'L2').value + '.xlsx')

    i = iter(list(contacts['Contacts']))
    j = iter(templates_to_send)
    k = iter(list(contacts['Publisher']))

    mail_item = 0x0
    client_object = win32com.client.Dispatch('Outlook.Application')

    for publisher in list(contacts['Publisher']):
        mail = client_object.CreateItem(mail_item)
        mail.Subject = k.next() + ' Passback Data ' + Range('passback', 'L1').value + ' - ' + Range('passback', 'L2').value
        mail.Body = 'Passback'
        mail.To = i.next()
        attachment = j.next()
        mail.Attachments.Add(attachment)
        mail.Send()
