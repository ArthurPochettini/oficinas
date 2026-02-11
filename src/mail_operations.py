import re, time
from datetime import datetime

import win32com.client as win32
from bs4 import BeautifulSoup, NavigableString

OUTLOOK = win32.Dispatch("outlook.application")
    
PROG = re.compile("Nova Ordem de Serviço: [0-9]{6}")

TODAY:datetime = datetime.now()

to_digest_mails = list()
digested_mails = list()

def find_mail():
    inbox = OUTLOOK.GetNamespace("MAPI").GetDefaultFolder(6).Folders("GIT")
    mails = inbox.Items
    for mail in mails:
        received_time = parse_mail_time(mail)
        if received_time < TODAY: continue
        # Redundante desde que o filtro seja aplicado no Outlook
        # if PROG.search(mail.Subject):
        #     if mail.Subject in mail_history_subject: continue
        #     # TODO implementar checagem de remetente
        #     mail_history.append(mail)
        #     mail_history_subject.append(mail.Subject)
        # else:
        #     continue
        if mail.Subject in get_subjects(digested_mails): continue
        to_digest_mails.append(mail)

def parse_mail_time(mail) -> datetime:
    return datetime.strptime(str(mail.ReceivedTime)[0:25], "%Y-%m-%d %H:%M:%S.%f")

#def get_last_mail():
#    most_recent = mail_history[0]
#    for mail in mail_history[1:]:
#        if parse_mail_time(most_recent) < parse_mail_time(mail):
#            most_recent = mail
#    return most_recent

def parse_mail(mail):
    mail_html = mail.HTMLBody
    parsed_html = BeautifulSoup(mail_html, "html.parser")
    table = parsed_html.find("table")
    tags = table.find_all()
    for tag in tags:
        for attr in ["style", "border", "cellpadding", "cellspacing", "class", "width", "colspan", "valign", "align"]:
            if tag.has_attr(attr):
                del tag.attrs[attr]

    return table

def digest_mail(mail) -> bool:
    if not (mail in digested_mails):
        digested_mails.append(mail)
        return True
    else:
        return False

def get_subjects(mails) -> list[str]:
    r_list = list()
    for mail in mails:
        r_list.append(mail.Subject)
    return r_list

if __name__ == "__main__":
    pass