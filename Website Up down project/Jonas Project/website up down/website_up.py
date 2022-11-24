import urllib.request as request  # used to fetch the websites content
import xlrd  # handling the excel integration
import os.path
import time  # managing the time of the script
import requests  # accessing web APIs
import smtplib  # sending mail
from email.mime.multipart import MIMEMultipart  # s.a.
from email.mime.text import MIMEText  # s.a.
from skpy import Skype  # Skype integration
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail


def makedict(path, headerrow):
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    dictionary = {sheet.cell_value(headerrow, i): [sheet.cell_value(j, i) for j in range(sheet.nrows) if
                                                   j > headerrow and sheet.cell_value(j, i) != ""] for i in
                  range(sheet.ncols) if sheet.cell_value(headerrow, i) != ""}
    dictionary["path"] = path
    dictionary["headerrow"] = headerrow
    for key in dictionary:
        if not (key == "path" or key == "headerrow"):
            dictionary[key] = [int(item) if isinstance(item, float) else item for item in dictionary[key]]
    return dictionary


def refreshdict(dictionary):
    return makedict(dictionary["path"], dictionary["headerrow"])


excel = ""
while not os.path.exists(excel):
    user_input = input("Please define path to options Excel file (leave blank to use ./options.xlsx): ")
    if user_input == "":
        excel = "./options.xlsx"
    else:
        excel = str(user_input)

print("Imported settings from Excel file")
options_dict = makedict(excel, 1)
if options_dict['log'][0] == "True":
    print("Starting...")
# User params

# log = options_dict["log"][0] == "True"
# logs all the events to the console

# sms_enabled = options_dict["sms_enabled"][0] == "True"
# skype_enabled = options_dict["skype_enabled"][0] == "True"
# whats_app_enabled = options_dict["whats_app_enabled"][0] == "True"
# choose the ways you want to be contacted

# email_contacts = options_dict["email_contacts"]
# sms_contacts = options_dict["sms_contacts"]
# whatsapp_contacts = options_dict["whatsapp_contacts"]
# skype_contacts = options_dict["skype_contacts"]
# the data of the persons to be contacted

# keywords = options_dict['keywords']
# List of keywords of which at least one must be present for the page do be fully recognized as "online"

# anti_keywords = options_dict['anti_keywords']
# List of so-to-say "anti-keywords" of which NONE must be on the page for it being recognized as "online"

# send_multiple_mails = options_dict["send_multiple_mails"][0] == "True"
# if activated, there will be an hourly update on all the websites (not recommended)

SMTP_server = None
Skype_account = None


def whatsapp(website, down_or_up):
    if options_dict["whats_app_enabled"][0] == "True":
        for contact in options_dict['whatsapp_contacts']:
            try:
                if down_or_up == 'down':
                    message = ("Attention: " + website + " Web Site is Down.")
                elif down_or_up == 'up':
                    message = ("Web Site " + website + " is up again.")
                elif down_or_up == 'hacked':
                    message = ("In Web Site " + website + " were no keywords found.")
                else:
                    message = ""
                requests.api.post(url="https://api.gupshup.io/sm/api/v1/msg",
                                  data={
                                      'apikey': options_dict['whatsapp_api_key'],
                                      'channel': 'whatsapp',
                                      'source': options_dict['whatsapp_api_telephone_number'],
                                      'destination': contact,
                                      'message': message})
            except:
                if options_dict['log'][0] == 'True':
                    print('Could not send WhatsApp message')
                continue


def skype(website, down_or_up):
    if options_dict["skype_enabled"][0] == "True":
        for skype_contact in options_dict['skype_contacts']:
            try:
                if down_or_up == 'down':
                    chat = Skype_account.contacts[skype_contact].chat
                    chat.sendMsg("Attention: " + website + " Web Site is Down.")
                elif down_or_up == 'up':
                    chat = Skype_account.contacts[skype_contact].chat
                    chat.sendMsg("Web Site " + website + " is up again.")
                elif down_or_up == 'hacked':
                    chat = Skype_account.contacts[skype_contact].chat
                    chat.sendMsg("In Web Site " + website + " were no keywords found.")
            except AttributeError:
                options_dict["skype_contacts"].remove(skype_contact)
            except:
                if options_dict['log'][0] == 'True':
                    print("Couldn't send skype message!")
                options_dict["skype_enabled"][0] = 'False'


def sms(website, down_or_up):  # SMS integration
    if options_dict["sms_enabled"][0] == "True":
        for sms_contact in options_dict["sms_contacts"]:
            try:
                if down_or_up == "down":
                    requests.api.get(
                        "http://203.212.70.200/smpp/sendsms?username=SERVER&password=server123@&to=" + str(
                            sms_contact) +
                        "&from=EDUSER&text=Attention: " + website + " Web Site is Down.")
            except:
                if options_dict['log'][0] == 'True':
                    print("Couldnt send SMS!")
                continue


def mail(website, down_or_up):  # mail integration
    if options_dict["email_enabled"][0] == "True":
        for api_key in options_dict['email_api_keys']:
            try:
                if down_or_up == 'down':
                    subject = 'Web Site down!'
                    message = "Attention: <strong>" + website + "</strong> Web Site is Down."
                elif down_or_up == 'up':
                    subject = 'Web Site back up!'
                    message = "Website <strong>" + website + "</strong> is up again!"
                elif down_or_up == 'hacked':
                    subject = 'Keywords not found!'
                    message = "In Web Site <strong>" + website + "</strong> were no keywords found."
                else:
                    subject = ''
                    message = ''
                SendGridAPIClient(api_key).send(
                    Mail(
                        from_email=options_dict['email_adress'][0],
                        to_emails=options_dict['email_contacts'],
                        subject=subject,
                        html_content=message
                    )
                )
                return
            except:
                if options_dict['log'][0] == 'True':
                    print("API Key " + api_key + " has exceeded the maximum amount of tries")
                continue


def check_if_online():
    result = False
    for item in ['https://google.com', 'http://amionline.net', 'https://duck.com']:
        try:
            if request.urlopen(item).getcode() == 200:
                result = True
        except:
            continue
    return result


print("Running...")
if __name__ == '__main__':

    down_sites = []
    key_missing_sites = []
    while True:
        if check_if_online():
            options_dict = refreshdict(options_dict)

            if options_dict["skype_enabled"][0] == "True" and not Skype_account:
                try:
                    if not os.path.exists("./tokenfile"):
                        open("tokenfile", "w").close()
                    Skype_account = Skype(options_dict['skype_username'][0], options_dict['skype_password'][0],
                                          "tokenfile")  # connecting to the Skype account
                    if options_dict["log"][0] == "True":
                        print("Successfully linked Skype account")
                except:
                    if options_dict['log'][0] == 'True':
                        print("Couldnt connect to Skype account, not sending any messages via Skype.")
                    options_dict['skype_enabled'][0] = 'False'
            if not options_dict["skype_enabled"][0] == "True":
                Skype_account = None

            for x in options_dict['websites']:
                try:
                    web_answer = request.urlopen(x)
                    s = web_answer.read().decode("utf-8")
                    keyword_condition = any([keyword in s for keyword in options_dict['keywords']]) and not any(
                        [keyword in s for
                         keyword in
                         options_dict['anti_keywords']])
                    if web_answer.getcode() == 200 and keyword_condition:
                        if x in down_sites or x in key_missing_sites:
                            if x in down_sites:
                                down_sites.remove(x)
                            if x in key_missing_sites:
                                key_missing_sites.remove(x)
                            mail(x, "up")
                            sms(x, "up")
                            whatsapp(x, "up")
                            skype(x, "up")
                        if options_dict["log"][0] == "True":
                            print("Successfully reached " + x)
                    elif ((x not in down_sites) or options_dict["send_multiple_mails"][0] == "True") and (
                            web_answer.getcode() != 200):
                        mail(x, 'down')
                        sms(x, "down")
                        whatsapp(x, "up")
                        skype(x, "down")
                        if options_dict["log"][0] == "True":
                            print(x + " is down.")
                        down_sites.append(x)
                    elif not keyword_condition and (
                            (x not in key_missing_sites) or options_dict["send_multiple_mails"][0] == "True"):
                        mail(x, "hacked")
                        sms(x, "hacked")
                        whatsapp(x, "hacked")
                        skype(x, "hacked")
                        if options_dict["log"][0] == "True":
                            print("Can't find any of the Keywords in " + x)
                        key_missing_sites.append(x)
                    del web_answer
                except:
                    if (x not in down_sites) or options_dict["send_multiple_mails"][0] == "True":
                        mail(x, 'down')
                        sms(x, "down")
                        whatsapp(x, "down")
                        skype(x, "down")
                        down_sites.append(x)
                        if options_dict["log"][0] == "True":
                            print(x + " is down.")
            if options_dict["log"][0] == "True":
                print("Sleeping for " + str(options_dict['sleep'][0]) + " seconds")
            time.sleep(options_dict['sleep'][0])
            down_sites = list(dict.fromkeys(down_sites))
