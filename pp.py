import argparse
from office365api import Mail
from datetime import datetime, timedelta

MAIL_ID = 'qa.ex@office365.ecknhhk.xyz'
MAIL_PASS = 'ew68I7W52p*W'
F_TIME = ('{date:%Y-%m-%dT%H:%M:%SZ}'.format(date=(datetime.now()-timedelta(days=0))))


def get_args():
    parser = argparse.ArgumentParser(
        description='Perception Point HW by Uri Dar')
    parser.add_argument(
        '-p', '--path', type=str, help='Path', required=False, default='')
    args = parser.parse_args()
    return args.path


def retrieve_emails(mail):
    """ retrieve messages by time filter """
    emails = mail.inbox.get_messages(order_by='DateTimeReceived desc', filters="DateTimeReceived gt " + F_TIME)
    if len(emails) < 1:
        emails = mail.inbox.get_messages(order_by='DateTimeReceived desc')[:10]
    return emails

if __name__ == '__main__':
    path = get_args()  # retrieve args
    mail = Mail(auth=(MAIL_ID, MAIL_PASS))  # initiate mail instance
    messages = retrieve_emails(mail)

    for message in messages:
        curr_email = path + '/' + message.data['Message']['Id'] + '.txt'
        with open(curr_email, 'wb') as f:
            for field in message.data['Message']:
                f.write(str(message.data))

