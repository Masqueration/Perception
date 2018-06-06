import argparse
from office365api import Mail
from datetime import datetime, timedelta
import boto3
import pickle
import webbrowser
from os import remove
from report import base, end_html, add_cell

MAIL_ID = 'qa.ex@office365.ecknhhk.xyz'
MAIL_PASS = 'ew68I7W52p*W'
F_TIME = ('{date:%Y-%m-%dT%H:%M:%SZ}'.format(date=(datetime.now()-timedelta(days=0))))
bucket_name = 'interview-exercises'

def get_args():
    parser = argparse.ArgumentParser(
        description='Perception Point HW by Uri Dar')
    parser.add_argument(
        '-p', '--path', type=str, help='Path', required=False, default='')
    parser.add_argument('-u', '--upload', action='store_true')
    parser.add_argument('-r', '--report', action='store_true')
    args = parser.parse_args()
    return args.path, args.upload, args.report


def retrieve_emails(mail):
    """ retrieve messages by time filter """
    emails = mail.inbox.get_messages(order_by='DateTimeReceived desc', filters="DateTimeReceived gt " + F_TIME)
    if len(emails) < 1:
        emails = mail.inbox.get_messages(order_by='DateTimeReceived desc')[:10]
    print 'Retrieved ' + str(len(emails)) + ' Emails'
    return emails


def save_locally(path, messages):
    print 'Storing emails in: ' + str(path)
    for message in messages:
        curr_email = path + '/' + message.data['Message']['Id'] + '.pkl'
        with open(curr_email, 'wb') as f:
            pickle.dump(message.data['Message'], f, pickle.HIGHEST_PROTOCOL)


def upload_to_s3(messages):
    print ('Uploading ' + str(len(messages)) + ' into S3 bucket: ' + bucket_name)
    s3 = boto3.client('s3')
    for message in messages:
        curr_email = message.data['Message']['Id'] + '.pkl'
        with open(curr_email, 'wb') as f:
            pickle.dump(message.data['Message'], f, pickle.HIGHEST_PROTOCOL)
        s3.put_object(Body=curr_email, Bucket=bucket_name, Key=curr_email)
        remove(curr_email)


def save_emails(upload, path, messages):
    if upload:
        upload_to_s3(messages)
    else:
        save_locally(path, messages)


def gen_report(messages, upload):
    name = str(datetime.now()).split('.')[0].replace(':', '-')
    html = base(name)
    if upload:
        html += """<th>S3 Link</th>"""
    for message in messages:
        html += """</tr><tr>"""
        html += add_cell(message.data['Message']['DateTimeReceived'])
        html += add_cell(message.data['Message']['From']['EmailAddress']['Address'])
        html += add_cell(message.data['Message']['ToRecipients'][0]['EmailAddress']['Address'])
        html += add_cell(message.data['Message']['Subject'])
        if upload:
            bucket_location = boto3.client('s3').get_bucket_location(Bucket=bucket_name)
            curr_email = message.data['Message']['Id'] + '.pkl'
            object_url = "https://s3-{0}.amazonaws.com/{1}/{2}".format(
                bucket_location['LocationConstraint'], bucket_name, curr_email)
            html += add_cell(object_url, link=True)
    html += end_html()
    with open('Report ' + name + '.html', 'wb') as f:
        f.write(html)
    return name

if __name__ == '__main__':
    path, upload, report = get_args()  # get_args()  # retrieve args
    mail = Mail(auth=(MAIL_ID, MAIL_PASS))  # initiate mail instance
    messages = retrieve_emails(mail)  # Lists emails
    save_emails(upload, path, messages)
    if report:
        report = gen_report(messages, upload)
        webbrowser.open('Report ' + report + '.html')
