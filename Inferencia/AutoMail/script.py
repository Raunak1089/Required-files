import base64
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import mimetypes
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# import openpyxl as xl
#
# wb = xl.load_workbook('database.xlsx')
# sheet = wb['Sheet1']

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/']


def get_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    return service


def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id,
                                                  body=message).execute()

        # print('Message Id: {}'.format(message['id']))

        return message
    except Exception as e:
        print('An error occurred: {}'.format(e))
        return None


def create_message(
        sender,
        to,
        subject,
        message_text,
        imgName
):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    message.attach(MIMEText(message_text, 'html'))

    # Attach the image to the message
    image_path = 'qrimages/' + imgName
    with open(image_path, 'rb') as f:
        img_data = f.read()
    img_type, _ = mimetypes.guess_type(image_path)
    img_name = os.path.basename(image_path)
    img = MIMEImage(img_data, img_type)
    img.add_header('Content-ID', f'<{image_path}>')
    img.add_header('Content-Disposition', 'inline', filename=img_name)
    message.attach(img)

    raw_message = \
        base64.urlsafe_b64encode(message.as_string().encode('utf-8'))
    return {'raw': raw_message.decode('utf-8')}


# with open('signup.html', 'r') as file:
#     emailBody = file.read()
#
# for i in range(2, 4):
#     service = get_service()
#     user_id = 'me'
#     this_body = emailBody.replace('ParticipantName', sheet.cell(i, 3).value.split(' ')[0])
#     this_body = this_body.replace('imgSrc', f'cid:qrimages/{sheet.cell(i, 1).value}.png')
#     msg = create_message('yyqoqok1009@gmail.com',
#                          sheet.cell(i, 2).value,
#                          'Confirmation for Inferencia',
#                          this_body,
#                          sheet.cell(i, 1).value + '.png')
#     sm = send_message(service, user_id, msg)
#     print(f'{i-1}. Message Id: {sm["id"]}')


