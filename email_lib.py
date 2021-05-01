import email.utils
import pickle
import os.path
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import errors
import json

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import base64

# If modifying these scopes, delete the file token.pickle.
_SCOPES = ('https://www.googleapis.com/auth/gmail.send',)
_CREDENTIALS_FILE = 'email_secrets/discord-token.json'
TOKEN_FOLDER = 'generated_tokens'


def set_credentials_file_name(file_name):
    global _CREDENTIALS_FILE
    _CREDENTIALS_FILE = file_name


def _get_cred_id():
    with open(_CREDENTIALS_FILE, 'r') as f:
        cred_data = json.load(f)
        cred_id = cred_data.get('installed').get('client_id')
        return cred_id


def _create_pkl_token_folder():
    if not os.path.exists(TOKEN_FOLDER):
        os.mkdir(TOKEN_FOLDER)


def _get_credentials(scope):
    """Get credentials for the scope
    """
    creds = None
    # look for specific token file based on the client_id string
    cred_id = _get_cred_id()
    _create_pkl_token_folder()
    pickle_file = os.path.join(TOKEN_FOLDER, f'{cred_id}.pickle')
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(_CREDENTIALS_FILE, scope)
            creds = flow.run_local_server(open_browser=True)

        # Save the credentials for the next run
        with open(pickle_file, 'wb') as token:
            pickle.dump(creds, token)
    return creds


def add_headers(mime_message, from_e, to_e, subject_e, cc_e=None, reply_to_e=None, **kwargs):
    mime_message['to'] = to_e
    mime_message['from'] = from_e
    mime_message['subject'] = subject_e
    if cc_e:
        mime_message['cc'] = cc_e
    if reply_to_e:
        mime_message.add_header('reply-to', reply_to_e)
    return mime_message


def _convert_attachements(att_dir):
    attachements = []
    if att_dir:
        for f in os.listdir(att_dir):
            if not f.startswith('.'):
                with open(os.path.join(att_dir, f), 'rb') as file:
                    mime_app = MIMEApplication(file.read())
                    mime_app.add_header('Content-Disposition',
                                        f'attachment; filename="{os.path.basename(f)}"')
                    attachements.append(mime_app)
    return attachements


def _encode_email(mime_message):
    return {'raw': base64.urlsafe_b64encode(mime_message.as_string().encode('utf-8')).decode('utf-8')}


def plain_txt_make_body(message_e, **kwargs):
    mime_message = MIMEText(message_e)
    mime_message = add_headers(mime_message, **kwargs)
    return _encode_email(mime_message)


def fancy_email(html=None, plain=None, att_dir=None, **kwargs):
    # main mime container
    main_message = MIMEMultipart()
    main_message = add_headers(mime_message=main_message, **kwargs)

    # textual message with plain and html alternatives
    textual_message = MIMEMultipart('alternative')

    # adding plain text message
    if plain:
        mime_plain_text = MIMEText(plain, 'plain')
        textual_message.attach(mime_plain_text)

    # adding html alternative
    if html:
        mime_html = MIMEText(html, 'html')
        textual_message.attach(mime_html)

    # get the attachements in mime format
    attachements = _convert_attachements(att_dir)

    # attaching all to the main email
    main_message.attach(textual_message)

    for att in attachements:
        main_message.attach(att)

    return _encode_email(main_message)


def send_email(message, scopes=_SCOPES, verbose=False):
    # Get credentials from the specified scopes
    creds = _get_credentials(scopes)
    service = build('gmail', 'v1', credentials=creds)

    try:
        e_mail = (service.users().messages().send(userId="me", body=message).execute())
        if verbose:
            print(f"Message Id: {e_mail['id']} To : {verbose}")
        return e_mail
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def add_display_name(display_name, address):
    return email.utils.formataddr((display_name, address))


if __name__ == '__main__':
    load_dotenv()
    email_from = os.getenv('FROM')
    email_to = os.getenv('TO')
    # email_copy_to = os.getenv('COPY_TO')
    email_reply_to = os.getenv('REPLY_TO')
    email_subject = "Test of email with python"
    email_body = "Hello there from python!"

    print(f"Testing sending a simple message to {email_to}")
    message = plain_txt_make_body(from_e=add_display_name("BOSS", email_from),
                                  to_e=email_to,
                                  subject_e=email_subject,
                                  message_e=email_body,
                                  # cc_e=email_copy_to,
                                  reply_to_e=email_reply_to)
    e_mail = send_email(message=message)
