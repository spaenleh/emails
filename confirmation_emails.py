import os
import gspread
import re

from dotenv import load_dotenv

import email_lib
import pandas as pd

# service account infos to access the sheets
SERVICE_ACCOUNT_FILE = 'drive_secrets/adele_drive_client_secret.json'
DEFAULT_EMAIL_HEADER_COL = 'Adresse e-mail'
DEFAULT_NAME_HEADER_COL = 'Nom prénom'

# regex
RE_EMAIL = r'^(\w|\.|\_|\-)+[@](\w|\_|\-|\.)+[.]\w{2,3}$'

# keys
NAME = 'name'
EMAIL = 'email'
COLOR = 'color'
SIZE = 'size'


def get_from_spreadsheet(sheet_title, sheet_index=0):
    gc = gspread.service_account(SERVICE_ACCOUNT_FILE)
    worksheet = gc.open(sheet_title).get_worksheet(sheet_index)
    df = pd.DataFrame(worksheet.get_all_records())
    return df


def rename_columns(df, mapdict):
    return df.rename(mapdict, axis=1)


def set_value(sheet_title, sheet_index=0, cell='A1', value='OK'):
    gc = gspread.service_account(SERVICE_ACCOUNT_FILE)
    worksheet = gc.open(sheet_title).get_worksheet(sheet_index)
    worksheet.update(cell, value)


def email_checker(email):
    return re.search(RE_EMAIL, email) is not None


if __name__ == '__main__':
    load_dotenv()
    # --------------------------------------------------------------------
    #                       spreadsheet settings
    # --------------------------------------------------------------------
    # Modify this to the spreadsheet you want
    # /!\ It has to be shared with the service account /!\
    GSHEET_TITLE = 'Commande Pulls de Section / Department Sweater Order  (réponses)'
    # GSHEET_TITLE = 'test_sheet_for_bot'

    # mapper is used to rename columns of the Dataframe to more readable names
    mapper = {"Quel est ton petit nom (Prénom Nom) ? -  What's your name sweety (Firstname Name)?": NAME,
              "Adresse e-mail": EMAIL,
              "Pour enfin mettre de la couleur dans notre quotidien l'ADELE te propose cette année une palette audacieuse. Fini les pulls d'hiver sombre, place au printemps et ces petits pulls qui donne du peps. Laquelle te ferais plaisir ? - To finally bring color to your life, ADELE present you this year an array of daring colour. No more dark colour, it's time to spice up your wardrobe.  Which colour will please you ?": COLOR,
              "On va pas se mentir, cet hiver on a tous abusé des raclettes et autres fondues (covid compatible selon les experts suisses ;) ). Alors en toute honnêteté à combien tu estimes ton tour de bidou ? - Let's be honest, during this winter we’ve all had our fair share of raclette and fondue (claimed covid-safe by the Swiss experts of course). So how big can you still fit in ?": SIZE,
              }

    # change this to the name of the column that determines if the email has to be sent
    CONDITION_HEADER_COL = "Mail"
    CONDITION_TO_SEND_VALUE = 'TRUE'

    # --------------------------------------------------------------------
    #                          email settings
    # --------------------------------------------------------------------
    # these are the template emails
    TXT_FILE = 'templates/confirm_paiement_email.txt'  # for the text version
    HTML_FILE = ''  # for the html version

    # this is the Subject of the email
    email_subject = "Confirmation de Paiement - Payement Confirmation"

    # "from" field
    EMAIL_FROM_DISPLAY = os.getenv('FROM_DISPLAY')
    EMAIL_FROM_ADDRESS = os.getenv('FROM')
    email_lib.set_credentials_file_name(os.getenv('EMAIL_TOKEN_FILE'))

    # "reply-to" field -- leave address empty to disable
    EMAIL_REPLY_TO_DISPLAY = os.getenv('REPLY_DISPLAY')
    EMAIL_REPLY_TO_ADDRESS = os.getenv('REPLY_TO')

    # --------------------------------------------------------------------
    #                    Confirmation
    # --------------------------------------------------------------------
    # this is to change values in the confirmation column -- no confirmation leave empty
    CONFIRMATION_COL = 'I'
    CONFIRMATION_VALUE = 'Confirmation ok'

    # --------------------------------------------------------------------
    #                    Attachements to email
    # --------------------------------------------------------------------
    # Absolute path to the attachement dir -- leave empty if no attachements
    attachement_dir = ''

    df_people = get_from_spreadsheet(GSHEET_TITLE)

    # rename columns
    df_people = rename_columns(df_people, mapper)

    # remove people that don't have valid emails
    a = df_people[EMAIL].map(email_checker)
    df_people = df_people[df_people[EMAIL].map(email_checker)]

    # remove people that don't need to receive the email
    df_people = df_people[df_people[CONDITION_HEADER_COL] == CONDITION_TO_SEND_VALUE]

    # title the names
    df_people[NAME] = df_people[NAME].map(lambda x: x.title().strip())

    # lowercase the colors
    df_people[COLOR] = df_people[COLOR].map(lambda x: x.lower().strip())

    print('\n'.join(list(df_people[EMAIL])))
    nb_to_send = len(df_people)
    if nb_to_send == 0:
        print("Nothing to send !")
        exit(0)
    validate = input(f"Send email to {nb_to_send} people ? ")
    if any([validate in v for v in ['no', 'n']]):
        print("\nok no emails were sent, exiting ...")
        exit(0)

    # add custom name display
    email_from = email_lib.add_display_name(display_name=EMAIL_FROM_DISPLAY,
                                            address=EMAIL_FROM_ADDRESS)

    reply_to_email = None
    if EMAIL_REPLY_TO_ADDRESS and EMAIL_REPLY_TO_ADDRESS != "":
        reply_to_email = email_lib.add_display_name(display_name=EMAIL_REPLY_TO_DISPLAY,
                                                    address=EMAIL_REPLY_TO_ADDRESS)

    txt_body = None
    html_body = None
    if os.path.exists(TXT_FILE):
        with open(TXT_FILE, 'r') as f:
            txt_body = f.read()
    else:
        print(f"The text alternative ({TXT_FILE}) does not exist in the working directory.")
        exit(1)
    if os.path.exists(HTML_FILE):
        with open(HTML_FILE, 'r') as f:
            html_body = f.read()
    else:
        print(f"Info: Skipping the html alternative as it does not exist in the working directory.")

    # send email to people
    for i, person in df_people.iterrows():
        # format email
        txt_body_perso = txt_body.format(person[NAME],
                                         person[COLOR],
                                         person[SIZE])
        message = email_lib.fancy_email(from_e=email_from,
                                        to_e=person.get(EMAIL),
                                        subject_e=email_subject,
                                        plain=txt_body_perso,
                                        reply_to_e=reply_to_email,
                                        att_dir=attachement_dir)
        # use the verbose as a flag to print the message id and person name
        email_lib.send_email(message=message, verbose=person.get(NAME))
        # print(txt_body_perso + '\n\n\n')

        # change value in spreadsheet
        if CONFIRMATION_COL:
            set_value(GSHEET_TITLE, cell=f'{CONFIRMATION_COL}{i + 2}', value=CONFIRMATION_VALUE)
