# Publiposting emails

This project aims at providing boilerplate code for sending personalised emails based on a template.

It includes adding attachements, alternative html and plain text emails as well as reading infos from Google spreadsheets and sending via gmail API.

## Requirements

### Emails

You will need a secrets file for the email account you want to use to send the emails. You can generate such a file in the [Google Cloud Console](https://console.developers.google.com/apis/credentials). The file will need to be placed in the folder called `email_secrets`. 

You will then have to specify it using the `set_credentials_file_name()` function from `email_lib.py`.

The first time you will connect to the account a browser window will open and prompt you to connect to the account. This will then cache the access tokens inside a pickle file in the `generated_tokens` folder.

### Google Drive

If you wish to use a Google spreadsheet to pull iformations from you will need to add the `service_account_secrets.json` file to the `drive_secrets` folder.

## Infos

Sensible informations are stored in a `.env` to preserv details from leaking in the source code
