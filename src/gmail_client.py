import os

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SECRETS_DIR = os.path.join(PROJECT_ROOT, 'secrets')
CREDENTIALS_FILE = os.path.join(SECRETS_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(SECRETS_DIR, 'token.json')


def get_gmail_service(secrets_dir=None):
    secrets_dir = secrets_dir or SECRETS_DIR
    credentials_file = os.path.join(secrets_dir, 'credentials.json')
    token_file = os.path.join(secrets_dir, 'token.json')

    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except RefreshError:
                # Refresh token itself is dead (expired/revoked) -- fall
                # through to a fresh interactive login instead of crashing.
                creds = None
        if not refreshed:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, 'w') as f:
            f.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)
