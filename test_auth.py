from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
from dotenv import load_dotenv

load_dotenv()
credentials = ClientCredential(os.getenv('SHP_ID_APP'), os.getenv('SHP_ID_APP_SECRET'))
ctx = ClientContext(os.getenv('SHP_SITE_URL')).with_credentials(credentials)

try:
    web = ctx.web.get().execute_query()
    print(f'Connected successfully to: {web.title}')
except Exception as e:
    print(f'Authentication failed: {e}')
