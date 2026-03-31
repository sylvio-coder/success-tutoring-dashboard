import os
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

from google_auth_oauthlib.flow import Flow

def create_flow():
    flow = Flow.from_client_secrets_file(
        "client_secrets.json",
        scopes=["openid", "email", "profile"],
        redirect_uri="http://localhost:8501",
    )
    flow.code_verifier = None
    return flow
