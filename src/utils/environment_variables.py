import os

def get_ms_teams_client_id() -> str:
    return os.getenv("MS_TEAMS_CLIENT_ID")

def get_ms_teams_client_secret() -> str:
    return os.getenv("MS_TEAMS_CLIENT_SECRET")

def get_ms_teams_tenant_id() -> str:
    return os.getenv("MS_TEAMS_TENANT_ID")

def get_ms_teams_chat_id() -> str:
    return os.getenv("MS_TEAMS_CHAT_ID")
