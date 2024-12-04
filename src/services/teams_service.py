from typing import Any, Dict
import requests
import time


class TeamsService:
    """
    A service class for interacting with Microsoft Teams via the Microsoft Graph API.

    Attributes:
        client_id (str): The client ID for Microsoft Graph API authentication.
        client_secret (str): The client secret for Microsoft Graph API authentication.
        tenant_id (str): The tenant ID for Microsoft Graph API authentication.
    """

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """
        Initialize the TeamsService.

        Args:
            client_id (str): Microsoft Graph API client ID.
            client_secret (str): Microsoft Graph API client secret.
            tenant_id (str): Microsoft Graph API tenant ID.
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token: str = ""
        self.token_expiry: float = 0.0

    def _fetch_access_token(self) -> str:
        """
        Fetch a new access token from Microsoft Graph API.

        Returns:
            str: A valid access token.
        """
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        response = requests.post(url, data=data)
        response.raise_for_status()
        token_data = response.json()
        self.access_token = token_data["access_token"]
        self.token_expiry = time.time() + token_data["expires_in"] - 60  # Buffer 1 minute
        return self.access_token

    def get_access_token(self) -> str:
        """
        Get the current access token, refetching it if necessary.

        Returns:
            str: A valid access token.
        """
        if not self.access_token or time.time() >= self.token_expiry:
            self._fetch_access_token()
        return self.access_token

    def send_message(self, chat_id: str, content: str) -> None:
        """
        Send a message to a Microsoft Teams chat.

        Args:
            chat_id (str): The ID of the chat.
            content (str): The message content.

        Raises:
            requests.exceptions.HTTPError: If the API call fails.
        """
        url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
        headers = {
            "Authorization": f"Bearer {self.get_access_token()}",
            "Content-Type": "application/json",
        }
        payload = {"body": {"content": content}}
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()

    def get_recent_messages(self, chat_id: str) -> Dict[str, Any]:
        """
        Retrieve recent messages from a Microsoft Teams chat.

        Args:
            chat_id (str): The ID of the chat.

        Returns:
            Dict[str, Any]: A dictionary containing the recent messages.

        Raises:
            requests.exceptions.HTTPError: If the API call fails.
        """
        url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
        headers = {
            "Authorization": f"Bearer {self.get_access_token()}",
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
