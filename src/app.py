import time
import threading
from collections import defaultdict
from typing import Any, Literal, Optional
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import PromptTemplate
from services.teams_service import TeamsService
from utils.environment_variables import get_ms_teams_chat_id, get_ms_teams_client_id, get_ms_teams_client_secret, get_ms_teams_tenant_id
from utils.get_llm import get_llm
from dotenv import load_dotenv

load_dotenv()

# Constants
MAX_MESSAGE_HISTORY = 20
MESSAGE_DELAY = 5  # Seconds to wait before processing messages
POLL_INTERVAL = 10  # Seconds between polls

# Initialize LLM
llm = get_llm()

# Data structures for chat history and message buffering
chat_history: dict[str, list[dict[str, str]]] = defaultdict(list)
message_buffers: dict[str, dict[str, Any]] = {}  # Format: {"chat_id": {"timer": Timer, "messages": []}}

user_names = {
    "chat_id": "first name",
}

# Functions
def generate_conversation(messages: list[dict[str, str]], user_name: str) -> str:
    """
    Generate a conversation string from a list of messages.
    """
    return '\n'.join([f"{user_name if message['user_type'] == 'user' else 'Me'}: {message['content']}" for message in messages])

def append_chat_history(chat_id: str, user_type: Literal['me', 'user'], content: str, message_id: Optional[str] = None) -> None:
    """
    Append a message to the chat history for a specific chat.

    Args:
        chat_id (str): The chat ID.
        message_id (str): The message ID.
        user_type (Literal['me', 'user']): Type of user ('me' for bot, 'user' for end user).
        content (str): Content of the message.
    """
    if len(chat_history[chat_id]) >= MAX_MESSAGE_HISTORY:
        chat_history[chat_id].pop(0)

    chat_history[chat_id].append({
        "message_id": message_id,
        "user_type": user_type,
        "content": content
    })

def generate_response(chat_id: str) -> str:
    """
    Generate a response based on the chat history for a specific chat.

    Args:
        chat_id (str): The chat ID.

    Returns:
        str: The generated response.
    """
    user_name = user_names.get(chat_id, "User")
    conversation = generate_conversation(chat_history[chat_id], user_name)

    query_output_parser = StrOutputParser()

    query_prompt = PromptTemplate(
        template="""
            You are a person who gets chat messages from your colleagues in a conversational style.
            Your tone should be witty, mischievous and incorporate jokes and humiliation to keep the interaction engaging and entertaining.
            Keep the answers short to make it look like you are busy and respond in the same language as used in the conversation (most often danish).
            Write an answer based on the conversation below, focusing on the latest messages from {user_name}:

            {conversation}
        """,
        input_variables=["conversation", "user_name"]
    )

    model = query_prompt | llm | query_output_parser

    response_text = model.invoke({
        "conversation": conversation,
        "user_name": user_name
    })

    append_chat_history(chat_id, 'me', response_text)
    return response_text

def process_buffered_messages(chat_id: str, message_id: str, teams_service: TeamsService) -> None:
    """
    Process buffered messages for a specific chat and send a response.

    Args:
        chat_id (str): The chat ID.
        teams_service (TeamsService): An instance of TeamsService.
    """
    if chat_id in message_buffers:
        buffered_messages = " ".join(message_buffers[chat_id]["messages"])

        append_chat_history(chat_id, 'user', buffered_messages, message_id)
        response = generate_response(chat_id)

        print(f"Sending response to chat {chat_id}: {response}")
        choice = input(f"Do you want to send the message? [y/N]\t")
        if choice and choice.lower() == 'y':
            teams_service.send_message(chat_id, response)

        del message_buffers[chat_id]

def buffer_message(message: str, message_id: str, chat_id: str, teams_service: TeamsService) -> None:
    """
    Buffer messages for a chat and start/reset the delay timer.

    Args:
        message (str): The content of the message.
        message_id (str): The ID of the message.
        chat_id (str): The chat ID.
        teams_service (TeamsService): An instance of TeamsService.
    """
    if chat_id not in message_buffers:
        message_buffers[chat_id] = {"timer": None, "messages": []}

    # Check if message is already processed
    if any(entry["message_id"] == message_id for entry in chat_history[chat_id]):
        return

    message_buffers[chat_id]["messages"].append(message)

    if message_buffers[chat_id]["timer"]:
        message_buffers[chat_id]["timer"].cancel()

    message_buffers[chat_id]["timer"] = threading.Timer(
        MESSAGE_DELAY, process_buffered_messages, [chat_id, message_id, teams_service]
    )
    message_buffers[chat_id]["timer"].start()

def fetch_new_messages(teams_service: TeamsService, chat_id: str) -> None:
    """
    Fetch new messages from the specified chat and process unprocessed ones.

    Args:
        teams_service (TeamsService): An instance of the TeamsService.
        chat_id (str): The ID of the chat to poll messages from.
    """
    messages = teams_service.get_recent_messages(chat_id)  # Fetch messages using TeamsService

    for message in messages.get("value", []):
        message_id = message["id"]
        content = message["body"]["content"]

        # Append the new message to the chat history and buffer it
        buffer_message(content, message_id, chat_id, teams_service)

def poll_messages(teams_service: TeamsService, chat_id: str) -> None:
    """
    Poll messages from the specified chat at regular intervals.

    Args:
        teams_service (TeamsService): An instance of the TeamsService.
        chat_id (str): The ID of the chat to poll messages from.
    """
    while True:
        fetch_new_messages(teams_service, chat_id)
        time.sleep(POLL_INTERVAL)

def main():
    # TODO support multiple chat ids
    specific_chat_id = get_ms_teams_chat_id()

    teams_service = TeamsService(
        client_id=get_ms_teams_client_id(),
        client_secret=get_ms_teams_client_secret(),
        tenant_id=get_ms_teams_tenant_id(),
    )

    print(f"Polling messages for chat {specific_chat_id} every {POLL_INTERVAL} seconds...")
    poll_messages(teams_service, specific_chat_id)

if __name__ == "__main__":
    main()
