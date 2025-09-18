# Author:
# Date: August 25 2025
# Description: Microsoft Graph API mailbox access toolkit for Exchange Online investigations.
#              Complete Python library providing programmatic access to Exchange Online mailboxes
#              through Microsoft Graph API. Supports authentication, folder enumeration, message
#              retrieval with filtering, detailed message analysis, and attachment extraction
#              for security investigations and email forensics.

import requests
from datetime import datetime, timedelta

# Default message fields to retrieve from Graph API for comprehensive email analysis
default_fields = [
    'id',
    'changeKey',
    'categories',
    'createDateTime',
    'lastModifiedDateTime',
    'subject',
    'body',
    'importance',
    'hasAttachments',
    'attachments',
    'parentFolderId',
    'from',
    'sender',
    'toRecipients',
    'ccRecipients',
    'bccRecipients',
    'replyTo',
    'conversationId',
    'conversationIndex',
    'receivedDateTime',
    'sentDateTime',
    'isDeliveryReceiptRequested',
    'isReadReceiptRequested',
    'isRead',
    'isDraft',
    'webLink',
    'internetMessageId',
    'internetMessageHeaders',
    'flag',
    'inferenceClassification',
    'uniqueBody',
    'singleValueExtendedProperties',
    'multiValueExtendedProperties'
]


def get_access_token(tenant_id, client_id, client_secret):
    """
    Authenticates to Microsoft Graph API using client credentials grant and returns an access token.

    :param tenant_id: The Azure AD tenant ID
    :param client_id: The application (client) ID
    :param client_secret: The client secret for authentication
    :return: Access token string for Microsoft Graph API requests
    """

    # OAuth2 token endpoint for the specified tenant
    authority_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    # Client credentials grant request payload
    payload = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }

    # Request access token from Azure AD
    response = requests.post(authority_url, data=payload)
    response.raise_for_status()

    # Extract and return the access token
    token = response.json().get('access_token')
    return token


def get_folders(mailbox, access_token):
    """
    Retrieves all folders from a specified mailbox and returns them as a dictionary.

    :param mailbox: The email address of the mailbox to query
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: Dictionary with folder display names as keys and folder IDs as values
    """
    # Graph API endpoint for mailbox folders
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()

    # Validate that we found the mailbox and it contains folders
    if not data.get('value') or len(data['value']) == 0:
        raise ValueError(f"Could not find Mailbox: {mailbox}")

    # Create dictionary mapping folder names to IDs for easy lookup
    folders = {folder['displayName']: folder['id'] for folder in data.get('value', [])}

    return folders


def get_folder_id(mailbox, folder_id_filter, access_token):
    """
    Retrieves the folder ID from a specified mailbox using a filter query.

    :param mailbox: The email address of the mailbox to query
    :param folder_id_filter: The OData filter expression to find the desired folder
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: The ID of the first folder matching the filter
    """

    # Graph API endpoint for mailbox folders
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    params = {
        '$filter': folder_id_filter
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    data = response.json()

    # Validate that we found matching folders
    if not data.get('value') or len(data['value']) == 0:
        raise ValueError(f"No folders found matching filter: {folder_id_filter}")

    folder_id = data['value'][0]['id']

    return folder_id

def get_child_folders(mailbox, main_folder_id, access_token):
    """
    Retrieves all child folders from a specified mailbox and folder and returns them as a dictionary.

    :param mailbox: The email address of the mailbox to query
    :param main_folder_id: The parent folder id to query
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: Dictionary with child folder display names as keys and child folder IDs as values
    """

    # Graph API endpoint for mailbox folders
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{main_folder_id}/childFolders"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()

    # Validate that we found matching folders
    if not data.get('value') or len(data['value']) == 0:
        raise ValueError(f"No child folders found matching under folder id: {main_folder_id}.")

    child_folders = {folder['displayName']: folder['id'] for folder in data.get('value', [])}

    return child_folders


def get_messages(mailbox, folder_id, access_token, filter_query=None, top=100):
    """
    Retrieves messages from a specified folder within a mailbox and returns them as a dictionary.

    :param mailbox: The email address of the mailbox to query
    :param folder_id: The ID of the folder containing the messages
    :param access_token: Bearer token for Microsoft Graph API authentication
    :param filter_query: Optional OData filter expression to filter messages
    :param top: Maximum number of messages to retrieve (default: 100)
    :return: Dictionary containing message data from the API response
    """


    # Graph API endpoint for messages in specific folder
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Set query parameters - always include top to avoid default limit of 10
    params = {
        '$top': top  # Default to 100 instead of API's default 10
    }

    # Add optional filter if provided
    if filter_query:
        params['$filter'] = filter_query

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    messages = response.json()

    # Validate that the folder contains messages
    if not messages.get('value') or len(messages['value']) == 0:
        raise ValueError(f"Could not find messages in folder for mailbox: {mailbox}")

    return messages


def get_message_id(mailbox, folder_id, message_id_filter, access_token):
    """
    Retrieves the message ID from a specified folder within a mailbox using a filter query.

    :param mailbox: The email address of the mailbox to query
    :param folder_id: The ID of the folder containing the messages
    :param message_id_filter: The OData filter expression to find the desired message
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: The ID of the first message matching the filter
    """

    # Graph API endpoint for messages in the specified folder
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"

    params = {
        '$filter': message_id_filter
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    data = response.json()

    # Validate that we found matching messages
    if not data.get('value') or len(data['value']) == 0:
        raise ValueError(f"No messages found matching filter: {message_id_filter}")

    message_id = data['value'][0]['id']

    return message_id


def get_message_details(mailbox, folder_id, message_id, access_token, select_fields=None):
    """
    Retrieves detailed information about a specific message and returns it as a dictionary.

    :param mailbox: The email address of the mailbox to query
    :param folder_id: The ID of the folder containing the message
    :param message_id: The ID of the message to retrieve details for
    :param access_token: Bearer token for Microsoft Graph API authentication
    :param select_fields: Optional list of specific fields to retrieve (defaults to default_fields)
    :return: Dictionary containing detailed message information
    """

    # Use default fields if none specified
    if select_fields is None:
        select_fields = default_fields

    # Graph API endpoint for specific message
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages/{message_id}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Use the provided select_fields parameter instead of hardcoded default_fields
    params = {
        '$select': ','.join(select_fields)
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    data = response.json()

    # Validate that message data was retrieved
    if not data:
        raise ValueError(f"Message metadata not found")

    message_details = data

    return message_details


def get_attachment_id(mailbox, folder_id, message_id, access_token):
    """
    Retrieves the attachment ID from the first attachment of a specified message.

    :param mailbox: The email address of the mailbox to query
    :param folder_id: The ID of the folder containing the message
    :param message_id: The ID of the message containing the attachment
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: The ID of the first attachment in the message
    """

    # Graph API endpoint for message attachments
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages/{message_id}/attachments"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()

    # Validate that message contains attachments
    if not data.get('value') or len(data['value']) == 0:
        raise ValueError(f"No attachments found in message")

    attachment_id = data['value'][0]['id']

    return attachment_id


def get_attachment_data(mailbox, folder_id, message_id, attachment_id, access_token):
    """
    Retrieves attachment metadata and raw content from a specified email attachment.

    :param mailbox: The email address of the mailbox to query
    :param folder_id: The ID of the folder containing the message
    :param message_id: The ID of the message containing the attachment
    :param attachment_id: The ID of the attachment to retrieve
    :param access_token: Bearer token for Microsoft Graph API authentication
    :return: Dictionary containing attachment name, contentType, and raw content bytes
    """

    # First, get attachment metadata to retrieve name and content type
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages/{message_id}/attachments/{attachment_id}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()

    # Validate metadata response
    if not data:
        raise ValueError(f"Attachment metadata not found")

    # Extract attachment metadata
    attachment_type = data.get('contentType')
    attachment_name = data.get('name')

    # Get attachment binary content using the $value endpoint
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages/{message_id}/attachments/{attachment_id}/$value"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    # Validate that attachment has content
    if not response.content:
        raise ValueError("Attachment has no content or is empty")

    attachment_raw = response.content

    # Return structured attachment data with metadata and content
    attachment_data = {
        'name': attachment_name,
        'contentType': attachment_type,
        'content': attachment_raw,
    }

    return attachment_data
