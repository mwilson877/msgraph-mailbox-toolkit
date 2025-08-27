# Microsoft Graph Mailbox Toolkit

A comprehensive Python library for programmatic access to Exchange Online mailboxes through the Microsoft Graph API. This toolkit provides a complete set of functions for authentication, folder enumeration, message retrieval with filtering, detailed message analysis, and attachment extraction - ideal for security investigations, email forensics, and mailbox automation.

## Features

- **Authentication**: OAuth2 client credentials flow with Microsoft Graph API
- **Folder Management**: Retrieve and enumerate mailbox folders with name-to-ID mapping
- **Message Operations**: Search, filter, and retrieve messages with customizable field selection
- **Attachment Handling**: Extract attachment metadata and binary content
- **Forensics Ready**: Comprehensive field retrieval for security investigations
- **Production Ready**: Proper error handling and validation

## Prerequisites

- Python 3.6+
- Azure AD application registration with appropriate Microsoft Graph API permissions
- Exchange Online mailbox access

### Required Azure AD Permissions

Your Azure AD application needs the following Microsoft Graph API permissions:
- `Mail.Read` or `Mail.ReadWrite` (Application permissions)
- `User.Read.All` (to access user mailboxes)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/mwilson877/msgraph-mailbox-toolkit.git
cd msgraph-mailbox-toolkit
```

2. Install required dependencies:
```bash
pip install requests
```

## Quick Start

```python
from msgraph_mailbox_toolkit import get_access_token, get_folders, get_messages

# Authenticate
token = get_access_token(
    tenant_id="your-tenant-id",
    client_id="your-client-id",
    client_secret="your-client-secret"
)

# Get mailbox folders
folders = get_folders("user@example.com", token)
print(f"Available folders: {list(folders.keys())}")

# Get messages from Inbox
inbox_id = folders["Inbox"]
messages = get_messages("user@example.com", inbox_id, token, top=50)
print(f"Found {len(messages['value'])} messages")
```

## API Reference

### Authentication

#### `get_access_token(tenant_id, client_id, client_secret)`
Authenticates to Microsoft Graph API and returns an access token.

**Parameters:**
- `tenant_id` (str): Azure AD tenant ID
- `client_id` (str): Application (client) ID
- `client_secret` (str): Client secret for authentication

**Returns:** Access token string

### Folder Operations

#### `get_folders(mailbox, access_token)`
Retrieves all folders from a mailbox as a dictionary.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `access_token` (str): Bearer token for authentication

**Returns:** Dictionary with folder display names as keys and folder IDs as values

#### `get_folder_id(mailbox, folder_id_filter, access_token)`
Retrieves a specific folder ID using OData filter query.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id_filter` (str): OData filter expression (e.g., `"displayName eq 'Sent Items'"`)
- `access_token` (str): Bearer token for authentication

**Returns:** Folder ID string

### Message Operations

#### `get_messages(mailbox, folder_id, access_token, filter_query=None, top=100)`
Retrieves messages from a specific folder.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id` (str): ID of the folder containing messages
- `access_token` (str): Bearer token for authentication
- `filter_query` (str, optional): OData filter expression for message filtering
- `top` (int): Maximum number of messages to retrieve (default: 100)

**Returns:** Dictionary containing message data from API response

#### `get_message_id(mailbox, folder_id, message_id_filter, access_token)`
Finds a specific message ID using OData filter.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id` (str): ID of the folder containing messages
- `message_id_filter` (str): OData filter expression to find the message
- `access_token` (str): Bearer token for authentication

**Returns:** Message ID string

#### `get_message_details(mailbox, folder_id, message_id, access_token, select_fields=None)`
Retrieves detailed information about a specific message.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id` (str): ID of the folder containing the message
- `message_id` (str): ID of the message to retrieve
- `access_token` (str): Bearer token for authentication
- `select_fields` (list, optional): Specific fields to retrieve (defaults to comprehensive field set)

**Returns:** Dictionary containing detailed message information

### Attachment Operations

#### `get_attachment_id(mailbox, folder_id, message_id, access_token)`
Retrieves the attachment ID from the first attachment of a message.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id` (str): ID of the folder containing the message
- `message_id` (str): ID of the message containing the attachment
- `access_token` (str): Bearer token for authentication

**Returns:** Attachment ID string

#### `get_attachment_data(mailbox, folder_id, message_id, attachment_id, access_token)`
Retrieves attachment metadata and binary content.

**Parameters:**
- `mailbox` (str): Email address of the target mailbox
- `folder_id` (str): ID of the folder containing the message
- `message_id` (str): ID of the message containing the attachment
- `attachment_id` (str): ID of the attachment to retrieve
- `access_token` (str): Bearer token for authentication

**Returns:** Dictionary with `name`, `contentType`, and `content` (raw bytes)

## Usage Examples

### Basic Message Search
```python
# Search for messages from a specific sender
filter_query = "from/emailAddress/address eq 'sender@example.com'"
messages = get_messages("user@example.com", folder_id, token, filter_query=filter_query)
```

### Download Attachments
```python
# Get message with attachments
message_id = get_message_id(mailbox, folder_id, "hasAttachments eq true", token)
attachment_id = get_attachment_id(mailbox, folder_id, message_id, token)
attachment_data = get_attachment_data(mailbox, folder_id, message_id, attachment_id, token)

# Save attachment to file
with open(attachment_data['name'], 'wb') as f:
    f.write(attachment_data['content'])
```

### Custom Field Selection
```python
# Retrieve only specific message fields for better performance
custom_fields = ['id', 'subject', 'from', 'receivedDateTime']
message_details = get_message_details(mailbox, folder_id, message_id, token, select_fields=custom_fields)
```

## OData Filter Examples

Common filter patterns for message searches:

```python
# Messages from specific sender
"from/emailAddress/address eq 'sender@example.com'"

# Messages received in last 24 hours
"receivedDateTime gt 2025-01-01T00:00:00Z"

# Messages with specific subject
"contains(subject, 'urgent')"

# Messages with attachments
"hasAttachments eq true"

# Unread messages
"isRead eq false"
```

## Error Handling

All functions use proper HTTP status code validation. Common exceptions:
- `ValueError`: Invalid parameters or no results found
- `requests.HTTPError`: API communication errors
- `requests.ConnectionError`: Network connectivity issues

## Security Considerations

- **Never hardcode credentials** in your source code
- Store sensitive information (tenant ID, client secrets) in environment variables or secure configuration
- Use least-privilege principle when configuring Azure AD permissions
- Implement proper token refresh logic for long-running applications
- Be mindful of API rate limits and implement appropriate backoff strategies

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- Check the [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/api/overview) for API reference

## Changelog

### v1.0.0
- Initial release with complete mailbox access functionality
- Authentication, folder, message, and attachment operations
- Comprehensive error handling and validation
- Production-ready codebase with full documentation
