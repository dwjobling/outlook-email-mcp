# Outlook Email MCP

An MCP (Model Context Protocol) server that enables sending emails via Microsoft Outlook COM interop on Windows.

## Features

- 📧 Send emails through Microsoft Outlook installed on Windows
- 📎 Support for attachments (multiple files)
- 👥 CC and BCC recipients
- 🎨 HTML and plain text body formats
- ⚠️ Email importance levels (low, normal, high)
- 🏷️ Outlook categories support
- ✅ Comprehensive email validation
- 🔒 No authentication required (uses current Outlook user account)

## Prerequisites

- **Windows OS** (7, 8, 10, 11 or Server editions)
- **Microsoft Outlook** (2016, 2019, 2021, or Microsoft 365) installed and configured
- **Node.js** 18+ (for running as MCP server)

## Installation

```bash
npm install outlook-email-mcp
```

Or for development:

```bash
git clone https://github.com/dwjobling/outlook-email-mcp.git
cd outlook-email-mcp
npm install
npm run build
```

## Usage

### GitHub Copilot Integration

Use this MCP server directly with GitHub Copilot in VS Code for natural language email sending:

**Quick Setup:**
```bash
npm install
npm run build
```

Then open the project in VS Code with GitHub Copilot installed. See [COPILOT_SETUP.md](COPILOT_SETUP.md) for detailed instructions.

**Example in Copilot Chat:**
> "Send an HTML email to alice@example.com and bob@example.com with CC to manager@example.com about the project status, set as high priority"

Copilot will use the `send_email` tool to execute your request!

#### Setup Files for Copilot
- **`.vscode/settings.json`** - VS Code configuration
- **`.vscode/mcp.json`** - MCP server configuration for Copilot
- **`COPILOT_SETUP.md`** - Complete setup and troubleshooting guide

### As an MCP Server

The package can be run as a standalone MCP server via stdio:

```bash
npm start
```

This will start the MCP server and listen for tool calls via standard input/output.

### MCP Tool: `send_email`

Send an email with the following parameters:

#### Parameters

- **`to`** (string, required): Email recipient(s). Multiple recipients can be comma-separated
  - Example: `"user@example.com"` or `"user1@example.com, user2@example.com"`

- **`subject`** (string, required): Email subject line

- **`body`** (string, required): Email body content

- **`bodyFormat`** (string, optional): Email body format
  - Options: `"text"` (default), `"html"`
  - Use `"html"` for formatted emails with CSS, tables, etc.

- **`cc`** (string, optional): Carbon copy recipient(s), comma-separated

- **`bcc`** (string, optional): Blind carbon copy recipient(s), comma-separated

- **`attachments`** (array of strings, optional): File paths to attach
  - Example: `["/path/to/file.pdf", "C:\\Users\\Name\\Document.docx"]`
  - Paths must be absolute or relative to the current working directory

- **`importance`** (string, optional): Email priority level
  - Options: `"low"`, `"normal"` (default), `"high"`

- **`categories`** (string, optional): Outlook categories, comma-separated
  - Example: `"Work, Urgent"`

#### Response

Success response:
```json
{
  "success": true,
  "message": "Email sent successfully to user@example.com",
  "messageId": "[outlook-entry-id]"
}
```

Error response:
```json
{
  "success": false,
  "message": "Failed to send email",
  "error": "Invalid email address: not-an-email"
}
```

### Example Usage (via Node.js)

```javascript
import { OutlookClient } from "outlook-email-mcp";

const client = new OutlookClient();

// Send a simple email
const response = await client.sendEmail({
  to: "recipient@example.com",
  subject: "Hello from Outlook MCP",
  body: "This is a test email.",
});

console.log(response);

// Send an HTML email with attachments
const response = await client.sendEmail({
  to: "recipient@example.com",
  subject: "Report",
  body: "<h1>Monthly Report</h1><p>Here is your report.</p>",
  bodyFormat: "html",
  cc: "manager@example.com",
  attachments: ["./report.pdf"],
  importance: "high",
});

console.log(response);
```

## Development

### Build

```bash
npm run build
```

### Watch Mode

```bash
npm run dev
```

## Error Handling

The server handles various error conditions:

- **Outlook not installed**: Cannot create Outlook.Application COM object
- **Outlook not running**: May need Outlook to be open for sending
- **Invalid email addresses**: Email validation against standard format
- **Attachment file not found**: Each attachment path is verified
- **COM errors**: Outlook-specific errors (network issues, permission errors, etc.)

All errors return structured error responses through the MCP protocol.

## Limitations

1. **Windows Only**: Requires Windows OS with Outlook installed via COM interop
2. **Outlook Installation**: Outlook must be installed and properly configured
3. **Outlook Access**: May require Outlook to be running or properly licensed
4. **File Paths**: Attachment paths must be accessible from the Node.js process
5. **Security Warnings**: Outlook may show security prompts for programmatic email access (depending on version)

## Troubleshooting

### "Failed to initialize Outlook: Cannot create Outlook.Application COM object"

- Ensure Microsoft Outlook is installed on your system
- Verify Outlook installation is not corrupted by trying to open Outlook manually
- Check that you have the necessary permissions to access COM objects

### "Attachment file not found"

- Verify file paths are absolute or relative to the current working directory
- Check file permissions - the process must have read access to the file

### Emails not appearing in Outlook

- Check Outlook is running (may need to be open)
- Verify the email addresses are valid and formatted correctly
- Check Outlook for any error dialogs or security prompts

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.
