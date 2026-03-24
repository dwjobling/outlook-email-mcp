# GitHub Copilot Integration Setup

This guide will help you set up the Outlook Email MCP package as a custom tool in GitHub Copilot for VS Code.

## Prerequisites

- **GitHub Copilot** activated in VS Code (requires GitHub account with Copilot subscription)
- **Windows OS** with Microsoft Outlook installed
- **Node.js 18+** installed
- **VS Code** with latest updates

## Setup Steps

### 1. Install Dependencies

```bash
npm install
```

This installs the MCP SDK and other required packages.

### 2. Build the Project

```bash
npm run build
```

This compiles the TypeScript source code into the `dist/` directory. Verify that `dist/server.js` is created.

### 3. Verify Server Runs

Test that the MCP server can run standalone:

```bash
node dist/server.js
```

The server should output to stderr:
```
Outlook Email MCP server started
Available tools: send_email
```

Press `Ctrl+C` to stop the server.

### 4. Configure VS Code for GitHub Copilot

The project includes two configuration files that enable Copilot integration:

#### A. VS Code Settings (`.vscode/settings.json`)
Already configured with:
- Copilot enabled for all file types
- Proper formatting settings

#### B. MCP Configuration (`.vscode/mcp.json`)
This file tells GitHub Copilot about the Outlook Email MCP server:

```json
{
  "servers": {
    "outlook-email": {
      "command": "node",
      "args": ["dist/server.js"],
      "env": {}
    }
  }
}
```

### 5. Open Project in VS Code

```bash
code .
```

This opens the current directory in VS Code.

### 6. Using the Tool in Copilot

Once the project is open in VS Code, you can use the tool in Copilot Chat:

#### Via Copilot Chat
1. Open Copilot Chat (`Ctrl+Shift+I` or `Cmd+Shift+I` on Mac)
2. Type a message asking to send an email, for example:
   ```
   Send an email to john@example.com with subject "Hello" and body "This is a test email"
   ```
3. Copilot will recognize the `send_email` tool and help you compose the request
4. The tool will execute and show the result

#### Available Tool: `send_email`

**Parameters:**
- `to` (required): Email recipient(s)
- `subject` (required): Email subject
- `body` (required): Email body text
- `cc` (optional): Carbon copy recipients
- `bcc` (optional): Blind carbon copy recipients
- `bodyFormat` (optional): "html" or "text" (default: "text")
- `attachments` (optional): Array of file paths
- `importance` (optional): "low", "normal", or "high"
- `categories` (optional): Outlook categories

**Example Usage in Copilot:**
```
Use the send_email tool to send an email to alice@example.com and bob@example.com
with the subject "Team Update" and body "Please check the attached report".
Also CC manager@example.com and set importance to high.
```

Copilot will generate a tool call like:
```json
{
  "to": "alice@example.com, bob@example.com",
  "cc": "manager@example.com",
  "subject": "Team Update",
  "body": "Please check the attached report",
  "importance": "high"
}
```

## Troubleshooting

### Tool Not Available in Copilot

**Problem:** Copilot doesn't show the `send_email` tool

**Solutions:**
1. Ensure `npm run build` completed successfully
2. Verify `dist/server.js` exists
3. Restart VS Code (`Ctrl+Shift+P` → "Developer: Reload Window")
4. Check that `.vscode/mcp.json` exists and is properly formatted

### Server Won't Start

**Problem:** `node dist/server.js` fails with errors

**Solutions:**
1. Rebuild the project: `npm run build`
2. Check that dependencies are installed: `npm install`
3. Verify Node.js is in PATH: `node --version`
4. Check the console output for specific errors

### Outlook Not Found

**Problem:** Error says "Failed to initialize Outlook"

**Solutions:**
1. Verify Outlook is installed: Open Outlook manually
2. Ensure Outlook is properly configured with an email account
3. Check that no Outlook windows are in a blocked state
4. Try restarting Outlook

### Email Not Sending

**Problem:** Tool executes but email doesn't appear in Outlook

**Solutions:**
1. Verify the recipient email address is valid
2. Check Outlook's Outbox folder for stuck messages
3. Look for security prompts from Outlook (accept if prompted)
4. Review the error message returned from the tool
5. Check Outlook's error logs

## Development Mode

For active development, use watch mode to automatically rebuild:

```bash
npm run dev
```

Then in another terminal, test the server:

```bash
node dist/server.js
```

## Running Tests

To run the validation tests:

```bash
npm test
```

## Advanced Configuration

### Using with Absolute File Paths

If you want Copilot to work with attachments, you can provide absolute file paths:

```
Send email with subject "Report" and body "Attached is the monthly report"
and attachments ["/path/to/report.pdf", "C:\\Users\\Name\\Document.docx"]
```

### HTML Emails

To send formatted HTML emails:

```
Send an HTML email to recipient@example.com with subject "Newsletter"
and body "<h1>Monthly Update</h1><p>Here's what happened this month...</p>"
setting bodyFormat to "html"
```

### Multiple Recipients

Comma-separate email addresses:

```
Send email to user1@example.com, user2@example.com, user3@example.com
```

## Performance Notes

- First tool invocation may be slightly slower as Outlook initializes
- Subsequent calls are faster (Outlook session is cached)
- Attachments are validated before sending (ensures files exist)
- Email validation happens before sending (prevents invalid addresses)

## Security Considerations

- Tool runs locally on your machine with Outlook installed
- No credentials needed (uses current Outlook user)
- Files must be in accessible paths
- Outlook may show security prompts (this is normal)
- No data is sent to external servers

## Next Steps

1. **Test with simple emails:** Start with basic to/subject/body
2. **Add attachments:** Verify file paths work correctly
3. **Try formatting:** Use HTML body format for styled emails
4. **Explore features:** Test CC, BCC, importance levels, categories
5. **Integration:** Use with larger projects as needed

## Support & Issues

For issues with the MCP server itself, check `dist/server.js` and `src/` files.

For GitHub Copilot integration issues:
1. Check VS Code console (View > Output > GitHub Copilot)
2. Verify `.vscode/mcp.json` is properly formatted
3. Restart VS Code and Outlook
4. Review the troubleshooting section above
