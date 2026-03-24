import { EmailRequest, EmailResponse } from "./types.js";
import * as fs from "fs";
import * as path from "path";
import { execFile } from "child_process";

/**
 * Wrapper around Outlook COM interop for sending emails
 */
export class OutlookClient {
  private initialized: boolean = false;

  /**
   * Initialize Outlook COM connection
   */
  async initialize(): Promise<void> {
    if (this.initialized) {
      return;
    }
    // PowerShell COM interop is invoked per request, so initialization is lightweight.
    this.initialized = true;
  }

  /**
   * Send email via Outlook
   */
  async sendEmail(request: EmailRequest): Promise<EmailResponse> {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      // Validate input
      this.validateEmailRequest(request);

      // Validate attachment paths
      if (request.attachments) {
        for (const attachment of request.attachments) {
          if (!fs.existsSync(attachment)) {
            return {
              success: false,
              message: `Attachment file not found: ${attachment}`,
              error: `File not found: ${attachment}`,
            };
          }
        }
      }
      const output = await this.sendViaPowerShell(request);
      const parsed = this.parseJsonResult(output);

      return {
        success: true,
        message: `Email sent successfully to ${request.to}`,
        messageId:
          parsed && typeof parsed.messageId === "string"
            ? parsed.messageId
            : undefined,
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      console.error("Error sending email:", errorMessage);

      return {
        success: false,
        message: "Failed to send email",
        error: errorMessage,
      };
    }
  }

  /**
   * Validate email request parameters
   */
  private validateEmailRequest(request: EmailRequest): void {
    if (!request.to || request.to.trim() === "") {
      throw new Error("Recipient (to) field is required");
    }

    if (!request.subject || request.subject.trim() === "") {
      throw new Error("Subject field is required");
    }

    if (!request.body || request.body.trim() === "") {
      throw new Error("Body field is required");
    }

    // Validate email addresses
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const recipients = request.to
      .split(",")
      .map((r) => r.trim())
      .filter((r) => r.length > 0);

    if (recipients.length === 0) {
      throw new Error("At least one valid recipient is required");
    }

    for (const recipient of recipients) {
      if (!emailRegex.test(recipient)) {
        throw new Error(`Invalid email address: ${recipient}`);
      }
    }

    // Validate CC if provided
    if (request.cc) {
      const ccRecipients = request.cc
        .split(",")
        .map((r) => r.trim())
        .filter((r) => r.length > 0);

      for (const recipient of ccRecipients) {
        if (!emailRegex.test(recipient)) {
          throw new Error(`Invalid CC email address: ${recipient}`);
        }
      }
    }

    // Validate BCC if provided
    if (request.bcc) {
      const bccRecipients = request.bcc
        .split(",")
        .map((r) => r.trim())
        .filter((r) => r.length > 0);

      for (const recipient of bccRecipients) {
        if (!emailRegex.test(recipient)) {
          throw new Error(`Invalid BCC email address: ${recipient}`);
        }
      }
    }
  }

  /**
   * Get Outlook version
   */
  async getVersion(): Promise<string> {
    try {
      const output = await this.runPowerShell(`
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
Write-Output $outlook.Version
`);
      const version = output.trim();
      return version.length > 0 ? version : "Unknown";
    } catch (error) {
      return "Unknown";
    }
  }

  /**
   * Close Outlook connection
   */
  async close(): Promise<void> {
    this.initialized = false;
  }

  private async sendViaPowerShell(request: EmailRequest): Promise<string> {
    const footnoteText = "This email was sent via an MCP server tool.";
    const bodyWithFootnote =
      request.bodyFormat === "html"
        ? `${request.body}<hr><p><em>${footnoteText}</em></p>`
        : `${request.body}\n\n---\n${footnoteText}`;

    const requestForScript = {
      ...request,
      body: bodyWithFootnote,
      attachments: (request.attachments ?? []).map((attachment) =>
        path.resolve(attachment)
      ),
    };

    const requestJsonBase64 = Buffer.from(
      JSON.stringify(requestForScript),
      "utf8"
    ).toString("base64");

    const script = `
$ErrorActionPreference = 'Stop'
$requestJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('${requestJsonBase64}'))
$request = $requestJson | ConvertFrom-Json

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.To = [string]$request.to
$mail.Subject = [string]$request.subject

if ($request.cc) { $mail.CC = [string]$request.cc }
if ($request.bcc) { $mail.BCC = [string]$request.bcc }

if ($request.bodyFormat -eq 'html') {
  $mail.HTMLBody = [string]$request.body
  $mail.BodyFormat = 2
} else {
  $mail.Body = [string]$request.body
  $mail.BodyFormat = 1
}

if ($request.importance -eq 'low') {
  $mail.Importance = 0
} elseif ($request.importance -eq 'high') {
  $mail.Importance = 2
} else {
  $mail.Importance = 1
}

if ($request.categories) { $mail.Categories = [string]$request.categories }

if ($request.attachments) {
  foreach ($attachment in $request.attachments) {
    if ($attachment) {
      $mail.Attachments.Add([string]$attachment, 1) | Out-Null
    }
  }
}

$mail.Send()
$result = @{ success = $true; messageId = [string]$mail.EntryID } | ConvertTo-Json -Compress
Write-Output $result
`;

    return this.runPowerShell(script);
  }

  private async runPowerShell(script: string): Promise<string> {
    const encodedCommand = Buffer.from(script, "utf16le").toString("base64");

    return new Promise((resolve, reject) => {
      execFile(
        "powershell.exe",
        [
          "-NoProfile",
          "-NonInteractive",
          "-ExecutionPolicy",
          "Bypass",
          "-EncodedCommand",
          encodedCommand,
        ],
        { maxBuffer: 1024 * 1024 },
        (error, stdout, stderr) => {
          if (error) {
            const message = stderr?.trim() || error.message;
            reject(new Error(message));
            return;
          }
          resolve(stdout ?? "");
        }
      );
    });
  }

  private parseJsonResult(output: string): Record<string, unknown> | null {
    const lines = output
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length > 0);

    for (let i = lines.length - 1; i >= 0; i--) {
      const line = lines[i];
      if (line.startsWith("{") && line.endsWith("}")) {
        try {
          const parsed = JSON.parse(line);
          if (typeof parsed === "object" && parsed !== null) {
            return parsed as Record<string, unknown>;
          }
        } catch {
          // Ignore non-JSON lines
        }
      }
    }

    return null;
  }
}
