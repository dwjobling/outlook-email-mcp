import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { EmailRequest } from "../types.js";

/**
 * MCP tool definition for sending emails via Outlook
 */
export const sendEmailTool: Tool = {
  name: "send_email",
  description:
    "Send an email via Microsoft Outlook. Requires Outlook to be installed and running on the system.",
  inputSchema: {
    type: "object" as const,
    properties: {
      to: {
        type: "string",
        description:
          "Email recipient(s). Multiple recipients can be separated by commas (e.g., 'user1@example.com, user2@example.com')",
      },
      subject: {
        type: "string",
        description: "Email subject line",
      },
      body: {
        type: "string",
        description: "Email body content",
      },
      bodyFormat: {
        type: "string",
        enum: ["html", "text"],
        description:
          "Format of the email body. Default is 'text'. Use 'html' for formatted content.",
        default: "text",
      },
      cc: {
        type: "string",
        description:
          "Carbon copy recipient(s). Multiple recipients can be separated by commas.",
      },
      bcc: {
        type: "string",
        description:
          "Blind carbon copy recipient(s). Multiple recipients can be separated by commas.",
      },
      attachments: {
        type: "array",
        items: {
          type: "string",
        },
        description:
          "Array of file paths to attach to the email. Paths must be absolute or relative to the working directory.",
      },
      importance: {
        type: "string",
        enum: ["low", "normal", "high"],
        description: "Set the importance/priority level of the email",
        default: "normal",
      },
      categories: {
        type: "string",
        description:
          "Email categories. Multiple categories can be separated by commas.",
      },
    },
    required: ["to", "subject", "body"],
  },
};

/**
 * Parse and validate email request from tool input
 */
export function parseEmailRequest(input: unknown): EmailRequest {
  if (typeof input !== "object" || input === null) {
    throw new Error("Invalid email request: input must be an object");
  }

  const data = input as Record<string, unknown>;

  // Validate required fields
  if (typeof data.to !== "string") {
    throw new Error("Invalid email request: 'to' field must be a string");
  }

  if (typeof data.subject !== "string") {
    throw new Error("Invalid email request: 'subject' field must be a string");
  }

  if (typeof data.body !== "string") {
    throw new Error("Invalid email request: 'body' field must be a string");
  }

  // Build request
  const request: EmailRequest = {
    to: data.to,
    subject: data.subject,
    body: data.body,
  };

  // Optional fields
  if (data.cc !== undefined && data.cc !== null) {
    if (typeof data.cc !== "string") {
      throw new Error("Invalid email request: 'cc' field must be a string");
    }
    request.cc = data.cc;
  }

  if (data.bcc !== undefined && data.bcc !== null) {
    if (typeof data.bcc !== "string") {
      throw new Error("Invalid email request: 'bcc' field must be a string");
    }
    request.bcc = data.bcc;
  }

  if (data.bodyFormat !== undefined && data.bodyFormat !== null) {
    if (typeof data.bodyFormat !== "string") {
      throw new Error(
        "Invalid email request: 'bodyFormat' field must be a string"
      );
    }
    if (data.bodyFormat === "html" || data.bodyFormat === "text") {
      request.bodyFormat = data.bodyFormat;
    } else {
      throw new Error(
        "Invalid email request: 'bodyFormat' must be 'html' or 'text'"
      );
    }
  }

  if (data.importance !== undefined && data.importance !== null) {
    if (typeof data.importance !== "string") {
      throw new Error(
        "Invalid email request: 'importance' field must be a string"
      );
    }
    if (
      data.importance === "low" ||
      data.importance === "normal" ||
      data.importance === "high"
    ) {
      request.importance = data.importance;
    } else {
      throw new Error(
        "Invalid email request: 'importance' must be 'low', 'normal', or 'high'"
      );
    }
  }

  if (data.categories !== undefined && data.categories !== null) {
    if (typeof data.categories !== "string") {
      throw new Error(
        "Invalid email request: 'categories' field must be a string"
      );
    }
    request.categories = data.categories;
  }

  if (data.attachments !== undefined && data.attachments !== null) {
    if (!Array.isArray(data.attachments)) {
      throw new Error(
        "Invalid email request: 'attachments' field must be an array"
      );
    }
    const attachments: string[] = [];
    for (const attachment of data.attachments) {
      if (typeof attachment !== "string") {
        throw new Error(
          "Invalid email request: each attachment must be a string (file path)"
        );
      }
      attachments.push(attachment);
    }
    request.attachments = attachments;
  }

  return request;
}
