#!/usr/bin/env node

import {
  Server,
} from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { OutlookClient } from "./outlook-client.js";
import { sendEmailTool, parseEmailRequest } from "./tools/send-email.js";

/**
 * Initialize MCP server with Outlook email functionality
 */
async function main() {
  // Create server instance
  const server = new Server(
    {
      name: "outlook-email-mcp",
      version: "1.0.3",
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Initialize Outlook client (singleton instance)
  const outlookClient = new OutlookClient();

  /**
   * Handle tool listing request
   */
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [sendEmailTool],
    };
  });

  /**
   * Handle tool call request
   */
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    if (name === "send_email") {
      try {
        // Parse and validate request
        const emailRequest = parseEmailRequest(args);

        // Send email via Outlook
        const response = await outlookClient.sendEmail(emailRequest);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(response, null, 2),
            },
          ],
          isError: !response.success,
        };
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : String(error);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  success: false,
                  message: "Failed to send email",
                  error: errorMessage,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(
            {
              success: false,
              message: `Unknown tool: ${name}`,
            },
            null,
            2
          ),
        },
      ],
      isError: true,
    };
  });

  // Start server with stdio transport
  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("Outlook Email MCP server started");
  console.error("Available tools: send_email");
}

// Run server
main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
