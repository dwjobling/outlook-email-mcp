import { describe, it, expect, beforeAll, afterAll } from "@jest/globals";
import { OutlookClient } from "../src/outlook-client";
import { EmailRequest } from "../src/types";
import { parseEmailRequest } from "../src/tools/send-email";

describe("OutlookClient", () => {
  let client: OutlookClient;

  beforeAll(async () => {
    client = new OutlookClient();
  });

  afterAll(async () => {
    await client.close();
  });

  describe("sendEmail", () => {
    it("should fail with invalid recipient", async () => {
      const request: EmailRequest = {
        to: "invalid-email",
        subject: "Test Subject",
        body: "Test Body",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Invalid email address");
    });

    it("should fail with empty subject", async () => {
      const request: EmailRequest = {
        to: "test@example.com",
        subject: "",
        body: "Test Body",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Subject field is required");
    });

    it("should fail with empty body", async () => {
      const request: EmailRequest = {
        to: "test@example.com",
        subject: "Test Subject",
        body: "",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Body field is required");
    });

    it("should fail with multiple invalid recipients", async () => {
      const request: EmailRequest = {
        to: "test@example.com, not-an-email",
        subject: "Test Subject",
        body: "Test Body",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Invalid email address");
    });

    it("should fail with invalid CC", async () => {
      const request: EmailRequest = {
        to: "test@example.com",
        subject: "Test Subject",
        body: "Test Body",
        cc: "invalid-cc",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Invalid CC email address");
    });

    it("should fail with invalid BCC", async () => {
      const request: EmailRequest = {
        to: "test@example.com",
        subject: "Test Subject",
        body: "Test Body",
        bcc: "invalid-bcc",
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("Invalid BCC email address");
    });

    it("should fail with non-existent attachment", async () => {
      const request: EmailRequest = {
        to: "test@example.com",
        subject: "Test Subject",
        body: "Test Body",
        attachments: ["/non/existent/file.txt"],
      };

      const response = await client.sendEmail(request);

      expect(response.success).toBe(false);
      expect(response.error).toContain("File not found");
    });

    // Note: Actual email sending tests require Outlook to be installed and running
    // These are integration tests that would need to be run in an appropriate environment
    it("should provide version info", async () => {
      const version = await client.getVersion();
      expect(typeof version).toBe("string");
    });
  });

  describe("parseEmailRequest", () => {
    it("should parse valid email request", () => {
      const request = parseEmailRequest({
        to: "test@example.com",
        subject: "Test",
        body: "Test body",
        cc: "cc@example.com",
        bcc: "bcc@example.com",
        bodyFormat: "html",
        importance: "high",
        categories: "Work",
        attachments: ["file.txt"],
      });

      expect(request.to).toBe("test@example.com");
      expect(request.subject).toBe("Test");
      expect(request.bodyFormat).toBe("html");
      expect(request.attachments).toEqual(["file.txt"]);
    });
  });
});
