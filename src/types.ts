/**
 * Email request parameters for sending via Outlook
 */
export interface EmailRequest {
  to: string;
  subject: string;
  body: string;
  cc?: string;
  bcc?: string;
  bodyFormat?: "html" | "text";
  attachments?: string[];
  importance?: "low" | "normal" | "high";
  categories?: string;
}

/**
 * Response from email sending operation
 */
export interface EmailResponse {
  success: boolean;
  message: string;
  messageId?: string;
  error?: string;
}

/**
 * Outlook COM configuration
 */
export interface OutlookConfig {
  timeout?: number;
  retryAttempts?: number;
}
