// Export types and client for use in other packages
export { OutlookClient } from "./outlook-client.js";
export { sendEmailTool, parseEmailRequest } from "./tools/send-email.js";
export type {
  EmailRequest,
  EmailResponse,
  OutlookConfig,
} from "./types.js";
