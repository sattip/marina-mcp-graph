export interface GraphConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  defaultFromEmail: string;
}

export interface SendEmailParams {
  from?: string;
  to: string | string[];
  subject: string;
  body: string;
  cc?: string | string[];
  bcc?: string | string[];
  contentType?: "Text" | "HTML";
}

export interface ListEmailsParams {
  mailbox: string;
  folder?: string;
  top?: number;
  filter?: string;
}

export interface ReadEmailParams {
  mailbox: string;
  messageId: string;
}

export interface ReplyEmailParams {
  mailbox: string;
  messageId: string;
  body: string;
}

export interface ListEventsParams {
  mailbox: string;
  start: string;
  end: string;
  top?: number;
}

export interface CreateEventParams {
  mailbox: string;
  subject: string;
  start: string;
  end: string;
  attendees?: string[];
  body?: string;
  location?: string;
}
