import { GraphClient } from "../graph-client.js";
import type { SendEmailParams, ListEmailsParams, ReadEmailParams, ReplyEmailParams } from "../types.js";

function toArray(val: string | string[] | undefined): string[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

function toRecipients(emails: string[]) {
  return emails.map((e) => ({ emailAddress: { address: e.trim() } }));
}

export async function sendEmail(client: GraphClient, params: SendEmailParams): Promise<string> {
  const from = params.from || client.defaultFromEmail;
  const toList = toArray(params.to);

  if (toList.length === 0) throw new Error("'to' is required");

  await client.request("POST", `/users/${from}/sendMail`, {
    message: {
      subject: params.subject,
      body: {
        contentType: params.contentType || "HTML",
        content: params.body,
      },
      toRecipients: toRecipients(toList),
      ccRecipients: toRecipients(toArray(params.cc)),
      bccRecipients: toRecipients(toArray(params.bcc)),
    },
  });

  return `Email sent from ${from} to ${toList.join(", ")}`;
}

export async function listEmails(client: GraphClient, params: ListEmailsParams): Promise<unknown> {
  const folder = params.folder || "inbox";
  const top = params.top || 10;
  let path = `/users/${params.mailbox}/mailFolders/${folder}/messages?$top=${top}&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview`;

  if (params.filter) {
    path += `&$filter=${encodeURIComponent(params.filter)}`;
  }

  const result = (await client.request("GET", path)) as { value: unknown[] };
  return result.value;
}

export async function readEmail(client: GraphClient, params: ReadEmailParams): Promise<unknown> {
  return client.request(
    "GET",
    `/users/${params.mailbox}/messages/${params.messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments`
  );
}

export async function replyEmail(client: GraphClient, params: ReplyEmailParams): Promise<string> {
  await client.request("POST", `/users/${params.mailbox}/messages/${params.messageId}/reply`, {
    comment: params.body,
  });

  return `Reply sent for message ${params.messageId}`;
}
