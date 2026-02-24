import { GraphClient } from "../graph-client.js";
import type { ListEventsParams, CreateEventParams } from "../types.js";

export async function listEvents(client: GraphClient, params: ListEventsParams): Promise<unknown> {
  const top = params.top || 20;
  const path = `/users/${params.mailbox}/calendarView?startDateTime=${encodeURIComponent(params.start)}&endDateTime=${encodeURIComponent(params.end)}&$top=${top}&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,attendees`;

  const result = (await client.request("GET", path)) as { value: unknown[] };
  return result.value;
}

export async function createEvent(client: GraphClient, params: CreateEventParams): Promise<string> {
  const event: Record<string, unknown> = {
    subject: params.subject,
    start: { dateTime: params.start, timeZone: "Europe/Athens" },
    end: { dateTime: params.end, timeZone: "Europe/Athens" },
  };

  if (params.attendees?.length) {
    event.attendees = params.attendees.map((email) => ({
      emailAddress: { address: email.trim() },
      type: "required",
    }));
  }

  if (params.body) {
    event.body = { contentType: "HTML", content: params.body };
  }

  if (params.location) {
    event.location = { displayName: params.location };
  }

  const result = (await client.request("POST", `/users/${params.mailbox}/events`, event)) as { id: string };
  return `Event created: ${params.subject} (id: ${result.id})`;
}
