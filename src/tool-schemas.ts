import { z } from 'zod';

interface ToolSchemaOverride {
  description: string;
  schema?: Record<string, z.ZodType<unknown>>;
  transform?: (params: Record<string, unknown>) => unknown;
  queryTransform?: (params: Record<string, unknown>) => Record<string, string>;
}

function parseRecipients(value: unknown): { emailAddress: { address: string; name?: string } }[] {
  if (!value || typeof value !== 'string') return [];
  return value
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((addr) => ({ emailAddress: { address: addr } }));
}

const DEFAULT_MAIL_SELECT =
  'id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments';

function buildMailQueryParams(p: Record<string, unknown>): Record<string, string> {
  const params: Record<string, string> = {};
  if (p.search) params['$search'] = `"${p.search}"`;
  const filters: string[] = [];
  if (p.from) filters.push(`from/emailAddress/address eq '${p.from}'`);
  if (p.subject) filters.push(`contains(subject, '${p.subject}')`);
  if (p.unreadOnly) filters.push('isRead eq false');
  if (filters.length > 0) params['$filter'] = filters.join(' and ');
  params['$top'] = String(p.count || 10);
  params['$orderby'] = 'receivedDateTime desc';
  params['$select'] = DEFAULT_MAIL_SELECT;
  return params;
}

const toolSchemaOverrides = new Map<string, ToolSchemaOverride>();

// ── Mail (Read) ─────────────────────────────────────────────────────────────────

const mailListSchema: Record<string, z.ZodType<unknown>> = {
  search: z.string().optional().describe('Search text to find in emails'),
  from: z.string().optional().describe('Filter by sender email address'),
  subject: z.string().optional().describe('Filter by subject text'),
  unreadOnly: z.boolean().optional().describe('Only return unread emails'),
  count: z.number().optional().describe('Number of emails to return (default: 10)'),
};

toolSchemaOverrides.set('list-mail-messages', {
  description: 'List emails from your mailbox. Can search and filter by sender, subject, or read status.',
  schema: mailListSchema,
  queryTransform: buildMailQueryParams,
});

toolSchemaOverrides.set('list-mail-folder-messages', {
  description:
    'List emails in a specific mail folder. Use list-mail-folders first to get folder IDs.',
  schema: mailListSchema,
  queryTransform: buildMailQueryParams,
});

toolSchemaOverrides.set('list-shared-mailbox-messages', {
  description:
    'List emails from a shared mailbox. Provide the user-id of the shared mailbox owner.',
  schema: mailListSchema,
  queryTransform: buildMailQueryParams,
});

toolSchemaOverrides.set('list-shared-mailbox-folder-messages', {
  description:
    'List emails in a specific folder of a shared mailbox. Provide user-id and mailFolder-id.',
  schema: mailListSchema,
  queryTransform: buildMailQueryParams,
});

toolSchemaOverrides.set('get-mail-message', {
  description:
    'Get a specific email by its message ID. Returns full email details including body, recipients, and attachments.',
});

toolSchemaOverrides.set('get-shared-mailbox-message', {
  description:
    'Get a specific email from a shared mailbox by user-id and message-id.',
});

toolSchemaOverrides.set('list-mail-folders', {
  description: 'List your mail folders (inbox, sent items, drafts, etc.). Returns folder names and IDs.',
});

toolSchemaOverrides.set('list-mail-attachments', {
  description: 'List all attachments on a specific email. Provide the message-id.',
});

toolSchemaOverrides.set('get-mail-attachment', {
  description: 'Get a specific attachment from an email. Provide message-id and attachment-id.',
});

// ── Mail (Write) ────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('send-mail', {
  description: 'Send an email. Provide recipients, subject, and content.',
  schema: {
    to: z.string().describe('Comma-separated recipient email addresses'),
    subject: z.string().describe('Email subject line'),
    content: z.string().describe('Email body content'),
    cc: z.string().optional().describe('Comma-separated CC email addresses'),
    bcc: z.string().optional().describe('Comma-separated BCC email addresses'),
    isHtml: z.boolean().optional().describe('Set true if content is HTML (default: plain text)'),
  },
  transform: (p) => ({
    message: {
      subject: p.subject,
      body: { contentType: p.isHtml ? 'HTML' : 'Text', content: p.content },
      toRecipients: parseRecipients(p.to),
      ...(p.cc ? { ccRecipients: parseRecipients(p.cc) } : {}),
      ...(p.bcc ? { bccRecipients: parseRecipients(p.bcc) } : {}),
    },
    saveToSentItems: true,
  }),
});

toolSchemaOverrides.set('send-shared-mailbox-mail', {
  description:
    'Send an email from a shared mailbox. Provide user-id (path), recipients, subject, and content.',
  schema: {
    to: z.string().describe('Comma-separated recipient email addresses'),
    subject: z.string().describe('Email subject line'),
    content: z.string().describe('Email body content'),
    cc: z.string().optional().describe('Comma-separated CC email addresses'),
    isHtml: z.boolean().optional().describe('Set true if content is HTML (default: plain text)'),
  },
  transform: (p) => ({
    message: {
      subject: p.subject,
      body: { contentType: p.isHtml ? 'HTML' : 'Text', content: p.content },
      toRecipients: parseRecipients(p.to),
      ...(p.cc ? { ccRecipients: parseRecipients(p.cc) } : {}),
    },
    saveToSentItems: true,
  }),
});

toolSchemaOverrides.set('create-draft-email', {
  description: 'Create a draft email that can be edited and sent later.',
  schema: {
    subject: z.string().describe('Email subject line'),
    content: z.string().describe('Email body content'),
    to: z.string().optional().describe('Comma-separated recipient email addresses'),
    isHtml: z.boolean().optional().describe('Set true if content is HTML (default: plain text)'),
  },
  transform: (p) => ({
    subject: p.subject,
    body: { contentType: p.isHtml ? 'HTML' : 'Text', content: p.content },
    ...(p.to ? { toRecipients: parseRecipients(p.to) } : {}),
  }),
});

toolSchemaOverrides.set('move-mail-message', {
  description:
    'Move an email to a folder. Use list-mail-folders to find folder IDs. Common names: "inbox", "drafts", "deleteditems", "sentitems".',
  schema: {
    destinationId: z.string().describe('Destination folder ID or well-known name'),
  },
  transform: (p) => ({ destinationId: p.destinationId }),
});

toolSchemaOverrides.set('add-mail-attachment', {
  description: 'Add a file attachment to a draft email message.',
  schema: {
    name: z.string().describe('File name, e.g. "report.pdf"'),
    contentBytes: z.string().describe('Base64-encoded file content'),
    contentType: z.string().optional().describe('MIME type, e.g. "application/pdf"'),
  },
  transform: (p) => ({
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: p.name,
    contentBytes: p.contentBytes,
    ...(p.contentType ? { contentType: p.contentType } : {}),
  }),
});

// ── Calendar ────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-calendar-event', {
  description:
    'Create a calendar event. Provide subject, start/end datetimes. Dates should be ISO 8601 format like "2025-03-15T09:00:00".',
  schema: {
    subject: z.string().describe('Event title'),
    startDateTime: z.string().describe('Start date/time in ISO 8601, e.g. "2025-03-15T09:00:00"'),
    endDateTime: z.string().describe('End date/time in ISO 8601, e.g. "2025-03-15T10:00:00"'),
    timeZone: z.string().optional().describe('IANA time zone, e.g. "America/New_York" (default: "UTC")'),
    location: z.string().optional().describe('Location name'),
    attendees: z.string().optional().describe('Comma-separated attendee email addresses'),
    body: z.string().optional().describe('Event description/notes'),
    isOnlineMeeting: z.boolean().optional().describe('Create as Teams online meeting'),
  },
  transform: (p) => ({
    subject: p.subject,
    start: { dateTime: p.startDateTime, timeZone: p.timeZone || 'UTC' },
    end: { dateTime: p.endDateTime, timeZone: p.timeZone || 'UTC' },
    ...(p.location ? { location: { displayName: p.location } } : {}),
    ...(p.attendees
      ? {
          attendees: parseRecipients(p.attendees).map((r) => ({
            emailAddress: r.emailAddress,
            type: 'required',
          })),
        }
      : {}),
    ...(p.body ? { body: { contentType: 'Text', content: p.body } } : {}),
    ...(p.isOnlineMeeting
      ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' }
      : {}),
  }),
});

toolSchemaOverrides.set('update-calendar-event', {
  description: 'Update a calendar event. Only provide the fields you want to change.',
  schema: {
    subject: z.string().optional().describe('New event title'),
    startDateTime: z.string().optional().describe('New start date/time in ISO 8601'),
    endDateTime: z.string().optional().describe('New end date/time in ISO 8601'),
    timeZone: z.string().optional().describe('IANA time zone (default: "UTC")'),
    location: z.string().optional().describe('New location name'),
    body: z.string().optional().describe('New event description'),
  },
  transform: (p) => ({
    ...(p.subject ? { subject: p.subject } : {}),
    ...(p.startDateTime
      ? { start: { dateTime: p.startDateTime, timeZone: p.timeZone || 'UTC' } }
      : {}),
    ...(p.endDateTime
      ? { end: { dateTime: p.endDateTime, timeZone: p.timeZone || 'UTC' } }
      : {}),
    ...(p.location ? { location: { displayName: p.location } } : {}),
    ...(p.body ? { body: { contentType: 'Text', content: p.body } } : {}),
  }),
});

// ── To-Do Tasks ─────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-todo-task', {
  description:
    'Create a To-Do task. Requires a title. Use list-todo-task-lists to get the task list ID.',
  schema: {
    title: z.string().describe('Task title'),
    dueDate: z.string().optional().describe('Due date in YYYY-MM-DD format, e.g. "2025-03-15"'),
    notes: z.string().optional().describe('Task notes/details'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
  },
  transform: (p) => ({
    title: p.title,
    ...(p.dueDate
      ? { dueDateTime: { dateTime: `${p.dueDate}T00:00:00`, timeZone: 'UTC' } }
      : {}),
    ...(p.notes ? { body: { content: p.notes as string, contentType: 'text' } } : {}),
    ...(p.importance ? { importance: p.importance } : {}),
  }),
});

toolSchemaOverrides.set('update-todo-task', {
  description: 'Update a To-Do task. Only provide the fields you want to change.',
  schema: {
    title: z.string().optional().describe('New task title'),
    status: z
      .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
      .optional()
      .describe('Task status'),
    dueDate: z.string().optional().describe('Due date in YYYY-MM-DD format'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
    notes: z.string().optional().describe('Task notes/details'),
  },
  transform: (p) => ({
    ...(p.title ? { title: p.title } : {}),
    ...(p.status ? { status: p.status } : {}),
    ...(p.dueDate
      ? { dueDateTime: { dateTime: `${p.dueDate}T00:00:00`, timeZone: 'UTC' } }
      : {}),
    ...(p.importance ? { importance: p.importance } : {}),
    ...(p.notes ? { body: { content: p.notes as string, contentType: 'text' } } : {}),
  }),
});

// ── Planner Tasks ───────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-planner-task', {
  description:
    'Create a Planner task. Requires planId and title. Use list-planner-tasks or get-planner-plan to find the planId.',
  schema: {
    planId: z.string().describe('Plan ID (from list-planner-tasks or get-planner-plan)'),
    title: z.string().describe('Task title'),
    bucketId: z.string().optional().describe('Bucket ID to place the task in'),
    dueDate: z.string().optional().describe('Due date in ISO 8601 format, e.g. "2025-03-15T00:00:00Z"'),
    assignedTo: z.string().optional().describe('Comma-separated user IDs to assign the task to'),
  },
  transform: (p) => ({
    planId: p.planId,
    title: p.title,
    ...(p.bucketId ? { bucketId: p.bucketId } : {}),
    ...(p.dueDate ? { dueDateTime: p.dueDate } : {}),
    ...(p.assignedTo
      ? {
          assignments: Object.fromEntries(
            (p.assignedTo as string)
              .split(',')
              .map((id) => id.trim())
              .filter(Boolean)
              .map((id) => [
                id,
                {
                  '@odata.type': '#microsoft.graph.plannerAssignment',
                  orderHint: ' !',
                },
              ])
          ),
        }
      : {}),
  }),
});

toolSchemaOverrides.set('update-planner-task', {
  description: 'Update a Planner task. Only provide the fields you want to change.',
  schema: {
    title: z.string().optional().describe('New task title'),
    percentComplete: z.number().optional().describe('Percentage complete (0-100)'),
    dueDate: z.string().optional().describe('Due date in ISO 8601 format, or empty to clear'),
    priority: z.number().optional().describe('Priority: 0=urgent, 1=important, 2=medium, 5=low'),
  },
  transform: (p) => ({
    ...(p.title ? { title: p.title } : {}),
    ...(p.percentComplete !== undefined ? { percentComplete: p.percentComplete } : {}),
    ...(p.dueDate !== undefined ? { dueDateTime: p.dueDate || null } : {}),
    ...(p.priority !== undefined ? { priority: p.priority } : {}),
  }),
});

toolSchemaOverrides.set('update-planner-task-details', {
  description: 'Update the description of a Planner task.',
  schema: {
    description: z.string().describe('Task description/notes'),
  },
  transform: (p) => ({
    description: p.description,
  }),
});

// ── Outlook Contacts ────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-outlook-contact', {
  description: 'Create an Outlook contact.',
  schema: {
    givenName: z.string().describe('First name'),
    surname: z.string().optional().describe('Last name'),
    email: z.string().optional().describe('Email address'),
    phone: z.string().optional().describe('Phone number'),
    company: z.string().optional().describe('Company name'),
    jobTitle: z.string().optional().describe('Job title'),
  },
  transform: (p) => ({
    givenName: p.givenName,
    ...(p.surname ? { surname: p.surname } : {}),
    ...(p.email ? { emailAddresses: [{ address: p.email, name: '' }] } : {}),
    ...(p.phone ? { businessPhones: [p.phone] } : {}),
    ...(p.company ? { companyName: p.company } : {}),
    ...(p.jobTitle ? { jobTitle: p.jobTitle } : {}),
  }),
});

toolSchemaOverrides.set('update-outlook-contact', {
  description: 'Update an Outlook contact. Only provide the fields you want to change.',
  schema: {
    givenName: z.string().optional().describe('First name'),
    surname: z.string().optional().describe('Last name'),
    email: z.string().optional().describe('Email address'),
    phone: z.string().optional().describe('Phone number'),
    company: z.string().optional().describe('Company name'),
    jobTitle: z.string().optional().describe('Job title'),
  },
  transform: (p) => ({
    ...(p.givenName ? { givenName: p.givenName } : {}),
    ...(p.surname ? { surname: p.surname } : {}),
    ...(p.email ? { emailAddresses: [{ address: p.email, name: '' }] } : {}),
    ...(p.phone ? { businessPhones: [p.phone] } : {}),
    ...(p.company ? { companyName: p.company } : {}),
    ...(p.jobTitle ? { jobTitle: p.jobTitle } : {}),
  }),
});

// ── Teams / Chat Messages ───────────────────────────────────────────────────────

toolSchemaOverrides.set('send-chat-message', {
  description: 'Send a message in a Teams chat.',
  schema: {
    content: z.string().describe('Message text'),
  },
  transform: (p) => ({
    body: { content: p.content },
  }),
});

toolSchemaOverrides.set('send-channel-message', {
  description: 'Send a message to a Teams channel.',
  schema: {
    content: z.string().describe('Message text'),
  },
  transform: (p) => ({
    body: { content: p.content },
  }),
});

toolSchemaOverrides.set('reply-to-chat-message', {
  description: 'Reply to a message in a Teams chat.',
  schema: {
    content: z.string().describe('Reply text'),
  },
  transform: (p) => ({
    body: { content: p.content },
  }),
});

// ── OneNote ─────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-onenote-page', {
  description: 'Create a OneNote page with a title and HTML content.',
  schema: {
    title: z.string().describe('Page title'),
    content: z.string().describe('Page content (plain text or HTML)'),
  },
  transform: (p) => ({
    contentType: 'text/html',
    content: `<html><head><title>${p.title}</title></head><body>${p.content}</body></html>`,
  }),
});

// ── Excel ───────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-excel-chart', {
  description: 'Create a chart in an Excel worksheet.',
  schema: {
    type: z
      .string()
      .describe('Chart type: "ColumnClustered", "Pie", "Line", "Bar", "Area", "XYScatter"'),
    sourceData: z.string().describe('Data range, e.g. "A1:B5"'),
    seriesBy: z.enum(['Auto', 'Columns', 'Rows']).optional().describe('Data series orientation (default: "Auto")'),
  },
  transform: (p) => ({
    type: p.type,
    sourceData: p.sourceData,
    seriesBy: p.seriesBy || 'Auto',
  }),
});

toolSchemaOverrides.set('format-excel-range', {
  description: 'Format cells in an Excel worksheet. Provide a JSON formatting object.',
  schema: {
    bold: z.boolean().optional().describe('Make text bold'),
    italic: z.boolean().optional().describe('Make text italic'),
    fontSize: z.number().optional().describe('Font size in points'),
    fontColor: z.string().optional().describe('Font color hex, e.g. "#FF0000"'),
    fillColor: z.string().optional().describe('Background color hex, e.g. "#FFFF00"'),
    numberFormat: z.string().optional().describe('Number format, e.g. "$#,##0.00", "0%"'),
  },
  transform: (p) => {
    const result: Record<string, unknown> = {};
    if (p.bold !== undefined || p.italic !== undefined || p.fontSize || p.fontColor) {
      result.font = {
        ...(p.bold !== undefined ? { bold: p.bold } : {}),
        ...(p.italic !== undefined ? { italic: p.italic } : {}),
        ...(p.fontSize ? { size: p.fontSize } : {}),
        ...(p.fontColor ? { color: p.fontColor } : {}),
      };
    }
    if (p.fillColor) {
      result.fill = { color: p.fillColor };
    }
    if (p.numberFormat) {
      result.numberFormat = p.numberFormat;
    }
    return result;
  },
});

toolSchemaOverrides.set('sort-excel-range', {
  description: 'Sort a range of cells in an Excel worksheet.',
  schema: {
    columnIndex: z.number().describe('Column index to sort by (0-based)'),
    ascending: z.boolean().optional().describe('Sort ascending (default: true)'),
    hasHeaders: z.boolean().optional().describe('Range has a header row (default: false)'),
  },
  transform: (p) => ({
    fields: [{ key: p.columnIndex, ascending: p.ascending !== false }],
    ...(p.hasHeaders !== undefined ? { hasHeaders: p.hasHeaders } : {}),
  }),
});

// ── OneDrive Upload ─────────────────────────────────────────────────────────────

toolSchemaOverrides.set('upload-file-content', {
  description: 'Upload or replace file content in OneDrive.',
  schema: {
    content: z.string().describe('File content (text or base64 for binary)'),
  },
  transform: (p) => p.content,
});

// ── Microsoft Search ────────────────────────────────────────────────────────────

toolSchemaOverrides.set('search-query', {
  description:
    'Search across Microsoft 365: emails, calendar, files, SharePoint, Teams. Provide a query and what to search.',
  schema: {
    query: z.string().describe('Search text, e.g. "budget report", "from:user@example.com"'),
    entityTypes: z
      .string()
      .describe(
        'Comma-separated types to search: "message" (email), "event" (calendar), "driveItem" (files), "chatMessage" (Teams), "site" (SharePoint), "person"'
      ),
    size: z.number().optional().describe('Number of results (default: 25)'),
  },
  transform: (p) => ({
    requests: [
      {
        entityTypes: (p.entityTypes as string)
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean),
        query: { queryString: p.query },
        ...(p.size ? { size: p.size } : {}),
      },
    ],
  }),
});

export { toolSchemaOverrides };
export type { ToolSchemaOverride };
