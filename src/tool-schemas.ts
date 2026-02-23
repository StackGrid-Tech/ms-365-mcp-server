import { z } from 'zod';

interface ToolSchemaOverride {
  description: string;
  schema?: Record<string, z.ZodType<unknown>>;
  transform?: (params: Record<string, unknown>) => unknown;
  queryTransform?: (params: Record<string, unknown>) => Record<string, string>;
  pathTransform?: (basePath: string, params: Record<string, unknown>) => string;
}

function parseRecipients(value: unknown): { emailAddress: { address: string; name?: string } }[] {
  if (!value || typeof value !== 'string') return [];
  return value
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((addr) => ({ emailAddress: { address: addr } }));
}

// ── Shared helpers ──────────────────────────────────────────────────────────────

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

function buildGenericListQuery(p: Record<string, unknown>): Record<string, string> {
  const params: Record<string, string> = {};
  if (p.search) params['$search'] = `"${p.search}"`;
  if (p.count) params['$top'] = String(p.count);
  return params;
}

const genericListSchema: Record<string, z.ZodType<unknown>> = {
  search: z.string().optional().describe('Search text'),
  count: z.number().optional().describe('Max results to return'),
};

const toolSchemaOverrides = new Map<string, ToolSchemaOverride>();

// ═════════════════════════════════════════════════════════════════════════════════
// ── MAIL ─────────────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Mail (Read) ──

toolSchemaOverrides.set('list-mail-messages', {
  description:
    'List emails. Searches your mailbox by default. Set folderId to list a specific folder, or userId to access a shared mailbox.',
  schema: {
    search: z.string().optional().describe('Search text to find in emails'),
    from: z.string().optional().describe('Filter by sender email address'),
    subject: z.string().optional().describe('Filter by subject text'),
    unreadOnly: z.boolean().optional().describe('Only return unread emails'),
    count: z.number().optional().describe('Number of emails to return (default: 10)'),
    folderId: z
      .string()
      .optional()
      .describe('Mail folder ID to list (use list-mail-folders to find IDs)'),
    userId: z
      .string()
      .optional()
      .describe('User ID or email for shared mailbox access'),
  },
  queryTransform: buildMailQueryParams,
  pathTransform: (_base, p) => {
    const root = p.userId ? `/users/${p.userId}` : '/me';
    if (p.folderId) return `${root}/mailFolders/${p.folderId}/messages`;
    return `${root}/messages`;
  },
});

toolSchemaOverrides.set('get-mail-message', {
  description:
    'Get a specific email by message-id. Set userId for shared mailbox access.',
  schema: {
    userId: z
      .string()
      .optional()
      .describe('User ID or email for shared mailbox access'),
  },
  pathTransform: (base, p) => {
    if (p.userId) return base.replace('/me/', `/users/${p.userId}/`);
    return base;
  },
});

toolSchemaOverrides.set('list-mail-folders', {
  description:
    'List your mail folders (inbox, sent items, drafts, etc.). Returns folder names and IDs.',
});

toolSchemaOverrides.set('list-mail-attachments', {
  description: 'List all attachments on a specific email. Provide the message-id.',
});

toolSchemaOverrides.set('get-mail-attachment', {
  description: 'Get a specific attachment from an email. Provide message-id and attachment-id.',
});

// ── Mail (Write) ──

toolSchemaOverrides.set('send-mail', {
  description:
    'Send an email. Set userId to send from a shared mailbox.',
  schema: {
    to: z.string().describe('Comma-separated recipient email addresses'),
    subject: z.string().describe('Email subject line'),
    content: z.string().describe('Email body content'),
    cc: z.string().optional().describe('Comma-separated CC email addresses'),
    bcc: z.string().optional().describe('Comma-separated BCC email addresses'),
    isHtml: z.boolean().optional().describe('Set true if content is HTML (default: plain text)'),
    userId: z
      .string()
      .optional()
      .describe('User ID or email to send from a shared mailbox'),
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
  pathTransform: (base, p) => {
    if (p.userId) return `/users/${p.userId}/sendMail`;
    return base;
  },
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
    'Move an email to a folder. Common folder names: "inbox", "drafts", "deleteditems", "sentitems". Use list-mail-folders to find other folder IDs.',
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

// ── Mail (Delete) ──

toolSchemaOverrides.set('delete-mail-message', {
  description: 'Delete an email. Provide the message-id of the email to delete.',
});

toolSchemaOverrides.set('delete-mail-attachment', {
  description:
    'Delete an attachment from an email. Provide message-id and attachment-id.',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── CALENDAR ─────────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Calendar (Read) ──

toolSchemaOverrides.set('list-calendar-events', {
  description:
    'List your calendar events. Can filter by date range, subject, or search text.',
  schema: {
    search: z.string().optional().describe('Search text to find in events'),
    startDate: z
      .string()
      .optional()
      .describe('Filter events starting after this date (YYYY-MM-DD)'),
    endDate: z
      .string()
      .optional()
      .describe('Filter events ending before this date (YYYY-MM-DD)'),
    count: z.number().optional().describe('Number of events to return (default: 10)'),
  },
  queryTransform: (p) => {
    const params: Record<string, string> = {};
    if (p.search) params['$search'] = `"${p.search}"`;
    const filters: string[] = [];
    if (p.startDate) filters.push(`start/dateTime ge '${p.startDate}T00:00:00Z'`);
    if (p.endDate) filters.push(`end/dateTime le '${p.endDate}T23:59:59Z'`);
    if (filters.length > 0) params['$filter'] = filters.join(' and ');
    params['$top'] = String(p.count || 10);
    params['$orderby'] = 'start/dateTime';
    params['$select'] = 'id,subject,start,end,location,organizer,isOnlineMeeting,bodyPreview';
    return params;
  },
});

toolSchemaOverrides.set('get-calendar-event', {
  description: 'Get a specific calendar event by its event-id. Returns full event details.',
});

toolSchemaOverrides.set('get-calendar-view', {
  description:
    'Get calendar events in a specific date/time range. Expands recurring events into individual occurrences.',
  schema: {
    startDateTime: z
      .string()
      .describe('Range start in ISO 8601, e.g. "2025-03-01T00:00:00Z"'),
    endDateTime: z
      .string()
      .describe('Range end in ISO 8601, e.g. "2025-03-31T23:59:59Z"'),
    count: z.number().optional().describe('Max events to return'),
  },
  queryTransform: (p) => {
    const params: Record<string, string> = {};
    params['startDateTime'] = p.startDateTime as string;
    params['endDateTime'] = p.endDateTime as string;
    if (p.count) params['$top'] = String(p.count);
    params['$select'] = 'id,subject,start,end,location,organizer,isOnlineMeeting';
    return params;
  },
});

toolSchemaOverrides.set('list-calendars', {
  description:
    'List all your calendars (primary, shared, group calendars). Returns calendar names and IDs.',
});

// ── Calendar (Write) ──

toolSchemaOverrides.set('create-calendar-event', {
  description:
    'Create a calendar event. Dates should be ISO 8601 format like "2025-03-15T09:00:00".',
  schema: {
    subject: z.string().describe('Event title'),
    startDateTime: z
      .string()
      .describe('Start date/time in ISO 8601, e.g. "2025-03-15T09:00:00"'),
    endDateTime: z
      .string()
      .describe('End date/time in ISO 8601, e.g. "2025-03-15T10:00:00"'),
    timeZone: z
      .string()
      .optional()
      .describe('IANA time zone, e.g. "America/New_York" (default: "UTC")'),
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

// ── Calendar (Delete) ──

toolSchemaOverrides.set('delete-calendar-event', {
  description: 'Delete a calendar event. Provide the event-id.',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── TO-DO TASKS ──────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── To-Do (Read) ──

toolSchemaOverrides.set('list-todo-task-lists', {
  description:
    'List your To-Do task lists. Returns list names and IDs needed for other To-Do operations.',
});

toolSchemaOverrides.set('list-todo-tasks', {
  description: 'List tasks in a To-Do task list. Provide the task list ID.',
  schema: {
    status: z
      .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
      .optional()
      .describe('Filter by task status'),
    count: z.number().optional().describe('Max tasks to return'),
  },
  queryTransform: (p) => {
    const params: Record<string, string> = {};
    if (p.status) params['$filter'] = `status eq '${p.status}'`;
    if (p.count) params['$top'] = String(p.count);
    return params;
  },
});

toolSchemaOverrides.set('get-todo-task', {
  description: 'Get a specific To-Do task by task list ID and task ID.',
});

// ── To-Do (Write) ──

toolSchemaOverrides.set('create-todo-task', {
  description:
    'Create a To-Do task. Requires a title. Use list-todo-task-lists to get the task list ID.',
  schema: {
    title: z.string().describe('Task title'),
    dueDate: z
      .string()
      .optional()
      .describe('Due date in YYYY-MM-DD format, e.g. "2025-03-15"'),
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

// ── To-Do (Delete) ──

toolSchemaOverrides.set('delete-todo-task', {
  description: 'Delete a To-Do task. Provide the task list ID and task ID.',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── PLANNER ──────────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Planner (Read) ──

toolSchemaOverrides.set('list-planner-tasks', {
  description: 'List all your Planner tasks across all plans.',
});

toolSchemaOverrides.set('get-planner-plan', {
  description: 'Get details of a specific Planner plan by its plan ID.',
});

toolSchemaOverrides.set('list-plan-tasks', {
  description: 'List all tasks in a specific Planner plan. Provide the plan ID.',
});

toolSchemaOverrides.set('get-planner-task', {
  description: 'Get details of a specific Planner task by its task ID.',
});

// ── Planner (Write) ──

toolSchemaOverrides.set('create-planner-task', {
  description:
    'Create a Planner task. Use list-planner-tasks or get-planner-plan to find the planId.',
  schema: {
    planId: z.string().describe('Plan ID (from list-planner-tasks or get-planner-plan)'),
    title: z.string().describe('Task title'),
    bucketId: z.string().optional().describe('Bucket ID to place the task in'),
    dueDate: z
      .string()
      .optional()
      .describe('Due date in ISO 8601 format, e.g. "2025-03-15T00:00:00Z"'),
    assignedTo: z
      .string()
      .optional()
      .describe('Comma-separated user IDs to assign the task to'),
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
    dueDate: z
      .string()
      .optional()
      .describe('Due date in ISO 8601 format, or empty to clear'),
    priority: z
      .number()
      .optional()
      .describe('Priority: 0=urgent, 1=important, 2=medium, 5=low'),
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

// ═════════════════════════════════════════════════════════════════════════════════
// ── OUTLOOK CONTACTS ─────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Contacts (Read) ──

toolSchemaOverrides.set('list-outlook-contacts', {
  description: 'List your Outlook contacts. Can search by name or email.',
  schema: genericListSchema,
  queryTransform: (p) => {
    const params = buildGenericListQuery(p);
    if (!p.count) params['$top'] = '25';
    params['$select'] = 'id,displayName,givenName,surname,emailAddresses,businessPhones,companyName,jobTitle';
    return params;
  },
});

toolSchemaOverrides.set('get-outlook-contact', {
  description: 'Get a specific Outlook contact by its contact-id.',
});

// ── Contacts (Write) ──

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

// ── Contacts (Delete) ──

toolSchemaOverrides.set('delete-outlook-contact', {
  description: 'Delete an Outlook contact. Provide the contact-id.',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── USER / PROFILE ───────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

toolSchemaOverrides.set('get-current-user', {
  description: 'Get your own profile info (name, email, job title, etc.).',
});

toolSchemaOverrides.set('list-users', {
  description: 'List users in your organization. Can search by name or email.',
  schema: genericListSchema,
  queryTransform: (p) => {
    const params = buildGenericListQuery(p);
    if (!p.count) params['$top'] = '25';
    params['$select'] = 'id,displayName,mail,userPrincipalName,jobTitle,department';
    return params;
  },
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── TEAMS / CHAT ─────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Teams (Read) ──

toolSchemaOverrides.set('list-joined-teams', {
  description: 'List all Teams you are a member of. Returns team names and IDs.',
});

toolSchemaOverrides.set('get-team', {
  description: 'Get details of a specific Team by its team-id.',
});

toolSchemaOverrides.set('list-team-channels', {
  description: 'List channels in a Team. Provide the team-id.',
});

toolSchemaOverrides.set('get-team-channel', {
  description: 'Get details of a specific channel. Provide team-id and channel-id.',
});

toolSchemaOverrides.set('list-team-members', {
  description: 'List members of a Team. Provide the team-id.',
});

toolSchemaOverrides.set('list-channel-messages', {
  description: 'List messages in a Teams channel. Provide team-id and channel-id.',
  schema: {
    count: z.number().optional().describe('Max messages to return'),
  },
  queryTransform: (p) => {
    const params: Record<string, string> = {};
    if (p.count) params['$top'] = String(p.count);
    return params;
  },
});

toolSchemaOverrides.set('get-channel-message', {
  description:
    'Get a specific message from a Teams channel. Provide team-id, channel-id, and message-id.',
});

// ── Chat (Read) ──

toolSchemaOverrides.set('list-chats', {
  description: 'List your Teams chats (1:1, group chats, meeting chats).',
});

toolSchemaOverrides.set('get-chat', {
  description: 'Get details of a specific Teams chat by its chat-id.',
});

toolSchemaOverrides.set('list-chat-messages', {
  description: 'List messages in a Teams chat. Provide the chat-id.',
  schema: {
    count: z.number().optional().describe('Max messages to return'),
  },
  queryTransform: (p) => {
    const params: Record<string, string> = {};
    if (p.count) params['$top'] = String(p.count);
    return params;
  },
});

toolSchemaOverrides.set('get-chat-message', {
  description: 'Get a specific message from a Teams chat. Provide chat-id and message-id.',
});

toolSchemaOverrides.set('list-chat-message-replies', {
  description:
    'List replies to a specific Teams chat message. Provide chat-id and message-id.',
});

// ── Teams/Chat (Write) ──

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

// ═════════════════════════════════════════════════════════════════════════════════
// ── ONEDRIVE / FILES ─────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── OneDrive (Read) ──

toolSchemaOverrides.set('list-drives', {
  description: 'List your OneDrive drives. Returns drive names and IDs.',
});

toolSchemaOverrides.set('get-drive-root-item', {
  description: 'Get the root item of a OneDrive drive. Provide the drive-id.',
});


toolSchemaOverrides.set('list-folder-files', {
  description:
    'List files and folders inside a OneDrive folder. Provide drive-id and folder item-id.',
});

toolSchemaOverrides.set('download-onedrive-file-content', {
  description:
    'Download file content from OneDrive. Provide drive-id, parent folder item-id, and file item-id.',
});

// ── OneDrive (Write) ──

toolSchemaOverrides.set('upload-file-content', {
  description: 'Upload or replace file content in OneDrive.',
  schema: {
    content: z.string().describe('File content (text or base64 for binary)'),
  },
  transform: (p) => p.content,
});

// ── OneDrive (Delete) ──

toolSchemaOverrides.set('delete-onedrive-file', {
  description: 'Delete a file or folder from OneDrive. Provide drive-id and item-id.',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── EXCEL ────────────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── Excel (Read) ──

toolSchemaOverrides.set('list-excel-worksheets', {
  description:
    'List worksheets in an Excel workbook. Provide drive-id and workbook item-id.',
});

toolSchemaOverrides.set('get-excel-range', {
  description:
    'Get values from a range of cells in an Excel worksheet. Provide drive-id, workbook item-id, worksheet-id, and cell range address.',
});

// ── Excel (Write) ──

toolSchemaOverrides.set('create-excel-chart', {
  description: 'Create a chart in an Excel worksheet.',
  schema: {
    type: z
      .string()
      .describe(
        'Chart type: "ColumnClustered", "Pie", "Line", "Bar", "Area", "XYScatter"'
      ),
    sourceData: z.string().describe('Data range, e.g. "A1:B5"'),
    seriesBy: z
      .enum(['Auto', 'Columns', 'Rows'])
      .optional()
      .describe('Data series orientation (default: "Auto")'),
  },
  transform: (p) => ({
    type: p.type,
    sourceData: p.sourceData,
    seriesBy: p.seriesBy || 'Auto',
  }),
});

toolSchemaOverrides.set('format-excel-range', {
  description: 'Format cells in an Excel worksheet.',
  schema: {
    bold: z.boolean().optional().describe('Make text bold'),
    italic: z.boolean().optional().describe('Make text italic'),
    fontSize: z.number().optional().describe('Font size in points'),
    fontColor: z.string().optional().describe('Font color hex, e.g. "#FF0000"'),
    fillColor: z.string().optional().describe('Background color hex, e.g. "#FFFF00"'),
    numberFormat: z
      .string()
      .optional()
      .describe('Number format, e.g. "$#,##0.00", "0%"'),
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
    hasHeaders: z
      .boolean()
      .optional()
      .describe('Range has a header row (default: false)'),
  },
  transform: (p) => ({
    fields: [{ key: p.columnIndex, ascending: p.ascending !== false }],
    ...(p.hasHeaders !== undefined ? { hasHeaders: p.hasHeaders } : {}),
  }),
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── ONENOTE ──────────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

// ── OneNote (Read) ──

toolSchemaOverrides.set('list-onenote-notebooks', {
  description: 'List your OneNote notebooks. Returns notebook names and IDs.',
});

toolSchemaOverrides.set('list-onenote-notebook-sections', {
  description:
    'List sections in a OneNote notebook. Provide the notebook-id.',
});

toolSchemaOverrides.set('list-onenote-section-pages', {
  description:
    'List pages in a OneNote section. Provide the section-id.',
});

toolSchemaOverrides.set('get-onenote-page-content', {
  description:
    'Get the HTML content of a OneNote page. Provide the page-id.',
});

// ── OneNote (Write) ──

toolSchemaOverrides.set('create-onenote-page', {
  description: 'Create a OneNote page with a title and content.',
  schema: {
    title: z.string().describe('Page title'),
    content: z.string().describe('Page content (plain text or HTML)'),
  },
  transform: (p) => ({
    contentType: 'text/html',
    content: `<html><head><title>${p.title}</title></head><body>${p.content}</body></html>`,
  }),
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── SHAREPOINT ───────────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

toolSchemaOverrides.set('search-sharepoint-sites', {
  description: 'Search for SharePoint sites by name or keyword.',
  schema: {
    search: z.string().describe('Search text to find SharePoint sites'),
  },
  queryTransform: (p) => ({
    search: p.search as string,
  }),
});

toolSchemaOverrides.set('get-sharepoint-site', {
  description: 'Get details of a specific SharePoint site. Provide the site-id.',
});

toolSchemaOverrides.set('list-sharepoint-site-drives', {
  description:
    'List document libraries (drives) in a SharePoint site. Provide the site-id.',
});

toolSchemaOverrides.set('get-sharepoint-site-drive-by-id', {
  description:
    'Get a specific document library in a SharePoint site. Provide site-id and drive-id.',
});

toolSchemaOverrides.set('list-sharepoint-site-items', {
  description: 'List items in a SharePoint site. Provide the site-id.',
});

toolSchemaOverrides.set('get-sharepoint-site-item', {
  description:
    'Get a specific item from a SharePoint site. Provide site-id and item-id.',
});

toolSchemaOverrides.set('list-sharepoint-site-lists', {
  description: 'List all lists in a SharePoint site. Provide the site-id.',
});

toolSchemaOverrides.set('get-sharepoint-site-list', {
  description:
    'Get a specific list from a SharePoint site. Provide site-id and list-id.',
});

toolSchemaOverrides.set('list-sharepoint-site-list-items', {
  description:
    'List items in a SharePoint list. Provide site-id and list-id.',
  schema: genericListSchema,
  queryTransform: (p) => {
    const params = buildGenericListQuery(p);
    if (!p.count) params['$top'] = '25';
    return params;
  },
});

toolSchemaOverrides.set('get-sharepoint-site-list-item', {
  description:
    'Get a specific item from a SharePoint list. Provide site-id, list-id, and item-id.',
});

toolSchemaOverrides.set('get-sharepoint-site-by-path', {
  description:
    'Get a SharePoint site by its URL path, e.g. "/sites/marketing".',
});

toolSchemaOverrides.set('get-sharepoint-sites-delta', {
  description: 'Get recently changed SharePoint sites (delta query).',
});

// ═════════════════════════════════════════════════════════════════════════════════
// ── MICROSOFT SEARCH ─────────────────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════════

toolSchemaOverrides.set('search-query', {
  description:
    'Search across Microsoft 365: emails, calendar, files, SharePoint, Teams. Provide a query and what to search.',
  schema: {
    query: z
      .string()
      .describe('Search text, e.g. "budget report", "from:user@example.com"'),
    entityTypes: z
      .string()
      .describe(
        'Comma-separated types: "message" (email), "event" (calendar), "driveItem" (files), "chatMessage" (Teams), "site" (SharePoint), "person"'
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
