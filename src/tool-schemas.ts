import { z } from 'zod';

interface ToolSchemaOverride {
  description: string;
  bodySchema: z.ZodType<unknown>;
}

const emailAddressSchema = z.object({
  address: z.string().describe('Email address (e.g. "user@example.com")'),
  name: z.string().optional().describe('Display name of the recipient'),
});

const recipientSchema = z.object({
  emailAddress: emailAddressSchema,
});

const itemBodySchema = z.object({
  contentType: z.enum(['Text', 'HTML']).describe('Format of the body content: "Text" or "HTML"'),
  content: z.string().describe('The body content'),
});

const dateTimeTimeZoneSchema = z.object({
  dateTime: z
    .string()
    .describe('Date and time in ISO 8601 format, e.g. "2025-03-15T09:00:00"'),
  timeZone: z
    .string()
    .describe('IANA time zone name, e.g. "America/New_York", "UTC", "Europe/London"'),
});

const toolSchemaOverrides = new Map<string, ToolSchemaOverride>();

// ── Mail ────────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('send-mail', {
  description:
    'Send an email message. Requires a message object with toRecipients, subject, and body.',
  bodySchema: z
    .object({
      message: z.object({
        subject: z.string().describe('Email subject line'),
        body: itemBodySchema.describe('Email body with contentType and content'),
        toRecipients: z
          .array(recipientSchema)
          .describe('List of To recipients, each with emailAddress.address'),
        ccRecipients: z
          .array(recipientSchema)
          .optional()
          .describe('List of CC recipients'),
        bccRecipients: z
          .array(recipientSchema)
          .optional()
          .describe('List of BCC recipients'),
        replyTo: z
          .array(recipientSchema)
          .optional()
          .describe('Reply-to addresses'),
        importance: z
          .enum(['low', 'normal', 'high'])
          .optional()
          .describe('Message importance'),
      }),
      saveToSentItems: z
        .boolean()
        .optional()
        .describe('Whether to save the message in Sent Items (default: true)'),
    })
    .describe('The sendMail request body'),
});

toolSchemaOverrides.set('send-shared-mailbox-mail', {
  description:
    'Send an email from a shared mailbox. Requires user-id path param and a message object with toRecipients, subject, and body.',
  bodySchema: z
    .object({
      message: z.object({
        subject: z.string().describe('Email subject line'),
        body: itemBodySchema.describe('Email body with contentType and content'),
        toRecipients: z
          .array(recipientSchema)
          .describe('List of To recipients, each with emailAddress.address'),
        ccRecipients: z
          .array(recipientSchema)
          .optional()
          .describe('List of CC recipients'),
        bccRecipients: z
          .array(recipientSchema)
          .optional()
          .describe('List of BCC recipients'),
        importance: z
          .enum(['low', 'normal', 'high'])
          .optional()
          .describe('Message importance'),
      }),
      saveToSentItems: z
        .boolean()
        .optional()
        .describe('Whether to save the message in Sent Items (default: true)'),
    })
    .describe('The sendMail request body'),
});

toolSchemaOverrides.set('create-draft-email', {
  description:
    'Create a draft email message. Returns the draft message which can later be sent or updated.',
  bodySchema: z
    .object({
      subject: z.string().optional().describe('Email subject line'),
      body: itemBodySchema.optional().describe('Email body with contentType and content'),
      toRecipients: z
        .array(recipientSchema)
        .optional()
        .describe('List of To recipients, each with emailAddress.address'),
      ccRecipients: z
        .array(recipientSchema)
        .optional()
        .describe('List of CC recipients'),
      bccRecipients: z
        .array(recipientSchema)
        .optional()
        .describe('List of BCC recipients'),
      importance: z
        .enum(['low', 'normal', 'high'])
        .optional()
        .describe('Message importance'),
    })
    .describe('The draft message properties'),
});

toolSchemaOverrides.set('move-mail-message', {
  description:
    'Move a message to a different mail folder. Provide the destination folder ID (e.g. use list-mail-folders to find folder IDs like "inbox", "drafts", "deleteditems").',
  bodySchema: z
    .object({
      destinationId: z
        .string()
        .describe(
          'The ID of the destination mail folder (e.g. "inbox", "drafts", "deleteditems", or a folder ID from list-mail-folders)'
        ),
    })
    .describe('The move request body'),
});

toolSchemaOverrides.set('add-mail-attachment', {
  description:
    'Add a file attachment to a message. The @odata.type field must be set to "#microsoft.graph.fileAttachment". Content must be base64-encoded.',
  bodySchema: z
    .object({
      '@odata.type': z
        .string()
        .describe(
          'Must be "#microsoft.graph.fileAttachment" for file attachments'
        )
        .default('#microsoft.graph.fileAttachment'),
      name: z.string().describe('The file name of the attachment (e.g. "report.pdf")'),
      contentBytes: z
        .string()
        .describe('Base64-encoded content of the file attachment'),
      contentType: z
        .string()
        .optional()
        .describe('MIME type of the attachment (e.g. "application/pdf", "image/png")'),
      isInline: z
        .boolean()
        .optional()
        .describe('Whether the attachment is an inline attachment (default: false)'),
    })
    .describe('The file attachment to add'),
});

// ── Calendar ────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-calendar-event', {
  description:
    'Create a new calendar event. Requires subject, start (dateTime + timeZone), and end (dateTime + timeZone).',
  bodySchema: z
    .object({
      subject: z.string().describe('Event title/subject'),
      body: itemBodySchema.optional().describe('Event description/body'),
      start: dateTimeTimeZoneSchema.describe(
        'Event start date/time with timeZone, e.g. {"dateTime": "2025-03-15T09:00:00", "timeZone": "America/New_York"}'
      ),
      end: dateTimeTimeZoneSchema.describe(
        'Event end date/time with timeZone, e.g. {"dateTime": "2025-03-15T10:00:00", "timeZone": "America/New_York"}'
      ),
      location: z
        .object({
          displayName: z.string().describe('Display name of the location'),
        })
        .optional()
        .describe('Event location'),
      attendees: z
        .array(
          z.object({
            emailAddress: emailAddressSchema,
            type: z
              .enum(['required', 'optional', 'resource'])
              .optional()
              .describe('Attendee type'),
          })
        )
        .optional()
        .describe('List of event attendees'),
      isOnlineMeeting: z
        .boolean()
        .optional()
        .describe('Whether this is an online meeting (Teams)'),
      onlineMeetingProvider: z
        .enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer'])
        .optional()
        .describe('Online meeting provider when isOnlineMeeting is true'),
      isAllDay: z.boolean().optional().describe('Whether the event lasts all day'),
      recurrence: z
        .object({
          pattern: z.object({
            type: z
              .enum(['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'])
              .describe('Recurrence pattern type'),
            interval: z.number().describe('Interval between occurrences'),
            daysOfWeek: z
              .array(
                z.enum([
                  'sunday',
                  'monday',
                  'tuesday',
                  'wednesday',
                  'thursday',
                  'friday',
                  'saturday',
                ])
              )
              .optional()
              .describe('Days of the week for weekly recurrence'),
          }),
          range: z.object({
            type: z
              .enum(['endDate', 'noEnd', 'numbered'])
              .describe('Recurrence range type'),
            startDate: z.string().describe('Start date in YYYY-MM-DD format'),
            endDate: z.string().optional().describe('End date in YYYY-MM-DD format'),
            numberOfOccurrences: z.number().optional().describe('Number of occurrences'),
          }),
        })
        .optional()
        .describe('Recurrence pattern for recurring events'),
      reminderMinutesBeforeStart: z
        .number()
        .optional()
        .describe('Minutes before event to show reminder'),
    })
    .describe('The calendar event to create'),
});

toolSchemaOverrides.set('update-calendar-event', {
  description:
    'Update an existing calendar event. Only provide the fields you want to change.',
  bodySchema: z
    .object({
      subject: z.string().optional().describe('Event title/subject'),
      body: itemBodySchema.optional().describe('Event description/body'),
      start: dateTimeTimeZoneSchema
        .optional()
        .describe(
          'Event start date/time with timeZone, e.g. {"dateTime": "2025-03-15T09:00:00", "timeZone": "America/New_York"}'
        ),
      end: dateTimeTimeZoneSchema
        .optional()
        .describe(
          'Event end date/time with timeZone, e.g. {"dateTime": "2025-03-15T10:00:00", "timeZone": "America/New_York"}'
        ),
      location: z
        .object({
          displayName: z.string().describe('Display name of the location'),
        })
        .optional()
        .describe('Event location'),
      attendees: z
        .array(
          z.object({
            emailAddress: emailAddressSchema,
            type: z
              .enum(['required', 'optional', 'resource'])
              .optional()
              .describe('Attendee type'),
          })
        )
        .optional()
        .describe('List of event attendees'),
      isOnlineMeeting: z
        .boolean()
        .optional()
        .describe('Whether this is an online meeting (Teams)'),
      isAllDay: z.boolean().optional().describe('Whether the event lasts all day'),
      reminderMinutesBeforeStart: z
        .number()
        .optional()
        .describe('Minutes before event to show reminder'),
    })
    .describe('The calendar event fields to update'),
});

// ── To-Do Tasks ─────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-todo-task', {
  description:
    'Create a new To-Do task in a specific task list. Requires the task list ID (from list-todo-task-lists) and at minimum a title.',
  bodySchema: z
    .object({
      title: z.string().describe('Title of the task'),
      body: z
        .object({
          content: z.string().describe('Task notes/details'),
          contentType: z
            .enum(['text', 'html'])
            .optional()
            .describe('Content format: "text" or "html"'),
        })
        .optional()
        .describe('Task body/notes'),
      dueDateTime: dateTimeTimeZoneSchema
        .optional()
        .describe('Due date/time for the task'),
      reminderDateTime: dateTimeTimeZoneSchema
        .optional()
        .describe('Reminder date/time for the task'),
      importance: z
        .enum(['low', 'normal', 'high'])
        .optional()
        .describe('Task importance level'),
      status: z
        .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
        .optional()
        .describe('Task status'),
      categories: z
        .array(z.string())
        .optional()
        .describe('Categories/tags for the task'),
      isReminderOn: z
        .boolean()
        .optional()
        .describe('Whether a reminder is set for the task'),
    })
    .describe('The To-Do task to create'),
});

toolSchemaOverrides.set('update-todo-task', {
  description:
    'Update an existing To-Do task. Only provide the fields you want to change.',
  bodySchema: z
    .object({
      title: z.string().optional().describe('Title of the task'),
      body: z
        .object({
          content: z.string().describe('Task notes/details'),
          contentType: z
            .enum(['text', 'html'])
            .optional()
            .describe('Content format: "text" or "html"'),
        })
        .optional()
        .describe('Task body/notes'),
      dueDateTime: dateTimeTimeZoneSchema
        .optional()
        .describe('Due date/time for the task'),
      reminderDateTime: dateTimeTimeZoneSchema
        .optional()
        .describe('Reminder date/time for the task'),
      importance: z
        .enum(['low', 'normal', 'high'])
        .optional()
        .describe('Task importance level'),
      status: z
        .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
        .optional()
        .describe('Task status'),
      categories: z
        .array(z.string())
        .optional()
        .describe('Categories/tags for the task'),
      isReminderOn: z
        .boolean()
        .optional()
        .describe('Whether a reminder is set for the task'),
      completedDateTime: dateTimeTimeZoneSchema
        .optional()
        .describe('Date/time when the task was completed'),
    })
    .describe('The To-Do task fields to update'),
});

// ── Planner Tasks ───────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-planner-task', {
  description:
    'Create a new Planner task. Requires planId and title. Use get-planner-plan or list-planner-tasks to find the planId.',
  bodySchema: z
    .object({
      planId: z
        .string()
        .describe('ID of the plan to create the task in (from get-planner-plan)'),
      bucketId: z
        .string()
        .optional()
        .describe('ID of the bucket to place the task in'),
      title: z.string().describe('Title of the task'),
      assignments: z
        .record(
          z.string(),
          z.object({
            '@odata.type': z
              .string()
              .optional()
              .describe('Set to "#microsoft.graph.plannerAssignment"')
              .default('#microsoft.graph.plannerAssignment'),
            orderHint: z
              .string()
              .optional()
              .describe('Hint for ordering, e.g. " !"'),
          })
        )
        .optional()
        .describe(
          'Task assignments as an object keyed by user ID, e.g. {"userId": {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}}'
        ),
      dueDateTime: z
        .string()
        .optional()
        .describe('Due date in ISO 8601 format, e.g. "2025-03-15T00:00:00Z"'),
      startDateTime: z
        .string()
        .optional()
        .describe('Start date in ISO 8601 format'),
      percentComplete: z
        .number()
        .optional()
        .describe('Percentage complete (0-100)'),
      priority: z
        .number()
        .optional()
        .describe('Priority: 0=urgent, 1=important, 2=medium, 3-4=low, 5-10=unset'),
      orderHint: z.string().optional().describe('Hint for ordering the task'),
    })
    .describe('The Planner task to create'),
});

toolSchemaOverrides.set('update-planner-task', {
  description:
    'Update an existing Planner task. Only provide the fields you want to change. Note: requires If-Match header with the task etag for concurrency control.',
  bodySchema: z
    .object({
      title: z.string().optional().describe('Title of the task'),
      bucketId: z
        .string()
        .optional()
        .describe('ID of the bucket to move the task to'),
      assignments: z
        .record(
          z.string(),
          z
            .object({
              '@odata.type': z
                .string()
                .optional()
                .describe('Set to "#microsoft.graph.plannerAssignment"')
                .default('#microsoft.graph.plannerAssignment'),
              orderHint: z.string().optional(),
            })
            .nullable()
        )
        .optional()
        .describe(
          'Task assignments. Set a user ID key to null to unassign.'
        ),
      dueDateTime: z
        .string()
        .nullable()
        .optional()
        .describe('Due date in ISO 8601 format, or null to clear'),
      startDateTime: z
        .string()
        .nullable()
        .optional()
        .describe('Start date in ISO 8601 format, or null to clear'),
      percentComplete: z
        .number()
        .optional()
        .describe('Percentage complete (0-100)'),
      priority: z
        .number()
        .optional()
        .describe('Priority: 0=urgent, 1=important, 2=medium, 3-4=low, 5-10=unset'),
    })
    .describe('The Planner task fields to update'),
});

toolSchemaOverrides.set('update-planner-task-details', {
  description:
    'Update the details of a Planner task (description, checklist, references). Requires If-Match header with the task details etag.',
  bodySchema: z
    .object({
      description: z
        .string()
        .optional()
        .describe('Description/notes for the task'),
      previewType: z
        .enum(['automatic', 'noPreview', 'checklist', 'description', 'reference'])
        .optional()
        .describe('Type of preview to show on the task card'),
      checklist: z
        .record(
          z.string(),
          z.object({
            '@odata.type': z
              .string()
              .optional()
              .default('microsoft.graph.plannerChecklistItem'),
            title: z.string().describe('Checklist item text'),
            isChecked: z.boolean().optional().describe('Whether the item is checked'),
          })
        )
        .optional()
        .describe(
          'Checklist items as an object keyed by a GUID. Each value has a title and optional isChecked.'
        ),
      references: z
        .record(
          z.string(),
          z.object({
            '@odata.type': z
              .string()
              .optional()
              .default('microsoft.graph.plannerExternalReference'),
            alias: z.string().optional().describe('Display name for the reference'),
            type: z.string().optional().describe('Type of reference'),
            previewPriority: z.string().optional(),
          })
        )
        .optional()
        .describe(
          'External references keyed by URL-encoded URL'
        ),
    })
    .describe('The Planner task details to update'),
});

// ── Outlook Contacts ────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-outlook-contact', {
  description:
    'Create a new Outlook contact. Provide at minimum a givenName or displayName.',
  bodySchema: z
    .object({
      givenName: z.string().optional().describe('First name'),
      surname: z.string().optional().describe('Last name'),
      displayName: z.string().optional().describe('Full display name'),
      emailAddresses: z
        .array(
          z.object({
            address: z.string().describe('Email address'),
            name: z.string().optional().describe('Display name for this email'),
          })
        )
        .optional()
        .describe('Email addresses for the contact'),
      businessPhones: z
        .array(z.string())
        .optional()
        .describe('Business phone numbers'),
      mobilePhone: z.string().optional().describe('Mobile phone number'),
      homePhones: z
        .array(z.string())
        .optional()
        .describe('Home phone numbers'),
      jobTitle: z.string().optional().describe('Job title'),
      companyName: z.string().optional().describe('Company name'),
      department: z.string().optional().describe('Department'),
      officeLocation: z.string().optional().describe('Office location'),
      businessAddress: z
        .object({
          street: z.string().optional(),
          city: z.string().optional(),
          state: z.string().optional(),
          countryOrRegion: z.string().optional(),
          postalCode: z.string().optional(),
        })
        .optional()
        .describe('Business address'),
      homeAddress: z
        .object({
          street: z.string().optional(),
          city: z.string().optional(),
          state: z.string().optional(),
          countryOrRegion: z.string().optional(),
          postalCode: z.string().optional(),
        })
        .optional()
        .describe('Home address'),
      personalNotes: z.string().optional().describe('Personal notes about the contact'),
      birthday: z
        .string()
        .optional()
        .describe('Birthday in ISO 8601 format (YYYY-MM-DD)'),
    })
    .describe('The contact to create'),
});

toolSchemaOverrides.set('update-outlook-contact', {
  description:
    'Update an existing Outlook contact. Only provide the fields you want to change.',
  bodySchema: z
    .object({
      givenName: z.string().optional().describe('First name'),
      surname: z.string().optional().describe('Last name'),
      displayName: z.string().optional().describe('Full display name'),
      emailAddresses: z
        .array(
          z.object({
            address: z.string().describe('Email address'),
            name: z.string().optional().describe('Display name for this email'),
          })
        )
        .optional()
        .describe('Email addresses for the contact'),
      businessPhones: z
        .array(z.string())
        .optional()
        .describe('Business phone numbers'),
      mobilePhone: z.string().nullable().optional().describe('Mobile phone number'),
      jobTitle: z.string().optional().describe('Job title'),
      companyName: z.string().optional().describe('Company name'),
      department: z.string().optional().describe('Department'),
      personalNotes: z.string().optional().describe('Personal notes about the contact'),
    })
    .describe('The contact fields to update'),
});

// ── Teams / Chat Messages ───────────────────────────────────────────────────────

const chatMessageBodySchema = z
  .object({
    body: z.object({
      content: z.string().describe('The message content (can be plain text or HTML)'),
      contentType: z
        .enum(['text', 'html'])
        .optional()
        .describe('Content format: "text" (default) or "html"'),
    }),
  })
  .describe('The chat message to send');

toolSchemaOverrides.set('send-chat-message', {
  description:
    'Send a message in a Teams chat. Provide the chat ID and the message body content.',
  bodySchema: chatMessageBodySchema,
});

toolSchemaOverrides.set('send-channel-message', {
  description:
    'Send a message to a Teams channel. Provide the team ID, channel ID, and the message body content.',
  bodySchema: chatMessageBodySchema,
});

toolSchemaOverrides.set('reply-to-chat-message', {
  description:
    'Reply to a specific message in a Teams chat. Provide the chat ID, message ID, and the reply body content.',
  bodySchema: chatMessageBodySchema,
});

// ── OneNote ─────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-onenote-page', {
  description:
    'Create a new OneNote page. The body should be HTML content wrapped in the OneNote page structure. Provide the section ID via the onenoteSection-id path parameter.',
  bodySchema: z
    .object({
      contentType: z
        .string()
        .optional()
        .describe(
          'Content type, usually "text/html" or "application/xhtml+xml"'
        )
        .default('text/html'),
      content: z
        .string()
        .describe(
          'HTML content for the page. Must include <html><head><title>Page Title</title></head><body>Content here</body></html>'
        ),
    })
    .describe('The OneNote page to create'),
});

// ── Excel ───────────────────────────────────────────────────────────────────────

toolSchemaOverrides.set('create-excel-chart', {
  description:
    'Create a new chart in an Excel worksheet. Requires the drive ID, item ID, worksheet ID, chart type, and source data range.',
  bodySchema: z
    .object({
      type: z
        .string()
        .describe(
          'Chart type, e.g. "ColumnClustered", "Pie", "Line", "Bar", "Area", "XYScatter"'
        ),
      sourceData: z
        .string()
        .describe('Range address for the chart data source, e.g. "A1:B5" or "Sheet1!A1:C10"'),
      seriesBy: z
        .enum(['Auto', 'Columns', 'Rows'])
        .optional()
        .describe('Whether data series are in columns or rows (default: "Auto")'),
    })
    .describe('The chart creation parameters'),
});

toolSchemaOverrides.set('format-excel-range', {
  description:
    'Format a range of cells in an Excel worksheet. Provide formatting options like font, fill, borders, alignment, etc.',
  bodySchema: z
    .object({
      font: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.string().optional().describe('"None", "Single", or "Double"'),
          color: z.string().optional().describe('Font color, e.g. "#FF0000"'),
          name: z.string().optional().describe('Font name, e.g. "Calibri"'),
          size: z.number().optional().describe('Font size in points'),
        })
        .optional()
        .describe('Font formatting options'),
      fill: z
        .object({
          color: z
            .string()
            .optional()
            .describe('Background fill color, e.g. "#FFFF00"'),
        })
        .optional()
        .describe('Fill/background formatting'),
      horizontalAlignment: z
        .string()
        .optional()
        .describe('"General", "Left", "Center", "Right", "Fill", "Justify"'),
      verticalAlignment: z
        .string()
        .optional()
        .describe('"Top", "Center", "Bottom", "Justify"'),
      wrapText: z.boolean().optional().describe('Whether to wrap text in cells'),
      numberFormat: z
        .string()
        .optional()
        .describe('Number format string, e.g. "$#,##0.00", "0%"'),
    })
    .describe('The formatting to apply to the range'),
});

toolSchemaOverrides.set('sort-excel-range', {
  description:
    'Sort a range of cells in an Excel worksheet.',
  bodySchema: z
    .object({
      fields: z
        .array(
          z.object({
            key: z.number().describe('Column index to sort by (0-based)'),
            ascending: z.boolean().optional().describe('Sort ascending (default: true)'),
          })
        )
        .describe('Sort fields specifying which columns to sort by and direction'),
      matchCase: z.boolean().optional().describe('Whether sort is case-sensitive'),
      hasHeaders: z
        .boolean()
        .optional()
        .describe('Whether the range has a header row (default: false)'),
      method: z
        .string()
        .optional()
        .describe('Sort method: "PinYin" or "StrokeCount"'),
    })
    .describe('The sort configuration'),
});

// ── OneDrive Upload ─────────────────────────────────────────────────────────────

toolSchemaOverrides.set('upload-file-content', {
  description:
    'Upload or replace file content in OneDrive. The body should be the raw file content (string). For binary files, use base64 encoding. Requires the drive ID and item ID path parameters.',
  bodySchema: z
    .string()
    .describe('The file content to upload (raw text or base64-encoded binary content)'),
});

// ── Microsoft Search ────────────────────────────────────────────────────────────

toolSchemaOverrides.set('search-query', {
  description:
    'Search across Microsoft 365 content (emails, events, files, sites, Teams messages). Provide entity types to search and a query string.',
  bodySchema: z
    .object({
      requests: z
        .array(
          z.object({
            entityTypes: z
              .array(
                z.enum([
                  'message',
                  'event',
                  'driveItem',
                  'site',
                  'list',
                  'listItem',
                  'chatMessage',
                  'person',
                ])
              )
              .describe(
                'Types of content to search: "message" (email), "event" (calendar), "driveItem" (files), "site" (SharePoint), "chatMessage" (Teams), "person" (people)'
              ),
            query: z.object({
              queryString: z
                .string()
                .describe(
                  'Search query text, supports KQL syntax (e.g. "project report", "from:user@example.com")'
                ),
            }),
            from: z
              .number()
              .optional()
              .describe('Starting index for pagination (default: 0)'),
            size: z
              .number()
              .optional()
              .describe('Number of results to return (default: 25, max: 1000)'),
            fields: z
              .array(z.string())
              .optional()
              .describe(
                'Specific fields to return in results (e.g. ["subject", "from", "receivedDateTime"])'
              ),
          })
        )
        .describe('Array of search request objects (typically one)'),
    })
    .describe('The search query request body'),
});

export { toolSchemaOverrides };
export type { ToolSchemaOverride };
