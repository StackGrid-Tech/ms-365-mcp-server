import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { toolSchemaOverrides } from '../src/tool-schemas.js';
import { z } from 'zod';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'src', 'endpoints.json'), 'utf8')
) as { toolName: string; method: string; pathPattern: string }[];

const writeEndpoints = endpointsData.filter(
  (e) => ['post', 'put', 'patch'].includes(e.method.toLowerCase())
);

describe('Tool Schema Overrides', () => {
  describe('coverage', () => {
    it('should have overrides for all POST/PUT/PATCH endpoints that need request bodies', () => {
      const missingOverrides = writeEndpoints
        .filter((e) => !toolSchemaOverrides.has(e.toolName))
        .map((e) => `${e.method.toUpperCase()} ${e.pathPattern} (${e.toolName})`);

      expect(
        missingOverrides,
        `Missing schema overrides for:\n  ${missingOverrides.join('\n  ')}`
      ).toHaveLength(0);
    });

    it('should not have overrides for DELETE endpoints', () => {
      const deleteEndpoints = endpointsData.filter((e) => e.method === 'delete');
      for (const endpoint of deleteEndpoints) {
        expect(toolSchemaOverrides.has(endpoint.toolName)).toBe(false);
      }
    });

    it('should not have overrides for GET endpoints', () => {
      const getEndpoints = endpointsData.filter((e) => e.method === 'get');
      for (const endpoint of getEndpoints) {
        expect(toolSchemaOverrides.has(endpoint.toolName)).toBe(false);
      }
    });
  });

  describe('override structure', () => {
    it('every override should have a non-empty description', () => {
      for (const [toolName, override] of toolSchemaOverrides) {
        expect(override.description, `${toolName} has empty description`).toBeTruthy();
        expect(override.description.length, `${toolName} description too short`).toBeGreaterThan(10);
      }
    });

    it('every override should have a bodySchema', () => {
      for (const [toolName, override] of toolSchemaOverrides) {
        expect(override.bodySchema, `${toolName} missing bodySchema`).toBeDefined();
      }
    });
  });

  describe('send-mail', () => {
    const override = toolSchemaOverrides.get('send-mail')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid send-mail payload', () => {
      const result = schema.safeParse({
        message: {
          subject: 'Test Subject',
          body: { contentType: 'Text', content: 'Hello world' },
          toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
        },
        saveToSentItems: true,
      });
      expect(result.success).toBe(true);
    });

    it('should accept send-mail with optional fields', () => {
      const result = schema.safeParse({
        message: {
          subject: 'Test',
          body: { contentType: 'HTML', content: '<b>Hello</b>' },
          toRecipients: [
            { emailAddress: { address: 'a@b.com', name: 'User A' } },
          ],
          ccRecipients: [{ emailAddress: { address: 'cc@b.com' } }],
          bccRecipients: [{ emailAddress: { address: 'bcc@b.com' } }],
          importance: 'high',
        },
      });
      expect(result.success).toBe(true);
    });

    it('should reject send-mail without message', () => {
      const result = schema.safeParse({
        saveToSentItems: true,
      });
      expect(result.success).toBe(false);
    });

    it('should reject send-mail without toRecipients', () => {
      const result = schema.safeParse({
        message: {
          subject: 'Test',
          body: { contentType: 'Text', content: 'Hello' },
        },
      });
      expect(result.success).toBe(false);
    });
  });

  describe('create-calendar-event', () => {
    const override = toolSchemaOverrides.get('create-calendar-event')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid calendar event', () => {
      const result = schema.safeParse({
        subject: 'Team Meeting',
        start: { dateTime: '2025-03-15T09:00:00', timeZone: 'America/New_York' },
        end: { dateTime: '2025-03-15T10:00:00', timeZone: 'America/New_York' },
      });
      expect(result.success).toBe(true);
    });

    it('should accept a calendar event with all optional fields', () => {
      const result = schema.safeParse({
        subject: 'Team Meeting',
        body: { contentType: 'HTML', content: '<p>Agenda</p>' },
        start: { dateTime: '2025-03-15T09:00:00', timeZone: 'UTC' },
        end: { dateTime: '2025-03-15T10:00:00', timeZone: 'UTC' },
        location: { displayName: 'Conference Room A' },
        attendees: [
          {
            emailAddress: { address: 'colleague@example.com', name: 'Colleague' },
            type: 'required',
          },
        ],
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness',
        isAllDay: false,
        reminderMinutesBeforeStart: 15,
      });
      expect(result.success).toBe(true);
    });

    it('should reject calendar event without required start/end', () => {
      const result = schema.safeParse({
        subject: 'Test Event',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('create-todo-task', () => {
    const override = toolSchemaOverrides.get('create-todo-task')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a minimal todo task', () => {
      const result = schema.safeParse({
        title: 'Buy groceries',
      });
      expect(result.success).toBe(true);
    });

    it('should accept a full todo task', () => {
      const result = schema.safeParse({
        title: 'Buy groceries',
        body: { content: 'Milk, eggs, bread', contentType: 'text' },
        dueDateTime: { dateTime: '2025-03-15T00:00:00', timeZone: 'UTC' },
        importance: 'high',
        status: 'inProgress',
        isReminderOn: true,
        reminderDateTime: { dateTime: '2025-03-14T09:00:00', timeZone: 'UTC' },
        categories: ['Shopping'],
      });
      expect(result.success).toBe(true);
    });

    it('should reject todo task without title', () => {
      const result = schema.safeParse({
        importance: 'high',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('create-planner-task', () => {
    const override = toolSchemaOverrides.get('create-planner-task')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a minimal planner task', () => {
      const result = schema.safeParse({
        planId: 'plan-123',
        title: 'Design review',
      });
      expect(result.success).toBe(true);
    });

    it('should accept a planner task with assignments', () => {
      const result = schema.safeParse({
        planId: 'plan-123',
        bucketId: 'bucket-456',
        title: 'Design review',
        assignments: {
          'user-id-789': {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !',
          },
        },
        dueDateTime: '2025-03-15T00:00:00Z',
        percentComplete: 50,
        priority: 1,
      });
      expect(result.success).toBe(true);
    });

    it('should reject planner task without planId', () => {
      const result = schema.safeParse({
        title: 'Test',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('add-mail-attachment', () => {
    const override = toolSchemaOverrides.get('add-mail-attachment')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid file attachment', () => {
      const result = schema.safeParse({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'report.pdf',
        contentBytes: 'base64encodedcontent==',
        contentType: 'application/pdf',
      });
      expect(result.success).toBe(true);
    });

    it('should preserve @odata.type in the schema', () => {
      const result = schema.safeParse({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'test.txt',
        contentBytes: 'dGVzdA==',
      });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data['@odata.type']).toBe('#microsoft.graph.fileAttachment');
      }
    });

    it('should reject attachment without name', () => {
      const result = schema.safeParse({
        '@odata.type': '#microsoft.graph.fileAttachment',
        contentBytes: 'dGVzdA==',
      });
      expect(result.success).toBe(false);
    });

    it('should reject attachment without contentBytes', () => {
      const result = schema.safeParse({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'test.txt',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('move-mail-message', () => {
    const override = toolSchemaOverrides.get('move-mail-message')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid move request', () => {
      const result = schema.safeParse({
        destinationId: 'inbox',
      });
      expect(result.success).toBe(true);
    });

    it('should reject without destinationId', () => {
      const result = schema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe('search-query', () => {
    const override = toolSchemaOverrides.get('search-query')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid search query', () => {
      const result = schema.safeParse({
        requests: [
          {
            entityTypes: ['message'],
            query: { queryString: 'project report' },
          },
        ],
      });
      expect(result.success).toBe(true);
    });

    it('should accept search with pagination and fields', () => {
      const result = schema.safeParse({
        requests: [
          {
            entityTypes: ['driveItem', 'site'],
            query: { queryString: 'budget 2025' },
            from: 0,
            size: 10,
            fields: ['name', 'lastModifiedDateTime'],
          },
        ],
      });
      expect(result.success).toBe(true);
    });

    it('should reject search without requests', () => {
      const result = schema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe('send-chat-message', () => {
    const override = toolSchemaOverrides.get('send-chat-message')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid chat message', () => {
      const result = schema.safeParse({
        body: { content: 'Hello team!' },
      });
      expect(result.success).toBe(true);
    });

    it('should accept html content', () => {
      const result = schema.safeParse({
        body: { content: '<b>Important</b>', contentType: 'html' },
      });
      expect(result.success).toBe(true);
    });

    it('should reject without body', () => {
      const result = schema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe('create-outlook-contact', () => {
    const override = toolSchemaOverrides.get('create-outlook-contact')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a minimal contact', () => {
      const result = schema.safeParse({
        givenName: 'John',
        surname: 'Doe',
      });
      expect(result.success).toBe(true);
    });

    it('should accept a contact with email and phone', () => {
      const result = schema.safeParse({
        givenName: 'Jane',
        surname: 'Smith',
        displayName: 'Jane Smith',
        emailAddresses: [{ address: 'jane@example.com', name: 'Jane Smith' }],
        businessPhones: ['+1-555-0100'],
        companyName: 'Acme Corp',
        jobTitle: 'Engineer',
      });
      expect(result.success).toBe(true);
    });
  });

  describe('create-excel-chart', () => {
    const override = toolSchemaOverrides.get('create-excel-chart')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a valid chart creation', () => {
      const result = schema.safeParse({
        type: 'ColumnClustered',
        sourceData: 'A1:B5',
        seriesBy: 'Auto',
      });
      expect(result.success).toBe(true);
    });

    it('should reject without type', () => {
      const result = schema.safeParse({
        sourceData: 'A1:B5',
      });
      expect(result.success).toBe(false);
    });

    it('should reject without sourceData', () => {
      const result = schema.safeParse({
        type: 'Pie',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('create-draft-email', () => {
    const override = toolSchemaOverrides.get('create-draft-email')!;
    const schema = override.bodySchema as z.ZodObject<z.ZodRawShape>;

    it('should accept a draft with all fields optional', () => {
      const result = schema.safeParse({});
      expect(result.success).toBe(true);
    });

    it('should accept a full draft', () => {
      const result = schema.safeParse({
        subject: 'Draft email',
        body: { contentType: 'Text', content: 'Draft content' },
        toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
        importance: 'high',
      });
      expect(result.success).toBe(true);
    });
  });

  describe('upload-file-content', () => {
    const override = toolSchemaOverrides.get('upload-file-content')!;

    it('should accept a string body', () => {
      const result = (override.bodySchema as z.ZodString).safeParse('file content here');
      expect(result.success).toBe(true);
    });

    it('should reject non-string body', () => {
      const result = (override.bodySchema as z.ZodString).safeParse(123);
      expect(result.success).toBe(false);
    });
  });
});
