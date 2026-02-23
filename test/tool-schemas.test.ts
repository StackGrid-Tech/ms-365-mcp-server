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

const writeEndpoints = endpointsData.filter((e) =>
  ['post', 'put', 'patch'].includes(e.method.toLowerCase())
);

describe('Tool Schema Overrides', () => {
  describe('coverage', () => {
    it('should have overrides for all POST/PUT/PATCH endpoints', () => {
      const missingOverrides = writeEndpoints
        .filter((e) => !toolSchemaOverrides.has(e.toolName))
        .map((e) => `${e.method.toUpperCase()} ${e.pathPattern} (${e.toolName})`);

      expect(
        missingOverrides,
        `Missing schema overrides for:\n  ${missingOverrides.join('\n  ')}`
      ).toHaveLength(0);
    });

    it('should not have overrides for DELETE or GET endpoints', () => {
      const readOnlyEndpoints = endpointsData.filter(
        (e) => e.method === 'delete' || e.method === 'get'
      );
      for (const endpoint of readOnlyEndpoints) {
        expect(toolSchemaOverrides.has(endpoint.toolName)).toBe(false);
      }
    });
  });

  describe('override structure', () => {
    it('every override has description, schema, and transform', () => {
      for (const [toolName, override] of toolSchemaOverrides) {
        expect(override.description, `${toolName} missing description`).toBeTruthy();
        expect(override.description.length, `${toolName} description too short`).toBeGreaterThan(10);
        expect(override.schema, `${toolName} missing schema`).toBeDefined();
        expect(typeof override.transform, `${toolName} transform not a function`).toBe('function');
      }
    });

    it('every override schema should have at most 8 params', () => {
      for (const [toolName, override] of toolSchemaOverrides) {
        const paramCount = Object.keys(override.schema).length;
        expect(paramCount, `${toolName} has too many params (${paramCount})`).toBeLessThanOrEqual(8);
      }
    });

    it('every schema param should be a Zod type', () => {
      for (const [toolName, override] of toolSchemaOverrides) {
        for (const [key, schema] of Object.entries(override.schema)) {
          expect(schema instanceof z.ZodType, `${toolName}.${key} is not a ZodType`).toBe(true);
        }
      }
    });
  });

  describe('send-mail', () => {
    const override = toolSchemaOverrides.get('send-mail')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept simple flat params', () => {
      const result = schema.safeParse({
        to: 'user@example.com',
        subject: 'Hello',
        content: 'Hi there',
      });
      expect(result.success).toBe(true);
    });

    it('should accept optional cc/bcc/isHtml', () => {
      const result = schema.safeParse({
        to: 'a@b.com',
        subject: 'Test',
        content: '<b>Bold</b>',
        cc: 'cc@b.com, cc2@b.com',
        bcc: 'bcc@b.com',
        isHtml: true,
      });
      expect(result.success).toBe(true);
    });

    it('should reject without required fields', () => {
      expect(schema.safeParse({ subject: 'Test' }).success).toBe(false);
      expect(schema.safeParse({ to: 'a@b.com' }).success).toBe(false);
    });

    it('transform should produce correct API body', () => {
      const body = override.transform({
        to: 'alice@example.com, bob@example.com',
        subject: 'Hello',
        content: 'Hi',
      });
      expect(body).toEqual({
        message: {
          subject: 'Hello',
          body: { contentType: 'Text', content: 'Hi' },
          toRecipients: [
            { emailAddress: { address: 'alice@example.com' } },
            { emailAddress: { address: 'bob@example.com' } },
          ],
        },
        saveToSentItems: true,
      });
    });

    it('transform should handle HTML and CC', () => {
      const body = override.transform({
        to: 'a@b.com',
        subject: 'Test',
        content: '<b>Bold</b>',
        cc: 'cc@b.com',
        isHtml: true,
      }) as Record<string, unknown>;
      const msg = body.message as Record<string, unknown>;
      const msgBody = msg.body as Record<string, unknown>;
      expect(msgBody.contentType).toBe('HTML');
      expect(msg.ccRecipients).toEqual([{ emailAddress: { address: 'cc@b.com' } }]);
    });
  });

  describe('create-calendar-event', () => {
    const override = toolSchemaOverrides.get('create-calendar-event')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept minimal params', () => {
      const result = schema.safeParse({
        subject: 'Meeting',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T10:00:00',
      });
      expect(result.success).toBe(true);
    });

    it('should reject without start/end', () => {
      expect(schema.safeParse({ subject: 'Meeting' }).success).toBe(false);
    });

    it('transform should produce correct API body', () => {
      const body = override.transform({
        subject: 'Team Meeting',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T10:00:00',
        timeZone: 'America/New_York',
        location: 'Room 101',
        attendees: 'alice@example.com, bob@example.com',
      });
      expect(body).toEqual({
        subject: 'Team Meeting',
        start: { dateTime: '2025-03-15T09:00:00', timeZone: 'America/New_York' },
        end: { dateTime: '2025-03-15T10:00:00', timeZone: 'America/New_York' },
        location: { displayName: 'Room 101' },
        attendees: [
          { emailAddress: { address: 'alice@example.com' }, type: 'required' },
          { emailAddress: { address: 'bob@example.com' }, type: 'required' },
        ],
      });
    });

    it('transform should default timeZone to UTC', () => {
      const body = override.transform({
        subject: 'Test',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T10:00:00',
      }) as Record<string, unknown>;
      const start = body.start as Record<string, unknown>;
      expect(start.timeZone).toBe('UTC');
    });

    it('transform should handle online meeting flag', () => {
      const body = override.transform({
        subject: 'Virtual Standup',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T09:15:00',
        isOnlineMeeting: true,
      }) as Record<string, unknown>;
      expect(body.isOnlineMeeting).toBe(true);
      expect(body.onlineMeetingProvider).toBe('teamsForBusiness');
    });
  });

  describe('create-todo-task', () => {
    const override = toolSchemaOverrides.get('create-todo-task')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept just a title', () => {
      expect(schema.safeParse({ title: 'Buy groceries' }).success).toBe(true);
    });

    it('should reject without title', () => {
      expect(schema.safeParse({ dueDate: '2025-03-15' }).success).toBe(false);
    });

    it('transform should produce correct API body with dueDate', () => {
      const body = override.transform({
        title: 'Buy groceries',
        dueDate: '2025-03-15',
        notes: 'Milk, bread, eggs',
        importance: 'high',
      });
      expect(body).toEqual({
        title: 'Buy groceries',
        dueDateTime: { dateTime: '2025-03-15T00:00:00', timeZone: 'UTC' },
        body: { content: 'Milk, bread, eggs', contentType: 'text' },
        importance: 'high',
      });
    });

    it('transform should produce minimal body', () => {
      const body = override.transform({ title: 'Simple task' });
      expect(body).toEqual({ title: 'Simple task' });
    });
  });

  describe('create-planner-task', () => {
    const override = toolSchemaOverrides.get('create-planner-task')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept planId and title', () => {
      expect(schema.safeParse({ planId: 'plan-123', title: 'Review' }).success).toBe(true);
    });

    it('should reject without planId', () => {
      expect(schema.safeParse({ title: 'Test' }).success).toBe(false);
    });

    it('transform should handle assignedTo', () => {
      const body = override.transform({
        planId: 'plan-123',
        title: 'Review',
        assignedTo: 'user-1, user-2',
      }) as Record<string, unknown>;
      expect(body.assignments).toEqual({
        'user-1': { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
        'user-2': { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      });
    });
  });

  describe('add-mail-attachment', () => {
    const override = toolSchemaOverrides.get('add-mail-attachment')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept name and contentBytes', () => {
      expect(schema.safeParse({ name: 'file.pdf', contentBytes: 'base64data==' }).success).toBe(
        true
      );
    });

    it('should reject without required fields', () => {
      expect(schema.safeParse({ name: 'file.pdf' }).success).toBe(false);
      expect(schema.safeParse({ contentBytes: 'data==' }).success).toBe(false);
    });

    it('transform should auto-inject @odata.type', () => {
      const body = override.transform({
        name: 'report.pdf',
        contentBytes: 'dGVzdA==',
        contentType: 'application/pdf',
      }) as Record<string, unknown>;
      expect(body['@odata.type']).toBe('#microsoft.graph.fileAttachment');
      expect(body.name).toBe('report.pdf');
      expect(body.contentBytes).toBe('dGVzdA==');
      expect(body.contentType).toBe('application/pdf');
    });
  });

  describe('move-mail-message', () => {
    const override = toolSchemaOverrides.get('move-mail-message')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept destinationId', () => {
      expect(schema.safeParse({ destinationId: 'inbox' }).success).toBe(true);
    });

    it('should reject without destinationId', () => {
      expect(schema.safeParse({}).success).toBe(false);
    });

    it('transform should pass through', () => {
      expect(override.transform({ destinationId: 'drafts' })).toEqual({
        destinationId: 'drafts',
      });
    });
  });

  describe('search-query', () => {
    const override = toolSchemaOverrides.get('search-query')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept query and entityTypes strings', () => {
      expect(
        schema.safeParse({ query: 'project report', entityTypes: 'message,driveItem' }).success
      ).toBe(true);
    });

    it('should reject without query', () => {
      expect(schema.safeParse({ entityTypes: 'message' }).success).toBe(false);
    });

    it('transform should produce correct nested API body', () => {
      const body = override.transform({
        query: 'budget 2025',
        entityTypes: 'message, driveItem',
        size: 10,
      });
      expect(body).toEqual({
        requests: [
          {
            entityTypes: ['message', 'driveItem'],
            query: { queryString: 'budget 2025' },
            size: 10,
          },
        ],
      });
    });
  });

  describe('send-chat-message', () => {
    const override = toolSchemaOverrides.get('send-chat-message')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept just content', () => {
      expect(schema.safeParse({ content: 'Hello team!' }).success).toBe(true);
    });

    it('should reject without content', () => {
      expect(schema.safeParse({}).success).toBe(false);
    });

    it('transform should wrap in body object', () => {
      expect(override.transform({ content: 'Hello!' })).toEqual({
        body: { content: 'Hello!' },
      });
    });
  });

  describe('create-outlook-contact', () => {
    const override = toolSchemaOverrides.get('create-outlook-contact')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept givenName', () => {
      expect(schema.safeParse({ givenName: 'John' }).success).toBe(true);
    });

    it('transform should expand email and phone', () => {
      const body = override.transform({
        givenName: 'Jane',
        surname: 'Doe',
        email: 'jane@example.com',
        phone: '+1-555-0100',
        company: 'Acme',
      }) as Record<string, unknown>;
      expect(body.givenName).toBe('Jane');
      expect(body.surname).toBe('Doe');
      expect(body.emailAddresses).toEqual([{ address: 'jane@example.com', name: '' }]);
      expect(body.businessPhones).toEqual(['+1-555-0100']);
      expect(body.companyName).toBe('Acme');
    });
  });

  describe('create-excel-chart', () => {
    const override = toolSchemaOverrides.get('create-excel-chart')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept type and sourceData', () => {
      expect(schema.safeParse({ type: 'Pie', sourceData: 'A1:B5' }).success).toBe(true);
    });

    it('should reject without type', () => {
      expect(schema.safeParse({ sourceData: 'A1:B5' }).success).toBe(false);
    });

    it('transform should pass through with default seriesBy', () => {
      expect(override.transform({ type: 'Line', sourceData: 'A1:C10' })).toEqual({
        type: 'Line',
        sourceData: 'A1:C10',
        seriesBy: 'Auto',
      });
    });
  });

  describe('create-onenote-page', () => {
    const override = toolSchemaOverrides.get('create-onenote-page')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept title and content', () => {
      expect(schema.safeParse({ title: 'My Page', content: 'Hello' }).success).toBe(true);
    });

    it('transform should wrap in HTML structure', () => {
      const body = override.transform({ title: 'Notes', content: 'Some text' }) as Record<
        string,
        unknown
      >;
      expect(body.contentType).toBe('text/html');
      expect(body.content).toContain('<title>Notes</title>');
      expect(body.content).toContain('Some text');
    });
  });

  describe('upload-file-content', () => {
    const override = toolSchemaOverrides.get('upload-file-content')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept content string', () => {
      expect(schema.safeParse({ content: 'file data' }).success).toBe(true);
    });

    it('transform should return raw content string', () => {
      expect(override.transform({ content: 'raw file data' })).toBe('raw file data');
    });
  });

  describe('format-excel-range', () => {
    const override = toolSchemaOverrides.get('format-excel-range')!;

    it('transform should build font/fill from flat params', () => {
      const body = override.transform({
        bold: true,
        fontSize: 14,
        fontColor: '#FF0000',
        fillColor: '#FFFF00',
      });
      expect(body).toEqual({
        font: { bold: true, size: 14, color: '#FF0000' },
        fill: { color: '#FFFF00' },
      });
    });
  });

  describe('sort-excel-range', () => {
    const override = toolSchemaOverrides.get('sort-excel-range')!;

    it('transform should wrap columnIndex into fields array', () => {
      const body = override.transform({ columnIndex: 2, ascending: false, hasHeaders: true });
      expect(body).toEqual({
        fields: [{ key: 2, ascending: false }],
        hasHeaders: true,
      });
    });
  });

  describe('create-draft-email', () => {
    const override = toolSchemaOverrides.get('create-draft-email')!;
    const schema = z.object(override.schema as z.ZodRawShape);

    it('should accept subject and content', () => {
      expect(schema.safeParse({ subject: 'Draft', content: 'Text' }).success).toBe(true);
    });

    it('transform should produce correct body', () => {
      const body = override.transform({
        subject: 'Draft',
        content: 'Hello',
        to: 'a@b.com',
      });
      expect(body).toEqual({
        subject: 'Draft',
        body: { contentType: 'Text', content: 'Hello' },
        toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
      });
    });
  });
});
