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

describe('Tool Schema Overrides', () => {
  // ── Coverage ──────────────────────────────────────────────────────────────

  describe('full coverage', () => {
    it('every endpoint in endpoints.json has an override', () => {
      const missing = endpointsData
        .filter((e) => !toolSchemaOverrides.has(e.toolName))
        .map((e) => `${e.method.toUpperCase()} ${e.pathPattern} (${e.toolName})`);
      expect(missing, `Missing overrides:\n  ${missing.join('\n  ')}`).toHaveLength(0);
    });

    it('no orphan overrides (every override maps to an endpoint)', () => {
      const endpointNames = new Set(endpointsData.map((e) => e.toolName));
      const orphans: string[] = [];
      for (const toolName of toolSchemaOverrides.keys()) {
        if (!endpointNames.has(toolName)) orphans.push(toolName);
      }
      expect(orphans, `Orphan overrides: ${orphans.join(', ')}`).toHaveLength(0);
    });
  });

  // ── Structure ─────────────────────────────────────────────────────────────

  describe('structure', () => {
    it('every override has a meaningful description', () => {
      for (const [name, o] of toolSchemaOverrides) {
        expect(o.description, `${name}`).toBeTruthy();
        expect(o.description.length, `${name} too short`).toBeGreaterThan(10);
      }
    });

    it('write endpoints have schema + transform', () => {
      for (const ep of endpointsData.filter((e) =>
        ['post', 'put', 'patch'].includes(e.method)
      )) {
        const o = toolSchemaOverrides.get(ep.toolName);
        if (!o) continue;
        expect(o.schema, `${ep.toolName} missing schema`).toBeDefined();
        expect(typeof o.transform, `${ep.toolName} transform not fn`).toBe('function');
      }
    });

    it('overrides with schema have at most 8 params', () => {
      for (const [name, o] of toolSchemaOverrides) {
        if (!o.schema) continue;
        const n = Object.keys(o.schema).length;
        expect(n, `${name} has ${n} params`).toBeLessThanOrEqual(8);
      }
    });

    it('every schema param is a Zod type', () => {
      for (const [name, o] of toolSchemaOverrides) {
        if (!o.schema) continue;
        for (const [key, s] of Object.entries(o.schema)) {
          expect(s instanceof z.ZodType, `${name}.${key}`).toBe(true);
        }
      }
    });
  });

  // ── Mail: consolidated list-mail-messages ─────────────────────────────────

  describe('list-mail-messages (consolidated)', () => {
    const o = toolSchemaOverrides.get('list-mail-messages')!;

    it('queryTransform builds OData', () => {
      const q = o.queryTransform!({ from: 'alice@x.com', unreadOnly: true, count: 5 });
      expect(q['$filter']).toContain("from/emailAddress/address eq 'alice@x.com'");
      expect(q['$filter']).toContain('isRead eq false');
      expect(q['$top']).toBe('5');
    });

    it('pathTransform: default → /me/messages', () => {
      expect(o.pathTransform!('/me/messages', {})).toBe('/me/messages');
    });

    it('pathTransform: folderId → /me/mailFolders/{id}/messages', () => {
      expect(o.pathTransform!('/me/messages', { folderId: 'inbox-id' })).toBe(
        '/me/mailFolders/inbox-id/messages'
      );
    });

    it('pathTransform: userId → /users/{id}/messages', () => {
      expect(o.pathTransform!('/me/messages', { userId: 'shared@x.com' })).toBe(
        '/users/shared@x.com/messages'
      );
    });

    it('pathTransform: userId + folderId', () => {
      expect(
        o.pathTransform!('/me/messages', { userId: 'u1', folderId: 'f1' })
      ).toBe('/users/u1/mailFolders/f1/messages');
    });
  });

  // ── Mail: consolidated get-mail-message ───────────────────────────────────

  describe('get-mail-message (consolidated)', () => {
    const o = toolSchemaOverrides.get('get-mail-message')!;

    it('pathTransform: default (no userId) → unchanged', () => {
      expect(o.pathTransform!('/me/messages/{message-id}', {})).toBe(
        '/me/messages/{message-id}'
      );
    });

    it('pathTransform: userId → /users/{id}/messages/{message-id}', () => {
      expect(
        o.pathTransform!('/me/messages/{message-id}', { userId: 'shared@x.com' })
      ).toBe('/users/shared@x.com/messages/{message-id}');
    });
  });

  // ── Mail: consolidated send-mail ──────────────────────────────────────────

  describe('send-mail (consolidated)', () => {
    const o = toolSchemaOverrides.get('send-mail')!;

    it('transform produces correct body', () => {
      expect(
        o.transform!({ to: 'a@b.com, c@d.com', subject: 'Hi', content: 'Hello' })
      ).toEqual({
        message: {
          subject: 'Hi',
          body: { contentType: 'Text', content: 'Hello' },
          toRecipients: [
            { emailAddress: { address: 'a@b.com' } },
            { emailAddress: { address: 'c@d.com' } },
          ],
        },
        saveToSentItems: true,
      });
    });

    it('transform handles HTML + CC + BCC', () => {
      const body = o.transform!({
        to: 'a@b.com',
        subject: 'T',
        content: '<b>X</b>',
        cc: 'cc@b.com',
        bcc: 'bcc@b.com',
        isHtml: true,
      }) as Record<string, unknown>;
      const msg = body.message as Record<string, unknown>;
      expect((msg.body as Record<string, unknown>).contentType).toBe('HTML');
      expect(msg.ccRecipients).toEqual([{ emailAddress: { address: 'cc@b.com' } }]);
      expect(msg.bccRecipients).toEqual([{ emailAddress: { address: 'bcc@b.com' } }]);
    });

    it('pathTransform: default → /me/sendMail', () => {
      expect(o.pathTransform!('/me/sendMail', {})).toBe('/me/sendMail');
    });

    it('pathTransform: userId → /users/{id}/sendMail', () => {
      expect(o.pathTransform!('/me/sendMail', { userId: 'shared@x.com' })).toBe(
        '/users/shared@x.com/sendMail'
      );
    });
  });

  // ── Other mail transforms ─────────────────────────────────────────────────

  describe('create-draft-email transform', () => {
    it('produces correct body', () => {
      const o = toolSchemaOverrides.get('create-draft-email')!;
      expect(o.transform!({ subject: 'D', content: 'Hi', to: 'a@b.com' })).toEqual({
        subject: 'D',
        body: { contentType: 'Text', content: 'Hi' },
        toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
      });
    });
  });

  describe('move-mail-message transform', () => {
    it('passes through', () => {
      const o = toolSchemaOverrides.get('move-mail-message')!;
      expect(o.transform!({ destinationId: 'inbox' })).toEqual({ destinationId: 'inbox' });
    });
  });

  describe('add-mail-attachment transform', () => {
    it('auto-injects @odata.type', () => {
      const o = toolSchemaOverrides.get('add-mail-attachment')!;
      const body = o.transform!({
        name: 'f.pdf',
        contentBytes: 'dGVzdA==',
        contentType: 'application/pdf',
      }) as Record<string, unknown>;
      expect(body['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    });
  });

  // ── Calendar ──────────────────────────────────────────────────────────────

  describe('create-calendar-event transform', () => {
    const o = toolSchemaOverrides.get('create-calendar-event')!;

    it('builds nested start/end', () => {
      const body = o.transform!({
        subject: 'Standup',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T09:15:00',
      }) as Record<string, unknown>;
      expect(body.start).toEqual({ dateTime: '2025-03-15T09:00:00', timeZone: 'UTC' });
    });

    it('handles attendees + online meeting', () => {
      const body = o.transform!({
        subject: 'Call',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T10:00:00',
        attendees: 'a@b.com, c@d.com',
        isOnlineMeeting: true,
      }) as Record<string, unknown>;
      expect(body.isOnlineMeeting).toBe(true);
      expect((body.attendees as unknown[]).length).toBe(2);
    });
  });

  describe('list-calendar-events queryTransform', () => {
    it('builds date range filter', () => {
      const o = toolSchemaOverrides.get('list-calendar-events')!;
      const q = o.queryTransform!({ startDate: '2025-03-01', endDate: '2025-03-31' });
      expect(q['$filter']).toContain("start/dateTime ge '2025-03-01T00:00:00Z'");
      expect(q['$filter']).toContain("end/dateTime le '2025-03-31T23:59:59Z'");
    });
  });

  describe('get-calendar-view queryTransform', () => {
    it('passes date range', () => {
      const o = toolSchemaOverrides.get('get-calendar-view')!;
      const q = o.queryTransform!({
        startDateTime: '2025-03-01T00:00:00Z',
        endDateTime: '2025-03-31T23:59:59Z',
      });
      expect(q.startDateTime).toBe('2025-03-01T00:00:00Z');
    });
  });

  // ── To-Do ─────────────────────────────────────────────────────────────────

  describe('create-todo-task transform', () => {
    const o = toolSchemaOverrides.get('create-todo-task')!;

    it('minimal', () => {
      expect(o.transform!({ title: 'Buy milk' })).toEqual({ title: 'Buy milk' });
    });

    it('full', () => {
      expect(
        o.transform!({ title: 'R', dueDate: '2025-03-15', notes: 'Q1', importance: 'high' })
      ).toEqual({
        title: 'R',
        dueDateTime: { dateTime: '2025-03-15T00:00:00', timeZone: 'UTC' },
        body: { content: 'Q1', contentType: 'text' },
        importance: 'high',
      });
    });
  });

  describe('list-todo-tasks queryTransform', () => {
    it('filters by status', () => {
      const o = toolSchemaOverrides.get('list-todo-tasks')!;
      expect(o.queryTransform!({ status: 'completed' })['$filter']).toBe(
        "status eq 'completed'"
      );
    });
  });

  // ── Planner ───────────────────────────────────────────────────────────────

  describe('create-planner-task transform', () => {
    it('handles assignedTo', () => {
      const o = toolSchemaOverrides.get('create-planner-task')!;
      const body = o.transform!({
        planId: 'p1',
        title: 'R',
        assignedTo: 'u1, u2',
      }) as Record<string, unknown>;
      expect(body.assignments).toEqual({
        u1: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
        u2: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      });
    });
  });

  // ── Contacts ──────────────────────────────────────────────────────────────

  describe('create-outlook-contact transform', () => {
    it('expands email and phone', () => {
      const o = toolSchemaOverrides.get('create-outlook-contact')!;
      const body = o.transform!({
        givenName: 'Jane',
        email: 'j@x.com',
        phone: '+1-555',
      }) as Record<string, unknown>;
      expect(body.emailAddresses).toEqual([{ address: 'j@x.com', name: '' }]);
      expect(body.businessPhones).toEqual(['+1-555']);
    });
  });

  describe('list-outlook-contacts queryTransform', () => {
    it('builds search', () => {
      const o = toolSchemaOverrides.get('list-outlook-contacts')!;
      expect(o.queryTransform!({ search: 'John' })['$search']).toBe('"John"');
    });
  });

  // ── Teams/Chat ────────────────────────────────────────────────────────────

  describe('send-chat-message transform', () => {
    it('wraps in body', () => {
      const o = toolSchemaOverrides.get('send-chat-message')!;
      expect(o.transform!({ content: 'Hi' })).toEqual({ body: { content: 'Hi' } });
    });
  });

  describe('list-chat-messages queryTransform', () => {
    it('sets $top', () => {
      const o = toolSchemaOverrides.get('list-chat-messages')!;
      expect(o.queryTransform!({ count: 50 })['$top']).toBe('50');
    });
  });

  // ── OneNote ───────────────────────────────────────────────────────────────

  describe('create-onenote-page transform', () => {
    it('wraps in HTML', () => {
      const o = toolSchemaOverrides.get('create-onenote-page')!;
      const body = o.transform!({ title: 'N', content: 'T' }) as Record<string, unknown>;
      expect(body.content).toContain('<title>N</title>');
    });
  });

  // ── Excel ─────────────────────────────────────────────────────────────────

  describe('create-excel-chart transform', () => {
    it('defaults seriesBy', () => {
      const o = toolSchemaOverrides.get('create-excel-chart')!;
      expect(o.transform!({ type: 'Line', sourceData: 'A1:C10' })).toEqual({
        type: 'Line',
        sourceData: 'A1:C10',
        seriesBy: 'Auto',
      });
    });
  });

  describe('format-excel-range transform', () => {
    it('builds font/fill', () => {
      const o = toolSchemaOverrides.get('format-excel-range')!;
      expect(o.transform!({ bold: true, fillColor: '#FF0' })).toEqual({
        font: { bold: true },
        fill: { color: '#FF0' },
      });
    });
  });

  describe('sort-excel-range transform', () => {
    it('wraps into fields', () => {
      const o = toolSchemaOverrides.get('sort-excel-range')!;
      expect(o.transform!({ columnIndex: 0, ascending: false })).toEqual({
        fields: [{ key: 0, ascending: false }],
      });
    });
  });

  // ── OneDrive ──────────────────────────────────────────────────────────────

  describe('upload-file-content transform', () => {
    it('returns raw string', () => {
      const o = toolSchemaOverrides.get('upload-file-content')!;
      expect(o.transform!({ content: 'data' })).toBe('data');
    });
  });

  // ── Search ────────────────────────────────────────────────────────────────

  describe('search-query transform', () => {
    it('builds requests array', () => {
      const o = toolSchemaOverrides.get('search-query')!;
      expect(
        o.transform!({ query: 'budget', entityTypes: 'message, driveItem', size: 10 })
      ).toEqual({
        requests: [
          { entityTypes: ['message', 'driveItem'], query: { queryString: 'budget' }, size: 10 },
        ],
      });
    });
  });

  // ── SharePoint ────────────────────────────────────────────────────────────

  describe('search-sharepoint-sites queryTransform', () => {
    it('passes search', () => {
      const o = toolSchemaOverrides.get('search-sharepoint-sites')!;
      expect(o.queryTransform!({ search: 'marketing' })).toEqual({ search: 'marketing' });
    });
  });

  describe('list-sharepoint-site-list-items queryTransform', () => {
    it('builds search + top', () => {
      const o = toolSchemaOverrides.get('list-sharepoint-site-list-items')!;
      const q = o.queryTransform!({ search: 'r', count: 15 });
      expect(q['$search']).toBe('"r"');
      expect(q['$top']).toBe('15');
    });
  });

  // ── Users ─────────────────────────────────────────────────────────────────

  describe('list-users queryTransform', () => {
    it('builds search', () => {
      const o = toolSchemaOverrides.get('list-users')!;
      expect(o.queryTransform!({ search: 'alice' })['$search']).toBe('"alice"');
    });
  });

  // ── Description-only overrides ────────────────────────────────────────────

  describe('description-only overrides', () => {
    const descOnly = [
      'list-mail-folders',
      'list-mail-attachments',
      'get-mail-attachment',
      'delete-mail-message',
      'delete-mail-attachment',
      'get-calendar-event',
      'list-calendars',
      'delete-calendar-event',
      'list-todo-task-lists',
      'get-todo-task',
      'delete-todo-task',
      'list-planner-tasks',
      'get-planner-plan',
      'list-plan-tasks',
      'get-planner-task',
      'get-outlook-contact',
      'delete-outlook-contact',
      'get-current-user',
      'list-joined-teams',
      'get-team',
      'list-team-channels',
      'get-team-channel',
      'list-team-members',
      'get-channel-message',
      'list-chats',
      'get-chat',
      'get-chat-message',
      'list-chat-message-replies',
      'list-drives',
      'get-drive-root-item',
      'list-folder-files',
      'download-onedrive-file-content',
      'delete-onedrive-file',
      'list-excel-worksheets',
      'get-excel-range',
      'list-onenote-notebooks',
      'list-onenote-notebook-sections',
      'list-onenote-section-pages',
      'get-onenote-page-content',
      'get-sharepoint-site',
      'list-sharepoint-site-drives',
      'get-sharepoint-site-drive-by-id',
      'list-sharepoint-site-items',
      'get-sharepoint-site-item',
      'list-sharepoint-site-lists',
      'get-sharepoint-site-list',
      'get-sharepoint-site-list-item',
      'get-sharepoint-site-by-path',
      'get-sharepoint-sites-delta',
    ];

    for (const name of descOnly) {
      it(`${name} has description only`, () => {
        const o = toolSchemaOverrides.get(name);
        expect(o, `${name} missing`).toBeDefined();
        expect(o!.description.length).toBeGreaterThan(10);
        expect(o!.transform).toBeUndefined();
        expect(o!.queryTransform).toBeUndefined();
      });
    }
  });
});
