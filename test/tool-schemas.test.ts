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
  // ── Coverage: every single endpoint must have an override ─────────────────

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

  // ── Structure validation ──────────────────────────────────────────────────

  describe('structure', () => {
    it('every override has a meaningful description', () => {
      for (const [name, o] of toolSchemaOverrides) {
        expect(o.description, `${name} missing description`).toBeTruthy();
        expect(o.description.length, `${name} description too short`).toBeGreaterThan(10);
      }
    });

    it('write endpoints (POST/PUT/PATCH) have schema + transform', () => {
      const writes = endpointsData.filter((e) =>
        ['post', 'put', 'patch'].includes(e.method)
      );
      for (const ep of writes) {
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
          expect(s instanceof z.ZodType, `${name}.${key} not ZodType`).toBe(true);
        }
      }
    });
  });

  // ── Mail transforms ───────────────────────────────────────────────────────

  describe('send-mail transform', () => {
    const o = toolSchemaOverrides.get('send-mail')!;

    it('produces correct API body', () => {
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

    it('handles HTML + CC + BCC', () => {
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
  });

  describe('create-draft-email transform', () => {
    const o = toolSchemaOverrides.get('create-draft-email')!;

    it('produces correct body', () => {
      expect(o.transform!({ subject: 'D', content: 'Hi', to: 'a@b.com' })).toEqual({
        subject: 'D',
        body: { contentType: 'Text', content: 'Hi' },
        toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
      });
    });
  });

  describe('move-mail-message transform', () => {
    it('passes through destinationId', () => {
      const o = toolSchemaOverrides.get('move-mail-message')!;
      expect(o.transform!({ destinationId: 'inbox' })).toEqual({
        destinationId: 'inbox',
      });
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
      expect(body.name).toBe('f.pdf');
    });
  });

  // ── Mail list queryTransform ──────────────────────────────────────────────

  describe('list-mail-messages queryTransform', () => {
    const o = toolSchemaOverrides.get('list-mail-messages')!;

    it('builds OData from simple params', () => {
      const q = o.queryTransform!({
        from: 'alice@example.com',
        unreadOnly: true,
        count: 5,
      });
      expect(q['$filter']).toContain("from/emailAddress/address eq 'alice@example.com'");
      expect(q['$filter']).toContain('isRead eq false');
      expect(q['$top']).toBe('5');
      expect(q['$orderby']).toBe('receivedDateTime desc');
    });

    it('defaults to 10 results', () => {
      expect(o.queryTransform!({})['$top']).toBe('10');
    });

    it('handles search', () => {
      expect(o.queryTransform!({ search: 'budget' })['$search']).toBe('"budget"');
    });
  });

  // ── Calendar transforms ───────────────────────────────────────────────────

  describe('create-calendar-event transform', () => {
    const o = toolSchemaOverrides.get('create-calendar-event')!;

    it('builds nested start/end with timeZone default', () => {
      const body = o.transform!({
        subject: 'Standup',
        startDateTime: '2025-03-15T09:00:00',
        endDateTime: '2025-03-15T09:15:00',
      }) as Record<string, unknown>;
      expect(body.subject).toBe('Standup');
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
      expect(body.onlineMeetingProvider).toBe('teamsForBusiness');
      expect((body.attendees as unknown[]).length).toBe(2);
    });
  });

  describe('list-calendar-events queryTransform', () => {
    const o = toolSchemaOverrides.get('list-calendar-events')!;

    it('builds date range filter', () => {
      const q = o.queryTransform!({
        startDate: '2025-03-01',
        endDate: '2025-03-31',
        count: 20,
      });
      expect(q['$filter']).toContain("start/dateTime ge '2025-03-01T00:00:00Z'");
      expect(q['$filter']).toContain("end/dateTime le '2025-03-31T23:59:59Z'");
      expect(q['$top']).toBe('20');
    });
  });

  describe('get-calendar-view queryTransform', () => {
    const o = toolSchemaOverrides.get('get-calendar-view')!;

    it('passes startDateTime and endDateTime', () => {
      const q = o.queryTransform!({
        startDateTime: '2025-03-01T00:00:00Z',
        endDateTime: '2025-03-31T23:59:59Z',
      });
      expect(q.startDateTime).toBe('2025-03-01T00:00:00Z');
      expect(q.endDateTime).toBe('2025-03-31T23:59:59Z');
    });
  });

  // ── To-Do transforms ─────────────────────────────────────────────────────

  describe('create-todo-task transform', () => {
    const o = toolSchemaOverrides.get('create-todo-task')!;

    it('minimal: just title', () => {
      expect(o.transform!({ title: 'Buy milk' })).toEqual({ title: 'Buy milk' });
    });

    it('full: with dueDate and notes', () => {
      expect(
        o.transform!({
          title: 'Report',
          dueDate: '2025-03-15',
          notes: 'Q1 financials',
          importance: 'high',
        })
      ).toEqual({
        title: 'Report',
        dueDateTime: { dateTime: '2025-03-15T00:00:00', timeZone: 'UTC' },
        body: { content: 'Q1 financials', contentType: 'text' },
        importance: 'high',
      });
    });
  });

  describe('list-todo-tasks queryTransform', () => {
    const o = toolSchemaOverrides.get('list-todo-tasks')!;

    it('filters by status', () => {
      const q = o.queryTransform!({ status: 'completed', count: 5 });
      expect(q['$filter']).toBe("status eq 'completed'");
      expect(q['$top']).toBe('5');
    });
  });

  // ── Planner transforms ───────────────────────────────────────────────────

  describe('create-planner-task transform', () => {
    const o = toolSchemaOverrides.get('create-planner-task')!;

    it('handles assignedTo', () => {
      const body = o.transform!({
        planId: 'plan-1',
        title: 'Review',
        assignedTo: 'user-1, user-2',
      }) as Record<string, unknown>;
      expect(body.assignments).toEqual({
        'user-1': {
          '@odata.type': '#microsoft.graph.plannerAssignment',
          orderHint: ' !',
        },
        'user-2': {
          '@odata.type': '#microsoft.graph.plannerAssignment',
          orderHint: ' !',
        },
      });
    });
  });

  // ── Contact transforms ───────────────────────────────────────────────────

  describe('create-outlook-contact transform', () => {
    const o = toolSchemaOverrides.get('create-outlook-contact')!;

    it('expands email and phone', () => {
      const body = o.transform!({
        givenName: 'Jane',
        email: 'jane@x.com',
        phone: '+1-555-0100',
        company: 'Acme',
      }) as Record<string, unknown>;
      expect(body.emailAddresses).toEqual([{ address: 'jane@x.com', name: '' }]);
      expect(body.businessPhones).toEqual(['+1-555-0100']);
      expect(body.companyName).toBe('Acme');
    });
  });

  describe('list-outlook-contacts queryTransform', () => {
    const o = toolSchemaOverrides.get('list-outlook-contacts')!;

    it('builds search query', () => {
      const q = o.queryTransform!({ search: 'John', count: 10 });
      expect(q['$search']).toBe('"John"');
      expect(q['$top']).toBe('10');
    });
  });

  // ── Teams/Chat transforms ────────────────────────────────────────────────

  describe('send-chat-message transform', () => {
    it('wraps in body object', () => {
      const o = toolSchemaOverrides.get('send-chat-message')!;
      expect(o.transform!({ content: 'Hello!' })).toEqual({
        body: { content: 'Hello!' },
      });
    });
  });

  describe('list-chat-messages queryTransform', () => {
    it('sets $top', () => {
      const o = toolSchemaOverrides.get('list-chat-messages')!;
      expect(o.queryTransform!({ count: 50 })['$top']).toBe('50');
    });
  });

  // ── OneNote transforms ────────────────────────────────────────────────────

  describe('create-onenote-page transform', () => {
    it('wraps in HTML structure', () => {
      const o = toolSchemaOverrides.get('create-onenote-page')!;
      const body = o.transform!({ title: 'Notes', content: 'Text' }) as Record<
        string,
        unknown
      >;
      expect(body.contentType).toBe('text/html');
      expect(body.content).toContain('<title>Notes</title>');
      expect(body.content).toContain('Text');
    });
  });

  // ── Excel transforms ─────────────────────────────────────────────────────

  describe('create-excel-chart transform', () => {
    it('passes through with default seriesBy', () => {
      const o = toolSchemaOverrides.get('create-excel-chart')!;
      expect(o.transform!({ type: 'Line', sourceData: 'A1:C10' })).toEqual({
        type: 'Line',
        sourceData: 'A1:C10',
        seriesBy: 'Auto',
      });
    });
  });

  describe('format-excel-range transform', () => {
    it('builds font/fill from flat params', () => {
      const o = toolSchemaOverrides.get('format-excel-range')!;
      expect(
        o.transform!({ bold: true, fontSize: 14, fillColor: '#FFFF00' })
      ).toEqual({
        font: { bold: true, size: 14 },
        fill: { color: '#FFFF00' },
      });
    });
  });

  describe('sort-excel-range transform', () => {
    it('wraps into fields array', () => {
      const o = toolSchemaOverrides.get('sort-excel-range')!;
      expect(
        o.transform!({ columnIndex: 2, ascending: false, hasHeaders: true })
      ).toEqual({
        fields: [{ key: 2, ascending: false }],
        hasHeaders: true,
      });
    });
  });

  // ── OneDrive ──────────────────────────────────────────────────────────────

  describe('upload-file-content transform', () => {
    it('returns raw content string', () => {
      const o = toolSchemaOverrides.get('upload-file-content')!;
      expect(o.transform!({ content: 'raw data' })).toBe('raw data');
    });
  });

  // ── Search transform ──────────────────────────────────────────────────────

  describe('search-query transform', () => {
    it('builds nested requests array', () => {
      const o = toolSchemaOverrides.get('search-query')!;
      expect(
        o.transform!({ query: 'budget', entityTypes: 'message, driveItem', size: 10 })
      ).toEqual({
        requests: [
          {
            entityTypes: ['message', 'driveItem'],
            query: { queryString: 'budget' },
            size: 10,
          },
        ],
      });
    });
  });

  // ── SharePoint queryTransform ─────────────────────────────────────────────

  describe('search-sharepoint-sites queryTransform', () => {
    it('passes search param', () => {
      const o = toolSchemaOverrides.get('search-sharepoint-sites')!;
      expect(o.queryTransform!({ search: 'marketing' })).toEqual({
        search: 'marketing',
      });
    });
  });

  describe('list-sharepoint-site-list-items queryTransform', () => {
    it('builds search + top', () => {
      const o = toolSchemaOverrides.get('list-sharepoint-site-list-items')!;
      const q = o.queryTransform!({ search: 'report', count: 15 });
      expect(q['$search']).toBe('"report"');
      expect(q['$top']).toBe('15');
    });
  });

  // ── Users queryTransform ──────────────────────────────────────────────────

  describe('list-users queryTransform', () => {
    it('builds search + select', () => {
      const o = toolSchemaOverrides.get('list-users')!;
      const q = o.queryTransform!({ search: 'alice' });
      expect(q['$search']).toBe('"alice"');
      expect(q['$select']).toContain('displayName');
    });
  });

  // ── Description-only overrides ────────────────────────────────────────────

  describe('description-only overrides', () => {
    const descOnly = [
      'get-mail-message',
      'get-shared-mailbox-message',
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
      'get-root-folder',
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
      it(`${name} has description and no transform`, () => {
        const o = toolSchemaOverrides.get(name);
        expect(o, `${name} missing override`).toBeDefined();
        expect(o!.description.length).toBeGreaterThan(10);
        expect(o!.transform).toBeUndefined();
        expect(o!.queryTransform).toBeUndefined();
      });
    }
  });
});
