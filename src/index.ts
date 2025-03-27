#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';

const ListResourcesRequestSchema = z.object({
  jsonrpc: z.string(),
  id: z.any(),
  method: z.literal('resources/list'),
  params: z.record(z.any())
});

const ListPromptsRequestSchema = z.object({
  jsonrpc: z.string(),
  id: z.any(),
  method: z.literal('prompts/list'),
  params: z.record(z.any())
});

// Environment variables required for OAuth
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;

if (!CLIENT_ID || !CLIENT_SECRET || !REFRESH_TOKEN) {
  throw new Error('Required Google OAuth credentials not found in environment variables');
}

class GoogleWorkspaceServer {
  private server: Server;
  private auth;
  private gmail;
  private calendar;
  private calendarNameToId: Map<string, string> = new Map();

  constructor() {
    this.server = new Server(
      {
        name: 'google-workspace-server',
        version: '0.1.0',
      },
      {
        capabilities: {
          tools: {},
          resources: {
            supported: true
          },
          prompts: {
            supported: true
          }
        },
      }
    );

    // Set up OAuth2 client
    this.auth = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET);
    this.auth.setCredentials({ refresh_token: REFRESH_TOKEN });

    // Initialize API clients
    this.gmail = google.gmail({ version: 'v1', auth: this.auth });
    this.calendar = google.calendar({ version: 'v3', auth: this.auth });

    this.setupToolHandlers();
    this.setupAdditionalHandlers();
    this.initializeCalendarMap();

    // Error handling
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_emails',
          description: 'List recent emails from Gmail inbox',
          inputSchema: {
            type: 'object',
            properties: {
              maxResults: {
                type: 'number',
                description: 'Maximum number of emails to return (default: 10)',
              },
              query: {
                type: 'string',
                description: 'Search query to filter emails',
              },
            },
          },
        },
        {
          name: 'search_emails',
          description: 'Search emails with advanced query',
          inputSchema: {
            type: 'object',
            properties: {
              query: {
                type: 'string',
                description: 'Gmail search query (e.g., "from:example@gmail.com has:attachment")',
                required: true
              },
              maxResults: {
                type: 'number',
                description: 'Maximum number of emails to return (default: 10)',
              },
            },
            required: ['query']
          },
        },
        {
          name: 'send_email',
          description: 'Send a new email',
          inputSchema: {
            type: 'object',
            properties: {
              to: {
                type: 'string',
                description: 'Recipient email address',
              },
              subject: {
                type: 'string',
                description: 'Email subject',
              },
              body: {
                type: 'string',
                description: 'Email body (can include HTML)',
              },
              cc: {
                type: 'string',
                description: 'CC recipients (comma-separated)',
              },
              bcc: {
                type: 'string',
                description: 'BCC recipients (comma-separated)',
              },
            },
            required: ['to', 'subject', 'body']
          },
        },
        {
          name: 'modify_email',
          description: 'Modify email labels (archive, trash, mark read/unread)',
          inputSchema: {
            type: 'object',
            properties: {
              id: {
                type: 'string',
                description: 'Email ID',
              },
              addLabels: {
                type: 'array',
                items: { type: 'string' },
                description: 'Labels to add',
              },
              removeLabels: {
                type: 'array',
                items: { type: 'string' },
                description: 'Labels to remove',
              },
            },
            required: ['id']
          },
        },
        {
          name: 'list_events',
          description: 'List upcoming calendar events',
          inputSchema: {
            type: 'object',
            properties: {
              maxResults: {
                type: 'number',
                description: 'Maximum number of events to return (default: 10)',
              },
              timeMin: {
                type: 'string',
                description: 'Start time in ISO format (default: now)',
              },
              timeMax: {
                type: 'string',
                description: 'End time in ISO format',
              },
            },
          },
        },
        {
          name: 'create_event',
          description: 'Create a new calendar event',
          inputSchema: {
            type: 'object',
            properties: {
              summary: {
                type: 'string',
                description: 'Event title',
              },
              location: {
                type: 'string',
                description: 'Event location',
              },
              description: {
                type: 'string',
                description: 'Event description',
              },
              start: {
                type: 'string',
                description: 'Start time in ISO format',
              },
              end: {
                type: 'string',
                description: 'End time in ISO format',
              },
              attendees: {
                type: 'array',
                items: { type: 'string' },
                description: 'List of attendee email addresses',
              },
              calendarId: {
                type: 'string',
                description: 'Calendar ID to create the event in (default: "primary")',
              },
            },
            required: ['summary', 'start', 'end']
          },
        },
        {
          name: 'update_event',
          description: 'Update an existing calendar event',
          inputSchema: {
            type: 'object',
            properties: {
              eventId: {
                type: 'string',
                description: 'Event ID to update',
              },
              calendarId: {
                type: 'string',
                description: 'Calendar ID containing the event (default: "primary")',
              },
              summary: {
                type: 'string',
                description: 'New event title',
              },
              location: {
                type: 'string',
                description: 'New event location',
              },
              description: {
                type: 'string',
                description: 'New event description',
              },
              start: {
                type: 'string',
                description: 'New start time in ISO format',
              },
              end: {
                type: 'string',
                description: 'New end time in ISO format',
              },
              attendees: {
                type: 'array',
                items: { type: 'string' },
                description: 'New list of attendee email addresses',
              },
            },
            required: ['eventId']
          },
        },
        {
          name: 'delete_event',
          description: 'Delete a calendar event',
          inputSchema: {
            type: 'object',
            properties: {
              eventId: {
                type: 'string',
                description: 'Event ID to delete',
              },
              calendarId: {
                type: 'string',
                description: 'Calendar ID containing the event (default: "primary")',
              },
            },
            required: ['eventId']
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      switch (request.params.name) {
        case 'list_emails':
          return await this.handleListEmails(request.params.arguments);
        case 'search_emails':
          return await this.handleSearchEmails(request.params.arguments);
        case 'send_email':
          return await this.handleSendEmail(request.params.arguments);
        case 'modify_email':
          return await this.handleModifyEmail(request.params.arguments);
        case 'list_events':
          return await this.handleListEvents(request.params.arguments);
        case 'create_event':
          return await this.handleCreateEvent(request.params.arguments);
        case 'update_event':
          return await this.handleUpdateEvent(request.params.arguments);
        case 'delete_event':
          return await this.handleDeleteEvent(request.params.arguments);
        default:
          throw new McpError(
            ErrorCode.MethodNotFound,
            `Unknown tool: ${request.params.name}`
          );
      }
    });
  }

  private setupAdditionalHandlers() {
    // resources/list
    this.server.setRequestHandler(ListResourcesRequestSchema, async (request) => {
      const defaultParams = { page: 1, limit: 10 };
      const finalParams = Object.assign({}, defaultParams, request.params);

      // finalParamsを元にリソース情報を取得して返す
      return {
        resources: [
          // 実際のリソース情報を取得するコードをここに追加
          // 例: { id: 'resource-1', name: 'Some resource' }
        ],
        paramsUsed: finalParams
      };
    });

    // prompts/list
    this.server.setRequestHandler(ListPromptsRequestSchema, async (request) => {
      const defaultParams = { page: 1, limit: 10 };
      const finalParams = Object.assign({}, defaultParams, request.params);

      // finalParamsを元にプロンプト情報を取得して返す
      return {
        prompts: [
          // 実際のプロンプト情報を取得するコードをここに追加
          // 例: { id: 'prompt-1', text: 'Some prompt text' }
        ],
        paramsUsed: finalParams
      };
    });
  }

  private async handleListEmails(args: any) {
    try {
      const maxResults = args?.maxResults || 10;
      const query = args?.query || '';

      const response = await this.gmail.users.messages.list({
        userId: 'me',
        maxResults,
        q: query,
      });

      const messages = response.data.messages || [];
      const emailDetails = await Promise.all(
        messages.map(async (msg) => {
          const detail = await this.gmail.users.messages.get({
            userId: 'me',
            id: msg.id!,
          });

          const headers = detail.data.payload?.headers;
          const subject = headers?.find((h) => h.name === 'Subject')?.value || '';
          const from = headers?.find((h) => h.name === 'From')?.value || '';
          const date = headers?.find((h) => h.name === 'Date')?.value || '';

          return {
            id: msg.id,
            subject,
            from,
            date,
          };
        })
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(emailDetails, null, 2),
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching emails: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSearchEmails(args: any) {
    try {
      const maxResults = args?.maxResults || 10;
      const query = args?.query || '';

      const response = await this.gmail.users.messages.list({
        userId: 'me',
        maxResults,
        q: query,
      });

      const messages = response.data.messages || [];
      const emailDetails = await Promise.all(
        messages.map(async (msg) => {
          const detail = await this.gmail.users.messages.get({
            userId: 'me',
            id: msg.id!,
          });

          const headers = detail.data.payload?.headers;
          const subject = headers?.find((h) => h.name === 'Subject')?.value || '';
          const from = headers?.find((h) => h.name === 'From')?.value || '';
          const date = headers?.find((h) => h.name === 'Date')?.value || '';

          return {
            id: msg.id,
            subject,
            from,
            date,
          };
        })
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(emailDetails, null, 2),
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching emails: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSendEmail(args: any) {
    try {
      const { to, subject, body, cc, bcc } = args;

      // Create email content
      const message = [
        'Content-Type: text/html; charset=utf-8',
        'MIME-Version: 1.0',
        `To: ${to}`,
        cc ? `Cc: ${cc}` : '',
        bcc ? `Bcc: ${bcc}` : '',
        `Subject: ${subject}`,
        '',
        body,
      ].filter(Boolean).join('\r\n');

      // Encode the email
      const encodedMessage = Buffer.from(message)
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');

      // Send the email
      const response = await this.gmail.users.messages.send({
        userId: 'me',
        requestBody: {
          raw: encodedMessage,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email sent successfully. Message ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error sending email: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleModifyEmail(args: any) {
    try {
      const { id, addLabels = [], removeLabels = [] } = args;

      const response = await this.gmail.users.messages.modify({
        userId: 'me',
        id,
        requestBody: {
          addLabelIds: addLabels,
          removeLabelIds: removeLabels,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email modified successfully. Updated labels for message ID: ${response.data.id}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error modifying email: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async initializeCalendarMap() {
    try {
      const response = await this.calendar.calendarList.list();
      const calendars = response.data.items || [];
      console.error('Available calendars:', calendars.map(c => ({ name: c.summary, id: c.id })));

      calendars.forEach(calendar => {
        if (calendar.id && calendar.summary) {
          // カレンダー名を小文字に変換して保存
          const name = calendar.summary.toLowerCase();
          this.calendarNameToId.set(name, calendar.id);
          // ドットを除去したバージョンも保存
          this.calendarNameToId.set(name.replace(/\./g, ''), calendar.id);
          // スペースを除去したバージョンも保存
          this.calendarNameToId.set(name.replace(/\s+/g, ''), calendar.id);
        }
      });
    } catch (error) {
      console.error('Failed to initialize calendar map:', error);
    }
  }

  private getCalendarId(calendarNameOrId: string): string {
    const searchName = calendarNameOrId.toLowerCase();

    // まず完全一致で検索
    if (this.calendarNameToId.has(searchName)) {
      return this.calendarNameToId.get(searchName)!;
    }

    // ドットを除去したバージョンで検索
    const noDotName = searchName.replace(/\./g, '');
    if (this.calendarNameToId.has(noDotName)) {
      return this.calendarNameToId.get(noDotName)!;
    }

    // スペースを除去したバージョンで検索
    const noSpaceName = searchName.replace(/\s+/g, '');
    if (this.calendarNameToId.has(noSpaceName)) {
      return this.calendarNameToId.get(noSpaceName)!;
    }

    // 部分一致で検索
    for (const [name, id] of this.calendarNameToId.entries()) {
      if (name.includes(searchName) || searchName.includes(name)) {
        return id;
      }
    }

    // 見つからない場合は、入力された値をそのまま使用
    return calendarNameOrId;
  }

  private convertToJST(dateTimeStr: string | null): string | null {
    if (!dateTimeStr) return null;
    const date = new Date(dateTimeStr);
    return date.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  }

  private convertToUTC(dateTimeStr: string): string {
    let date: Date;

    // タイムゾーン指定がない場合は日本時間(JST)として処理
    if (!/Z|\+|\-/.test(dateTimeStr)) {
      date = new Date(`${dateTimeStr}+09:00`);
    } else {
      date = new Date(dateTimeStr);
    }

    if (isNaN(date.getTime())) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        'Invalid date format. Please use a valid ISO 8601 format (e.g., "2025-03-20T10:00:00+09:00")'
      );
    }
    return date.toISOString();
  }

  private async handleCreateEvent(args: any) {
    try {
      const {
        summary,
        location,
        description,
        start,
        end,
        attendees,
        calendarId = 'primary'
      } = args;

      const resolvedCalendarId = this.getCalendarId(calendarId);

      // Validate calendar ID
      try {
        await this.calendar.calendars.get({ calendarId: resolvedCalendarId });
      } catch (error) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `Invalid calendar ID or name: ${calendarId}`
        );
      }

      // Validate date formats and convert to UTC
      const startDate = new Date(start);
      const endDate = new Date(end);

      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'Invalid date format. Please use ISO 8601 format (e.g., "2024-03-20T10:00:00Z")'
        );
      }

      if (endDate <= startDate) {
        throw new McpError(
          ErrorCode.InvalidRequest,
          'End time must be after start time'
        );
      }

      const event = {
        summary,
        location,
        description,
        start: {
          dateTime: this.convertToUTC(start),
          timeZone: 'UTC',
        },
        end: {
          dateTime: this.convertToUTC(end),
          timeZone: 'UTC',
        },
        attendees: attendees?.map((email: string) => ({ email })),
      };

      const response = await this.calendar.events.insert({
        calendarId: resolvedCalendarId,
        requestBody: event,
        sendUpdates: 'all',
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              eventId: response.data.id,
              htmlLink: response.data.htmlLink,
              start: {
                utc: response.data!.start!.dateTime!,
                jst: this.convertToJST(response.data!.start!.dateTime!)
              },
              end: {
                utc: response.data!.end!.dateTime!,
                jst: this.convertToJST(response.data!.end!.dateTime!)
              }
            }, null, 2)
          }
        ]
      };
    } catch (error) {
      if (error instanceof McpError) {
        throw error;
      }
      throw new McpError(
        ErrorCode.InternalError,
        `Failed to create event: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
  }

  private async handleUpdateEvent(args: any) {
    try {
      const { eventId, calendarId = 'primary', summary, location, description, start, end, attendees } = args;

      const event: any = {};
      if (summary) event.summary = summary;
      if (location) event.location = location;
      if (description) event.description = description;
      if (start) {
        event.start = {
          dateTime: this.convertToUTC(start),
          timeZone: 'UTC',
        };
      }
      if (end) {
        event.end = {
          dateTime: this.convertToUTC(end),
          timeZone: 'UTC',
        };
      }
      if (attendees) {
        event.attendees = attendees.map((email: string) => ({ email }));
      }

      const response = await this.calendar.events.patch({
        calendarId,
        eventId,
        requestBody: event,
      });

      return {
        content: [
          {
            type: 'text',
            text: `Event updated successfully in calendar ${calendarId}. Event ID: ${response.data.id}`,
          },
        ],
        start: {
          utc: response.data!.start!.dateTime!,
          jst: this.convertToJST(response.data!.start!.dateTime!)
        },
        end: {
          utc: response.data!.end!.dateTime!,
          jst: this.convertToJST(response.data!.end!.dateTime!)
        }
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error updating event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDeleteEvent(args: any) {
    try {
      const { eventId, calendarId = 'primary' } = args;

      await this.calendar.events.delete({
        calendarId,
        eventId,
      });

      return {
        content: [
          {
            type: 'text',
            text: `Event deleted successfully from calendar ${calendarId}. Event ID: ${eventId}`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error deleting event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleListEvents(args: any) {
    try {
      const maxResults = args?.maxResults || 15;
      const timeMin = args?.timeMin || new Date().toISOString();
      const timeMax = args?.timeMax;

      // まずアカウントに紐付いた全カレンダーのリストを取得
      const calendarListResponse = await this.calendar.calendarList.list();
      const calendars = calendarListResponse.data.items || [];

      let allEvents: any[] = [];

      // 各カレンダーごとにイベントを取得
      for (const calendarItem of calendars) {
        const calId = calendarItem.id!;
        const eventsResponse = await this.calendar.events.list({
          calendarId: calId,
          timeMin,
          timeMax,
          singleEvents: true,
          orderBy: 'startTime', // API側でもある程度並びますが、全体での整合性のために後で手動ソート
        });

        // JST に変換する関数の例
        const convertToJST = (dateTimeStr: string | null) => {
          if (!dateTimeStr) return null;
          const date = new Date(dateTimeStr);
          return date.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
        };

        const events = eventsResponse.data.items || [];
        allEvents = allEvents.concat(
          events.map((event) => ({
            calendarId: calendarItem.id,
            id: event.id,
            summary: event.summary,
            start: {
              ...event.start,
              jst: event.start!.dateTime ? convertToJST(event.start!.dateTime) : convertToJST(event.start!.date!)
            },
            end: {
              ...event.end,
              jst: event.end!.dateTime ? convertToJST(event.end!.dateTime) : convertToJST(event.end!.date!)
            },
            location: event.location,
          }))
        );
      }

      // すべてのイベントを開始時刻（dateTime もしくは date）で昇順にソート
      allEvents.sort((a, b) => {
        const aTime = new Date(a.start.dateTime || a.start.date);
        const bTime = new Date(b.start.dateTime || b.start.date);
        return aTime.getTime() - bTime.getTime();
      });

      // 指定された件数だけ抽出
      const limitedEvents = allEvents.slice(0, maxResults);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(limitedEvents, null, 2),
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching calendar events: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Google Workspace MCP server running on stdio');
  }
}

const server = new GoogleWorkspaceServer();
server.run().catch(console.error);
