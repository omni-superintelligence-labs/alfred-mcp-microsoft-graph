#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { fetch } from 'undici';

// Environment variables
const BACKEND_BASE_URL = process.env.BACKEND_BASE_URL || 'http://localhost:3000';
const SERVICE_JWT = process.env.SERVICE_JWT || '';

if (!SERVICE_JWT) {
  console.error('ERROR: SERVICE_JWT environment variable is required');
  process.exit(1);
}

// Tool definitions
const tools: Tool[] = [
  {
    name: 'excel_read_range',
    description: 'Read a range of cells from an Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: 'The ID of the Excel workbook' },
        worksheetName: { type: 'string', description: 'The name of the worksheet' },
        range: { type: 'string', description: 'The range to read, e.g., A1:C10' },
      },
      required: ['workbookId', 'worksheetName', 'range'],
    },
  },
  {
    name: 'excel_write_range',
    description: 'Write data to a range of cells in an Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        workbookId: { type: 'string', description: 'The ID of the Excel workbook' },
        worksheetName: { type: 'string', description: 'The name of the worksheet' },
        range: { type: 'string', description: 'The range to write to, e.g., A1:C10' },
        values: { type: 'array', description: '2D array of values to write' },
      },
      required: ['workbookId', 'worksheetName', 'range', 'values'],
    },
  },
  {
    name: 'outlook_list_messages',
    description: 'List recent email messages from Outlook',
    inputSchema: {
      type: 'object',
      properties: {
        top: { type: 'number', description: 'Number of messages to retrieve', default: 10 },
        filter: { type: 'string', description: 'OData filter query (optional)' },
      },
    },
  },
  {
    name: 'outlook_send_message',
    description: 'Send an email message via Outlook',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body (HTML or plain text)' },
        cc: { type: 'array', items: { type: 'string' }, description: 'CC recipients (optional)' },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'teams_send_chat_message',
    description: 'Send a message to a Microsoft Teams chat',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string', description: 'The ID of the Teams chat' },
        message: { type: 'string', description: 'The message to send' },
      },
      required: ['chatId', 'message'],
    },
  },
];

// Server setup
const server = new Server(
  {
    name: 'alfred-mcp-microsoft-graph',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List tools handler
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools };
});

// Call tool handler
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    // Call backend Graph action endpoint
    const response = await fetch(`${BACKEND_BASE_URL}/v1/graph/tools/${name}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${SERVICE_JWT}`,
      },
      body: JSON.stringify(args),
    });

    if (!response.ok) {
      const error = await response.text();
      return {
        content: [
          {
            type: 'text',
            text: `Error calling ${name}: ${error}`,
          },
        ],
        isError: true,
      };
    }

    const result = await response.json();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error: any) {
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${error.message}`,
        },
      ],
      isError: true,
    };
  }
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Alfred Microsoft Graph MCP server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});

