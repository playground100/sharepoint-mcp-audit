// src/server.ts
import 'dotenv/config';
import express from 'express';
import type { Request, Response } from 'express';
import cors from 'cors';
import fetch from 'node-fetch';
import { ConfidentialClientApplication } from '@azure/msal-node';
import pLimit from 'p-limit';
import { z } from 'zod';

// MCP (Streamable HTTP) – Copilot Studio supports this transport
// Docs: https://github.com/modelcontextprotocol/typescript-sdk (Streamable HTTP)
// Copilot Studio uses Streamable transport and deprecates SSE. 
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';

const app = express();
app.use(express.json());

// If your server may be called from browser-hosted MCP clients, expose the header.
// For Copilot Studio, it's fine to leave origin '*' or restrict to your domain(s).
app.use(
  cors({
    origin: '*',
    exposedHeaders: ['Mcp-Session-Id'],
    allowedHeaders: ['Content-Type', 'mcp-session-id']
  })
);

// ---------- ENV ----------
const TENANT_ID = process.env.TENANT_ID!;
const CLIENT_ID = process.env.CLIENT_ID!;
const CLIENT_SECRET = process.env.CLIENT_SECRET!;
const TENANT_PRIMARY_HOST = process.env.TENANT_PRIMARY_HOST!;
const PORT = parseInt(process.env.PORT || '3000', 10);

for (const [k, v] of Object.entries({
  TENANT_ID, CLIENT_ID, CLIENT_SECRET, TENANT_PRIMARY_HOST
})) {
  if (!v) throw new Error(`Missing env var: ${k}`);
}

// ---------- MSAL (app-only) ----------
const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET
  }
});

async function getToken(scope: 'graph' | 'sharepoint'): Promise<string> {
  const resource = scope === 'graph'
    ? 'https://graph.microsoft.com/.default'
    : `https://${TENANT_PRIMARY_HOST}/.default`;
  const result = await msalApp.acquireTokenByClientCredential({ scopes: [resource] });
  if (!result?.accessToken) throw new Error(`Failed to acquire ${scope} token`);
  return result.accessToken;
}

// ---------- SharePoint/Graph helpers ----------
async function graphGet(url: string): Promise<any> {
  const token = await getToken('graph');
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph GET ${url} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

async function spPostJson(relativeUrl: string, body: any): Promise<any> {
  const token = await getToken('sharepoint');
  const url = `https://${TENANT_PRIMARY_HOST}${relativeUrl}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata'
    },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`SPO POST ${relativeUrl} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

async function spSiteGet(siteWebUrl: string, relativeUrl: string): Promise<any> {
  const token = await getToken('sharepoint');
  const url = `${siteWebUrl.replace(/\/$/, '')}${relativeUrl}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' }
  });
  if (!res.ok) throw new Error(`SPO GET ${relativeUrl} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// ---------- Search helpers ----------
type Row = Record<string, string>;
type SearchTable = { Rows: { Cells: { Key: string; Value: string }[] }[] };

async function searchInactiveFiles(inactivityDays: number): Promise<Row[]> {
  const thresholdISO = new Date(Date.now() - inactivityDays * 86400000).toISOString();
  const kql = `IsDocument:1 AND LastModifiedTime<="${thresholdISO}"`;
  const body = {
    request: {
      Querytext: kql,
      TrimDuplicates: false,
      RowLimit: 500,
      SelectProperties: { results: ['Title', 'Path', 'SPWebUrl', 'SiteTitle', 'ListId', 'LastModifiedTime', 'FileExtension'] },
      SourceId: '8413cd39-2156-4e00-b54d-11efd9abdb89'
    }
  };
  return searchAllPages(body);
}

async function searchByContentClass(contentClass: string): Promise<Row[]> {
  const body = {
    request: {
      Querytext: `contentclass:${contentClass}`,
      TrimDuplicates: false,
      RowLimit: 500,
      SelectProperties: { results: ['Title', 'Path', 'SPWebUrl', 'SiteTitle', 'ListId'] },
      SourceId: '8413cd39-2156-4e00-b54d-11efd9abdb89'
    }
  };
  return searchAllPages(body);
}

async function searchAllPages(body: any): Promise<Row[]> {
  const rows: Row[] = [];
  let startRow = 0;
  while (true) {
    const pageBody = structuredClone(body);
    pageBody.request.StartRow = startRow;
    const json = await spPostJson('/_api/search/postquery', pageBody);
    const table: SearchTable | undefined = json?.PrimaryQueryResult?.RelevantResults?.Table;
    const page = (table?.Rows || []).map(r => {
      const obj: Row = {};
      for (const c of r.Cells) obj[c.Key] = c.Value;
      return obj;
    });
    rows.push(...page);
    const total = json?.PrimaryQueryResult?.RelevantResults?.TotalRows || 0;
    startRow += page.length;
    if (startRow >= total || page.length === 0) break;
  }
  return rows;
}

async function getSiteOwners(siteWebUrl: string): Promise<string[]> {
  try {
    const json = await spSiteGet(siteWebUrl, '/_api/web/AssociatedOwnerGroup/Users');
    const users = json?.value ?? [];
    const names = users.map((u: any) => u.Title || u.LoginName).filter(Boolean);
    return names.length ? names : ['(no owners found)'];
  } catch {
    return ['(owners lookup failed)'];
  }
}

// ---------- Audit tool ----------
const AuditInput = z.object({ inactivityDays: z.number().int().positive().default(180) }).strict();

type AuditRecord =
  | { type: 'inactiveFile'; fileUrl: string; title?: string; siteUrl?: string; siteName?: string; listId?: string; lastModifiedTime?: string; extension?: string; owners?: string[] }
  | { type: 'documentLibrary'; title?: string; libraryUrl: string; siteUrl?: string; siteName?: string; listId?: string; owners?: string[] }
  | { type: 'list'; title?: string; listUrl: string; siteUrl?: string; siteName?: string; listId?: string; owners?: string[] };

async function runAudit(inactivityDays: number): Promise<AuditRecord[]> {
  const [files, libraries, lists] = await Promise.all([
    searchInactiveFiles(inactivityDays),
    searchByContentClass('STS_List_DocumentLibrary'),
    searchByContentClass('STS_List')
  ]);

  // collect unique sites for owner lookup
  const siteSet = new Set<string>();
  for (const r of [...files, ...libraries, ...lists]) if (r.SPWebUrl) siteSet.add(r.SPWebUrl);
  const limit = pLimit(6);
  const ownersMap = new Map<string, string[]>();
  await Promise.all(
    Array.from(siteSet).map(url => limit(async () => ownersMap.set(url, await getSiteOwners(url))))
  );

  // de-dupe “lists” that are libraries
  const libraryIds = new Set(libraries.map(l => l.ListId).filter(Boolean));
  const listOnly = lists.filter(l => l.ListId && !libraryIds.has(l.ListId!));

  const fileRecords: AuditRecord[] = files
    .filter(f => typeof f.Path === 'string')
    .map(f => {
      const record: AuditRecord = {
        type: 'inactiveFile',
        fileUrl: f.Path!
      };
      if (f.Title !== undefined) record.title = f.Title;
      if (f.SPWebUrl !== undefined) record.siteUrl = f.SPWebUrl;
      if (f.SiteTitle !== undefined) record.siteName = f.SiteTitle;
      if (f.ListId !== undefined) record.listId = f.ListId;
      if (f.LastModifiedTime !== undefined) record.lastModifiedTime = f.LastModifiedTime;
      if (f.FileExtension !== undefined) record.extension = f.FileExtension;
      const owners = ownersMap.get(f.SPWebUrl || '');
      if (owners !== undefined) record.owners = owners;
      return record;
    });

  const libraryRecords: AuditRecord[] = libraries.map(l => {
    const record: AuditRecord = {
      type: 'documentLibrary',
      title: l.Title ?? '',
      libraryUrl: l.Path ?? ''
    };
    if (l.SPWebUrl !== undefined) record.siteUrl = l.SPWebUrl;
    if (l.SiteTitle !== undefined) record.siteName = l.SiteTitle;
    if (l.ListId !== undefined) record.listId = l.ListId;
    const owners = ownersMap.get(l.SPWebUrl || '');
    if (owners !== undefined) record.owners = owners;
    return record;
  });

  const listRecords: AuditRecord[] = listOnly.map(l => {
    const record: AuditRecord = {
      type: 'list',
      title: l.Title ?? '',
      listUrl: l.Path ?? ''
    };
    if (l.SPWebUrl !== undefined) record.siteUrl = l.SPWebUrl;
    if (l.SiteTitle !== undefined) record.siteName = l.SiteTitle;
    if (l.ListId !== undefined) record.listId = l.ListId;
    const owners = ownersMap.get(l.SPWebUrl || '');
    if (owners !== undefined) record.owners = owners;
    return record;
  });

  return [...fileRecords, ...libraryRecords, ...listRecords];
}

// ---------- Build the MCP server instance ----------
function getMcpServer(): McpServer {
  const server = new McpServer({
    name: 'sharepoint-audit-mcp',
    version: '1.0.0',
    description: 'Audit SharePoint tenant for inactive files, libraries, and lists with site owners'
  });

  server.registerTool(
    'sharepoint_tenant_audit',
    {
      title: 'SharePoint Tenant Audit',
      description: 'Find inactive files (by inactivityDays), all document libraries and lists, with site owners',
      inputSchema: { inactivityDays: z.number().int().positive().default(180) }
    },
    async ({ inactivityDays }) => {
      const { inactivityDays: days } = AuditInput.parse({ inactivityDays });
      const results = await runAudit(days);
      return {
        content: [{ type: 'text', text: JSON.stringify({ inactivityDays: days, count: results.length, results }, null, 2) }]
      };
    }
  );

  return server;
}

// ---------- Streamable HTTP transport with session management ----------
import { randomUUID } from 'node:crypto';
import type { StreamableHTTPServerTransport as TTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';

const transports: Record<string, TTransport> = {};

app.post('/mcp', async (req: Request, res: Response) => {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;
  let transport: TTransport | undefined = sessionId ? transports[sessionId] : undefined;

  if (!transport && isInitializeRequest(req.body)) {
    transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID()
    });
    transport.onclose = () => {
      if (transport?.sessionId) delete transports[transport.sessionId];
    };
    const server = getMcpServer();
    await server.connect(transport);
    transports[transport.sessionId!] = transport;
  } else if (!transport) {
    res.status(400).json({
      jsonrpc: '2.0',
      error: { code: -32000, message: 'Bad Request: No valid session ID provided' },
      id: null
    });
    return;
  }

  await transport.handleRequest(req, res, req.body);
});

// Notifications (server → client) and session close:
app.get('/mcp', async (req, res) => {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;
  const transport = sessionId && transports[sessionId];
  if (!transport) return res.status(400).send('Invalid or missing session ID');
  await transport.handleRequest(req, res);
});

app.delete('/mcp', async (req, res) => {
  const sessionId = req.headers['mcp-session-id'] as string | undefined;
  const transport = sessionId && transports[sessionId];
  if (!transport) return res.status(400).send('Invalid or missing session ID');
  await transport.handleRequest(req, res);
});

// simple health
app.get('/healthz', (_req, res) => res.status(200).send('ok'));

app.listen(PORT, () => {
  console.log(`MCP server listening on :${PORT}`);
});
