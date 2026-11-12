#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { DeviceCodeCredential } from '@azure/identity';
import fetch from 'node-fetch';
import { JSDOM } from 'jsdom';
import { z } from 'zod';

// Load environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Create the MCP server
const server = new McpServer(
  { 
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server" 
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

// Try to read the stored access token
let accessToken = null;
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

let graphClient = null;

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID
const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

// Function to ensure Graph client is created
async function ensureGraphClient() {
  if (!graphClient) {
    // Read token from file if it exists
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
        try {
          // Try to parse as JSON first (new format)
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.token;
        } catch (parseError) {
          // Fall back to using the raw token (old format)
          accessToken = tokenData;
        }
      }
    } catch (error) {
      console.error("Error reading token file:", error);
    }

    if (!accessToken) {
      throw new Error("Access token not found. Please save access token first.");
    }

    // Create Microsoft Graph client
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
  }
  return graphClient;
}

// Create graph client with device code auth or access token
async function createGraphClient() {
  if (accessToken) {
    // Use access token if available
    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });
    return { type: 'token', client: graphClient };
  } else {
    // Use device code flow
    const credential = new DeviceCodeCredential({
      clientId: clientId,
      userPromptCallback: (info) => {
        // This will be shown to the user with the URL and code
        console.error('\n' + info.message);
      }
    });

    try {
      // Get an access token using device code flow
      const tokenResponse = await credential.getToken(scopes);
      
      // Save the token for future use
      accessToken = tokenResponse.token;
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));
      
      // Initialize Graph client with the token
      graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return accessToken;
          }
        }
      });
      
      return { type: 'device_code', client: graphClient };
    } catch (error) {
      console.error('Authentication error:', error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
}

// Tool for starting authentication flow
server.tool(
  "authenticate",
  "Start the authentication flow with Microsoft Graph",
  async () => {
    try {
      const result = await createGraphClient();
      if (result.type === 'device_code') {
        return { 
          content: [
            {
              type: "text",
              text: "Authentication started. Please check the console for the URL and code."
            }
          ]
        };
      } else {
        return { 
          content: [
            {
              type: "text",
              text: "Already authenticated with an access token."
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in authentication:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// Tool for saving an access token provided by the user
server.tool(
  "saveAccessToken",
  "Save a Microsoft Graph access token for later use",
  async (params) => {
    try {
      // Save the token for future use
      accessToken = params.random_string;
      const tokenData = JSON.stringify({ token: accessToken });
      fs.writeFileSync(tokenFilePath, tokenData);
      await createGraphClient();
      return { 
        content: [
          {
            type: "text",
            text: "Access token saved successfully"
          }
        ]
      };
    } catch (error) {
      console.error("Error saving access token:", error);
      throw new Error(`Failed to save access token: ${error.message}`);
    }
  }
);

// Tool for listing all notebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      // Return content as an array of text items
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/notebooks`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value[0])
          }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections in a notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/sections`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        return { 
          content: [
            {
              type: "text",
              text: "[]"
            }
          ]
        };
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      console.error("GetPage called with params:", params);
      await ensureGraphClient();
      
      // First, list all pages to find the one we want
      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      console.error("Got", pagesResponse.value.length, "pages");
      
      let targetPage;
      
      // If a page ID is provided, use it to find the page
      if (params.random_string && params.random_string.length > 0) {
        const pageId = params.random_string;
        console.error("Looking for page with ID:", pageId);
        
        // Look for exact match first
        targetPage = pagesResponse.value.find(p => p.id === pageId);
        
        // If no exact match, try matching by title
        if (!targetPage) {
          console.error("No exact match, trying title search");
          targetPage = pagesResponse.value.find(p => 
            p.title && p.title.toLowerCase().includes(params.random_string.toLowerCase())
          );
        }
        
        // If still no match, try partial ID match
        if (!targetPage) {
          console.error("No title match, trying partial ID match");
          targetPage = pagesResponse.value.find(p => 
            p.id.includes(pageId) || pageId.includes(p.id)
          );
        }
      } else {
        // If no ID provided, use the first page
        console.error("No ID provided, using first page");
        targetPage = pagesResponse.value[0];
      }
      
      if (!targetPage) {
        throw new Error("Page not found");
      }
      
      console.error("Target page found:", targetPage.title);
      console.error("Page ID:", targetPage.id);
      
      try {
        const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
        console.error("Fetching content from:", url);
        
        // Make direct HTTP request with fetch
        const response = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });
        
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
        }
        
        const content = await response.text();
        console.error(`Content received! Length: ${content.length} characters`);
        
        // Return the raw HTML content
        return {
          content: [
            {
              type: "text",
              text: content
            }
          ]
        };
      } catch (error) {
        console.error("Error getting content:", error);
        
        // Return a simple error message
        return {
          content: [
            {
              type: "text",
              text: `Error retrieving page content: ${error.message}`
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in getPage:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error in getPage: ${error.message}`
          }
        ]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        throw new Error("No sections found");
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      
      // Create simple HTML content
      const simpleHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(simpleHtml);
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response)
          }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages across notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      
      // Get all pages
      const response = await graphClient.api(`/me/onenote/pages`).get();
      
      // If search string is provided, filter the results
      if (params.random_string && params.random_string.length > 0) {
        const searchTerm = params.random_string.toLowerCase();
        const filteredPages = response.value.filter(page => {
          // Search in title
          if (page.title && page.title.toLowerCase().includes(searchTerm)) {
            return true;
          }
          return false;
        });
        
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(filteredPages)
            }
          ]
        };
      } else {
        // Return all pages if no search term
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(response.value)
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// ---------------------------------------------------------------------------
// updatePage: replace, append to, or create a OneNote page.
// Used by the "Master List ABC" artifact (host:onenote -> updatePage).
// ---------------------------------------------------------------------------

// Read the freshest saved token on every call, so re-running `npm run auth`
// takes effect without restarting the server.
function readFreshToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const raw = fs.readFileSync(tokenFilePath, 'utf8');
      try {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.token) return parsed.token;
      } catch (e) {
        if (raw.trim()) return raw.trim();
      }
    }
  } catch (e) {
    console.error('Error reading token file:', e.message);
  }
  return accessToken || process.env.GRAPH_ACCESS_TOKEN || null;
}

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function graphFetch(pathOrUrl, { method = 'GET', headers = {}, body } = {}) {
  const token = readFreshToken();
  if (!token) {
    throw new Error('No access token found. Run `npm run auth` in the onenote-mcp folder, then try again.');
  }
  const url = pathOrUrl.startsWith('http') ? pathOrUrl : GRAPH + pathOrUrl;
  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...headers },
    body
  });
  if (res.status === 401) {
    throw new Error('Microsoft rejected the saved sign-in (401, usually an expired token). Run `npm run auth` in the onenote-mcp folder, then try again.');
  }
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${pathOrUrl.split('?')[0]} failed (${res.status}): ${text}`.slice(0, 900));
  }
  return res;
}

async function graphListAll(path, maxPages = 10) {
  let url = path;
  const out = [];
  for (let i = 0; i < maxPages && url; i++) {
    const res = await graphFetch(url);
    const json = await res.json();
    out.push(...(json.value || []));
    url = json['@odata.nextLink'] || null;
  }
  return out;
}

function xmlEscape(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

async function findSection(notebook, section) {
  const sections = await graphListAll('/me/onenote/sections?$expand=parentNotebook&$top=100');
  const wantSection = section.trim().toLowerCase();
  const wantNotebook = notebook ? notebook.trim().toLowerCase() : null;
  const matches = sections.filter(s =>
    (s.displayName || '').trim().toLowerCase() === wantSection &&
    (!wantNotebook || ((s.parentNotebook && s.parentNotebook.displayName) || '').trim().toLowerCase() === wantNotebook)
  );
  if (matches.length === 0) {
    throw new Error(`No section named "${section}"${notebook ? ` in notebook "${notebook}"` : ''} was found.`);
  }
  if (matches.length > 1) {
    throw new Error(`More than one section named "${section}" matched. Pass a notebook name to pick one.`);
  }
  return matches[0];
}

async function findPageByTitle(sectionId, title) {
  const pages = await graphListAll(`/me/onenote/sections/${sectionId}/pages?$select=id,title,links&$top=100`);
  const want = title.trim().toLowerCase();
  const matches = pages.filter(p => (p.title || '').trim().toLowerCase() === want);
  if (matches.length === 0) throw new Error(`No page titled "${title}" was found in that section.`);
  if (matches.length > 1) throw new Error(`${matches.length} pages are titled "${title}" in that section. Rename or remove the duplicates first.`);
  return matches[0];
}

async function createPageInSection(sectionId, title, bodyHtml) {
  const doc = `<!DOCTYPE html><html><head><title>${xmlEscape(title)}</title></head><body>${bodyHtml}</body></html>`;
  const res = await graphFetch(`/me/onenote/sections/${sectionId}/pages`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/html' },
    body: doc
  });
  return res.json();
}

server.tool(
  "updatePage",
  "Replace, append to, or create a OneNote page. mode 'replace' swaps the text on an existing page (after saving a backup copy of it), 'append' adds to the end of it, 'new' creates a separate dated page. html is a body fragment such as <p>..</p><ul><li>..</li></ul>.",
  {
    notebook: z.string().optional().describe("Notebook name, e.g. My 2024 Notebook"),
    section: z.string().describe("Section name, e.g. Personal"),
    title: z.string().describe("Title of the existing page, e.g. Master List"),
    html: z.string().min(1).max(200000).describe("Body HTML fragment to write"),
    mode: z.enum(["replace", "append", "new"]).default("replace"),
    backup: z.boolean().default(true).describe("In replace mode, first save a text copy of the old page as a new page in the same section")
  },
  async ({ notebook, section, title, html, mode, backup }) => {
    try {
      const sec = await findSection(notebook, section);
      const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ');

      if (mode === 'new') {
        const created = await createPageInSection(sec.id, `${title} (prioritized ${stamp.slice(0, 10)})`, html);
        return { content: [{ type: "text", text: JSON.stringify({
          ok: true, mode, pageId: created.id, pageTitle: created.title,
          pageUrl: created.links && created.links.oneNoteWebUrl && created.links.oneNoteWebUrl.href
        }) }] };
      }

      const page = await findPageByTitle(sec.id, title);
      const pageUrl = page.links && page.links.oneNoteWebUrl && page.links.oneNoteWebUrl.href;

      if (mode === 'append') {
        await graphFetch(`/me/onenote/pages/${page.id}/content`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify([{ target: 'body', action: 'append', content: `<div>${html}</div>` }])
        });
        return { content: [{ type: "text", text: JSON.stringify({ ok: true, mode, pageId: page.id, pageUrl }) }] };
      }

      // mode === 'replace'
      // 1. Backup: text copy of the current page (images and ink are not copied).
      let backupTitle = null;
      if (backup) {
        const plain = await (await graphFetch(`/me/onenote/pages/${page.id}/content`)).text();
        const oldBody = new JSDOM(plain).window.document.body.innerHTML;
        backupTitle = `${title} (backup ${stamp})`;
        await createPageInSection(sec.id, backupTitle, oldBody);
      }

      // 2. Find the page's outlines (top-level containers) via generated ids.
      const withIds = await (await graphFetch(`/me/onenote/pages/${page.id}/content?includeIDs=true`)).text();
      const doc = new JSDOM(withIds).window.document;
      const outlines = Array.from(doc.body.children).filter(el => el.tagName === 'DIV' && el.id);

      // 3. Replace the first outline with the new content and empty the rest.
      const commands = [];
      if (outlines.length === 0) {
        commands.push({ target: 'body', action: 'append', content: `<div>${html}</div>` });
      } else {
        outlines.forEach((el, i) => {
          commands.push({ target: el.id, action: 'replace', content: i === 0 ? `<div>${html}</div>` : '<div></div>' });
        });
      }
      await graphFetch(`/me/onenote/pages/${page.id}/content`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(commands)
      });

      return { content: [{ type: "text", text: JSON.stringify({
        ok: true, mode, pageId: page.id, pageUrl, backupTitle
      }) }] };
    } catch (error) {
      console.error("Error updating page:", error);
      throw new Error(`Failed to update page: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    // Connect to standard I/O
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('Server started successfully.');
    console.error('Use the "authenticate" tool to start the authentication flow,');
    console.error('or use "saveAccessToken" if you already have a token.');
    
    // Keep the process alive
    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main(); 