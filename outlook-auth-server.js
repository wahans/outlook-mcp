#!/usr/bin/env node
const http = require('http');
const url = require('url');
const querystring = require('querystring');
const https = require('https');
const fs = require('fs');
const path = require('path');

// Load environment variables from .env file
require('dotenv').config();

// Log to console
console.log('Starting Outlook Authentication Server');

// Authentication configuration
const AUTH_CONFIG = {
  clientId: process.env.MS_CLIENT_ID || '', // Set your client ID as an environment variable
  clientSecret: process.env.MS_CLIENT_SECRET || '', // Set your client secret as an environment variable
  redirectUri: 'http://localhost:3333/auth/callback',
  scopes: [
    'offline_access',
    'User.Read',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.Read',
    'Calendars.ReadWrite',
    'Contacts.Read'
  ],
  tokenStorePath: path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json')
};

// Power Automate / Flow authentication configuration
const FLOW_AUTH_CONFIG = {
  clientId: process.env.MS_CLIENT_ID || '',
  clientSecret: process.env.MS_CLIENT_SECRET || '',
  redirectUri: 'http://localhost:3333/auth/flow/callback',
  scopes: [
    'https://service.flow.microsoft.com//.default',
    'offline_access'
  ],
  tokenStorePath: path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json')
};

// Create HTTP server
const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  
  console.log(`Request received: ${pathname}`);
  
  if (pathname === '/auth/callback') {
    const query = parsedUrl.query;
    
    if (query.error) {
      console.error(`Authentication error: ${query.error} - ${query.error_description}`);
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head>
            <title>Authentication Error</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
              h1 { color: #d9534f; }
              .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
            </style>
          </head>
          <body>
            <h1>Authentication Error</h1>
            <div class="error-box">
              <p><strong>Error:</strong> ${query.error}</p>
              <p><strong>Description:</strong> ${query.error_description || 'No description provided'}</p>
            </div>
            <p>Please close this window and try again.</p>
          </body>
        </html>
      `);
      return;
    }
    
    if (query.code) {
      console.log('Authorization code received, exchanging for tokens...');
      
      // Exchange code for tokens
      exchangeCodeForTokens(query.code)
        .then((tokens) => {
          console.log('Token exchange successful');
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <head>
                <title>Authentication Successful</title>
                <style>
                  body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
                  h1 { color: #5cb85c; }
                  .success-box { background-color: #d4edda; border: 1px solid #c3e6cb; padding: 15px; border-radius: 4px; }
                </style>
              </head>
              <body>
                <h1>Authentication Successful!</h1>
                <div class="success-box">
                  <p>You have successfully authenticated with Microsoft Graph API.</p>
                  <p>The access token has been saved securely.</p>
                </div>
                <p>You can now close this window and return to Claude.</p>
              </body>
            </html>
          `);
        })
        .catch((error) => {
          console.error(`Token exchange error: ${error.message}`);
          res.writeHead(500, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <head>
                <title>Token Exchange Error</title>
                <style>
                  body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
                  h1 { color: #d9534f; }
                  .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
                </style>
              </head>
              <body>
                <h1>Token Exchange Error</h1>
                <div class="error-box">
                  <p>${error.message}</p>
                </div>
                <p>Please close this window and try again.</p>
              </body>
            </html>
          `);
        });
    } else {
      console.error('No authorization code provided');
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head>
            <title>Missing Authorization Code</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
              h1 { color: #d9534f; }
              .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
            </style>
          </head>
          <body>
            <h1>Missing Authorization Code</h1>
            <div class="error-box">
              <p>No authorization code was provided in the callback.</p>
            </div>
            <p>Please close this window and try again.</p>
          </body>
        </html>
      `);
    }
  } else if (pathname === '/auth') {
    // Handle the /auth route - redirect to Microsoft's OAuth authorization endpoint
    console.log('Auth request received, redirecting to Microsoft login...');
    
    // Verify credentials are set
    if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
      res.writeHead(500, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head>
            <title>Configuration Error</title>
            <style>
              body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
              h1 { color: #d9534f; }
              .error-box { background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px; }
              code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
            </style>
          </head>
          <body>
            <h1>Configuration Error</h1>
            <div class="error-box">
              <p>Microsoft Graph API credentials are not set. Please set the following environment variables:</p>
              <ul>
                <li><code>MS_CLIENT_ID</code></li>
                <li><code>MS_CLIENT_SECRET</code></li>
              </ul>
            </div>
          </body>
        </html>
      `);
      return;
    }
    
    // Get client_id from query parameters or use the default
    const query = parsedUrl.query;
    const clientId = query.client_id || AUTH_CONFIG.clientId;
    
    // Build the authorization URL
    const authParams = {
      client_id: clientId,
      response_type: 'code',
      redirect_uri: AUTH_CONFIG.redirectUri,
      scope: AUTH_CONFIG.scopes.join(' '),
      response_mode: 'query',
      state: Date.now().toString() // Simple state parameter for security
    };
    
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${querystring.stringify(authParams)}`;
    console.log(`Redirecting to: ${authUrl}`);
    
    // Redirect to Microsoft's login page
    res.writeHead(302, { 'Location': authUrl });
    res.end();
  } else if (pathname === '/auth/flow') {
    // Handle the /auth/flow route - redirect to Microsoft's OAuth authorization endpoint for Power Automate
    console.log('Flow auth request received, redirecting to Microsoft login...');

    // Verify credentials are set
    if (!FLOW_AUTH_CONFIG.clientId || !FLOW_AUTH_CONFIG.clientSecret) {
      res.writeHead(500, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head><title>Configuration Error</title></head>
          <body>
            <h1>Configuration Error</h1>
            <p>Microsoft credentials are not set. Please set MS_CLIENT_ID and MS_CLIENT_SECRET.</p>
          </body>
        </html>
      `);
      return;
    }

    // Build the authorization URL for Flow
    const flowAuthParams = {
      client_id: FLOW_AUTH_CONFIG.clientId,
      response_type: 'code',
      redirect_uri: FLOW_AUTH_CONFIG.redirectUri,
      scope: FLOW_AUTH_CONFIG.scopes.join(' '),
      response_mode: 'query',
      state: 'flow_' + Date.now().toString()
    };

    const flowAuthUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${querystring.stringify(flowAuthParams)}`;
    console.log(`Redirecting to Flow auth: ${flowAuthUrl}`);

    res.writeHead(302, { 'Location': flowAuthUrl });
    res.end();
  } else if (pathname === '/auth/flow/callback') {
    // Handle Flow OAuth callback
    const query = parsedUrl.query;

    if (query.error) {
      console.error(`Flow auth error: ${query.error} - ${query.error_description}`);
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(`
        <html>
          <head><title>Flow Authentication Error</title></head>
          <body>
            <h1 style="color: #d9534f;">Flow Authentication Error</h1>
            <p><strong>Error:</strong> ${query.error}</p>
            <p><strong>Description:</strong> ${query.error_description || 'No description'}</p>
            <p>Please close this window and try again.</p>
          </body>
        </html>
      `);
      return;
    }

    if (query.code) {
      console.log('Flow authorization code received, exchanging for tokens...');

      exchangeCodeForFlowTokens(query.code)
        .then((tokens) => {
          console.log('Flow token exchange successful');
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <head><title>Power Automate Authentication Successful</title></head>
              <body>
                <h1 style="color: #5cb85c;">Power Automate Authentication Successful!</h1>
                <p>You have successfully authenticated with Power Automate.</p>
                <p>You can now close this window and use Flow commands in Claude.</p>
              </body>
            </html>
          `);
        })
        .catch((error) => {
          console.error(`Flow token exchange error: ${error.message}`);
          res.writeHead(500, { 'Content-Type': 'text/html' });
          res.end(`
            <html>
              <head><title>Flow Token Exchange Error</title></head>
              <body>
                <h1 style="color: #d9534f;">Flow Token Exchange Error</h1>
                <p>${error.message}</p>
                <p>Please close this window and try again.</p>
              </body>
            </html>
          `);
        });
    } else {
      res.writeHead(400, { 'Content-Type': 'text/plain' });
      res.end('Missing authorization code');
    }
  } else if (pathname === '/') {
    // Root path - provide instructions
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`
      <html>
        <head>
          <title>Outlook Authentication Server</title>
          <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            h1 { color: #0078d4; }
            .info-box { background-color: #e7f6fd; border: 1px solid #b3e0ff; padding: 15px; border-radius: 4px; }
            code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
          </style>
        </head>
        <body>
          <h1>Outlook Authentication Server</h1>
          <div class="info-box">
            <p>This server is running to handle Microsoft Graph API authentication callbacks.</p>
            <p>Don't navigate here directly. Instead, use the <code>authenticate</code> tool in Claude to start the authentication process.</p>
            <p>Make sure you've set the <code>MS_CLIENT_ID</code> and <code>MS_CLIENT_SECRET</code> environment variables.</p>
          </div>
          <p>Server is running at http://localhost:3333</p>
        </body>
      </html>
    `);
  } else {
    // Not found
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found');
  }
});

function exchangeCodeForTokens(code) {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: AUTH_CONFIG.clientId,
      client_secret: AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: AUTH_CONFIG.scopes.join(' ')
    });
    
    const options = {
      hostname: 'login.microsoftonline.com',
      path: '/common/oauth2/v2.0/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };
    
    const req = https.request(options, (res) => {
      let data = '';
      
      res.on('data', (chunk) => {
        data += chunk;
      });
      
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);
            
            // Calculate expiration time (current time + expires_in seconds)
            const expiresAt = Date.now() + (tokenResponse.expires_in * 1000);
            
            // Add expires_at for easier expiration checking
            tokenResponse.expires_at = expiresAt;
            
            // Save tokens to file
            fs.writeFileSync(AUTH_CONFIG.tokenStorePath, JSON.stringify(tokenResponse, null, 2), 'utf8');
            console.log(`Tokens saved to ${AUTH_CONFIG.tokenStorePath}`);
            
            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });
    
    req.on('error', (error) => {
      reject(error);
    });
    
    req.write(postData);
    req.end();
  });
}

function exchangeCodeForFlowTokens(code) {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: FLOW_AUTH_CONFIG.clientId,
      client_secret: FLOW_AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: FLOW_AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: FLOW_AUTH_CONFIG.scopes.join(' ')
    });

    const options = {
      hostname: 'login.microsoftonline.com',
      path: '/common/oauth2/v2.0/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);

            // Calculate expiration time
            const expiresAt = Date.now() + (tokenResponse.expires_in * 1000);

            // Load existing tokens and merge Flow tokens
            let existingTokens = {};
            try {
              if (fs.existsSync(FLOW_AUTH_CONFIG.tokenStorePath)) {
                existingTokens = JSON.parse(fs.readFileSync(FLOW_AUTH_CONFIG.tokenStorePath, 'utf8'));
              }
            } catch (e) {
              console.log('No existing tokens found, creating new file');
            }

            // Merge Flow tokens with existing Graph tokens
            const mergedTokens = {
              ...existingTokens,
              flow_access_token: tokenResponse.access_token,
              flow_refresh_token: tokenResponse.refresh_token,
              flow_expires_at: expiresAt,
              flow_scope: tokenResponse.scope
            };

            fs.writeFileSync(FLOW_AUTH_CONFIG.tokenStorePath, JSON.stringify(mergedTokens, null, 2), 'utf8');
            console.log(`Flow tokens saved to ${FLOW_AUTH_CONFIG.tokenStorePath}`);

            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing Flow token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Flow token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });

    req.on('error', (error) => {
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}

// Start server
const PORT = 3333;
server.listen(PORT, () => {
  console.log(`Authentication server running at http://localhost:${PORT}`);
  console.log(`Waiting for authentication callback at ${AUTH_CONFIG.redirectUri}`);
  console.log(`Token will be stored at: ${AUTH_CONFIG.tokenStorePath}`);
  
  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log('\n⚠️  WARNING: Microsoft Graph API credentials are not set.');
    console.log('   Please set the MS_CLIENT_ID and MS_CLIENT_SECRET environment variables.');
  }
});

// Handle termination
process.on('SIGINT', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});
