/**
 * Microsoft Graph client with OBO token exchange
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { config } from '../config.js';

let msalClient: ConfidentialClientApplication | null = null;

/**
 * Initialize MSAL confidential client for OBO flow
 */
function getMSALClient(): ConfidentialClientApplication {
  if (msalClient) return msalClient;

  msalClient = new ConfidentialClientApplication({
    auth: {
      clientId: config.azure.clientId,
      authority: `https://login.microsoftonline.com/${config.azure.tenantId}`,
      clientSecret: process.env.AZURE_CLIENT_SECRET || '',
    },
  });

  return msalClient;
}

/**
 * Exchange user token for Graph token using OBO flow
 */
export async function getOBOToken(userToken: string): Promise<string> {
  const client = getMSALClient();

  const oboRequest = {
    oboAssertion: userToken,
    scopes: ['https://graph.microsoft.com/.default'],
  };

  try {
    const response = await client.acquireTokenOnBehalfOf(oboRequest);
    
    if (!response || !response.accessToken) {
      throw new Error('Failed to acquire OBO token');
    }

    return response.accessToken;
  } catch (err) {
    console.error('[OBO] Token exchange failed:', err);
    throw new Error('Failed to exchange token for Graph access');
  }
}

/**
 * Create Microsoft Graph client with access token
 */
export function createGraphClient(accessToken: string): Client {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
    defaultVersion: 'v1.0',
  });
}

/**
 * Retry middleware for Graph client (handles 429 and transient errors)
 */
export async function graphRetry<T>(
  fn: () => Promise<T>,
  maxRetries: number = 3,
  baseDelayMs: number = 1000
): Promise<T> {
  let lastError: any;

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await fn();
    } catch (err: any) {
      lastError = err;

      // Don't retry on 4xx errors except 429
      if (err.statusCode && err.statusCode >= 400 && err.statusCode < 500 && err.statusCode !== 429) {
        throw err;
      }

      // Check for Retry-After header
      const retryAfter = err.headers?.['retry-after'];
      let delay = baseDelayMs * Math.pow(2, attempt);

      if (retryAfter) {
        delay = parseInt(retryAfter, 10) * 1000;
      }

      // Add jitter
      delay += Math.random() * 1000;

      console.log(`[GraphRetry] Attempt ${attempt + 1}/${maxRetries} failed, retrying in ${delay}ms`, {
        statusCode: err.statusCode,
        message: err.message,
      });

      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }

  throw lastError;
}

