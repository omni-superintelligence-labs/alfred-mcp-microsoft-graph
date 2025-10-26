/**
 * JWT validation for AAD tokens and Alfred add-in tokens
 */

import { createRemoteJWKSet, jwtVerify, SignJWT, importPKCS8 } from 'jose';
import { config } from '../config.js';
import crypto from 'crypto';

const JWKS_URI = `https://login.microsoftonline.com/${config.azure.tenantId}/discovery/v2.0/keys`;
const jwks = createRemoteJWKSet(new URL(JWKS_URI));

interface JWTPayload {
  sub: string;
  aud: string;
  iss: string;
  exp: number;
  iat: number;
  tid: string;
  oid: string;
  preferred_username?: string;
  name?: string;
}

/**
 * Validate JWT token from AAD
 */
export async function validateJWT(token: string): Promise<JWTPayload> {
  try {
    const { payload } = await jwtVerify(token, jwks, {
      issuer: `https://login.microsoftonline.com/${config.azure.tenantId}/v2.0`,
      audience: config.azure.clientId,
    });

    return payload as unknown as JWTPayload;
  } catch (err) {
    console.error('[JWT] Validation failed:', err);
    throw new Error('Invalid JWT token');
  }
}

/**
 * Decode JWT without validation (for debugging)
 */
export function decodeJWT(token: string): any {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('Invalid JWT format');
  }

  const payload = Buffer.from(parts[1], 'base64').toString('utf-8');
  return JSON.parse(payload);
}

// ==========================================
// Alfred Add-in Token Management
// ==========================================

/**
 * Generate a signing key for Alfred tokens (in-memory for now; use KMS in production)
 */
const ALFRED_JWT_SECRET = process.env.ALFRED_JWT_SECRET || crypto.randomBytes(32).toString('hex');
const ALFRED_ISSUER = 'alfred-api';
const ALFRED_ADDIN_AUDIENCE = 'alfred-addin';

interface AlfredTokenPayload {
  sub: string; // user ID
  aud: string; // 'alfred-addin'
  iss: string; // 'alfred-api'
  exp: number;
  iat: number;
  nonce: string;
  tabId: string;
}

/**
 * Issue a short-lived JWT for the add-in (10 minutes)
 */
export async function issueAddinToken(
  userId: string,
  tabId: string,
  nonce: string
): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const exp = now + 600; // 10 minutes

  // Use HS256 for simplicity (symmetric key)
  const secret = new TextEncoder().encode(ALFRED_JWT_SECRET);

  const token = await new SignJWT({
    sub: userId,
    nonce,
    tabId,
  })
    .setProtectedHeader({ alg: 'HS256' })
    .setIssuedAt(now)
    .setIssuer(ALFRED_ISSUER)
    .setAudience(ALFRED_ADDIN_AUDIENCE)
    .setExpirationTime(exp)
    .sign(secret);

  return token;
}

/**
 * Validate an Alfred add-in token
 */
export async function validateAddinToken(token: string): Promise<AlfredTokenPayload> {
  try {
    const secret = new TextEncoder().encode(ALFRED_JWT_SECRET);

    const { payload } = await jwtVerify(token, secret, {
      issuer: ALFRED_ISSUER,
      audience: ALFRED_ADDIN_AUDIENCE,
    });

    return payload as unknown as AlfredTokenPayload;
  } catch (err) {
    console.error('[JWT] Add-in token validation failed:', err);
    throw new Error('Invalid add-in token');
  }
}

