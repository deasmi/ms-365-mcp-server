import { spawn } from 'child_process';
import logger from './logger.js';

export interface ExternalBrokerToken {
  accessToken: string;
  expiresOn: Date | null;
}

export interface ExternalBrokerContext {
  clientId: string;
  tenantId: string;
  scopes: string[];
  resource: string;
}

/**
 * Invokes an external command that returns an access token as JSON on stdout.
 *
 * Designed for the WSL2 → Windows-WAM bridge: from a Linux ms-365-mcp-server
 * instance, shell out to e.g. azureauth.exe via WSL interop, let WAM use the
 * device's Primary Refresh Token (which carries the device-compliance claim),
 * and consume the resulting bearer. The same hook works for any external
 * token source — cloud STS, corp tooling, etc.
 *
 * The command template supports {clientId}, {tenantId}, {scopes}, {resource}
 * placeholders. The same values are also exported as MS365_MCP_BROKER_*
 * environment variables for commands that prefer env over args.
 *
 * Expected stdout JSON (azureauth's shape is supported natively):
 *   { "access_token": "...", "expires_on": <epoch seconds> }
 *   { "token":        "...", "expirationTimestamp": <epoch seconds> }
 */
export async function acquireTokenFromExternalBroker(
  commandTemplate: string,
  ctx: ExternalBrokerContext
): Promise<ExternalBrokerToken> {
  const expanded = commandTemplate
    .replace(/\{clientId\}/g, ctx.clientId)
    .replace(/\{tenantId\}/g, ctx.tenantId)
    .replace(/\{scopes\}/g, ctx.scopes.join(' '))
    .replace(/\{resource\}/g, ctx.resource);

  logger.info(`Invoking external broker command for ${ctx.scopes.length} scope(s)`);

  return new Promise((resolve, reject) => {
    const child = spawn(expanded, {
      shell: true,
      env: {
        ...process.env,
        MS365_MCP_BROKER_CLIENT_ID: ctx.clientId,
        MS365_MCP_BROKER_TENANT_ID: ctx.tenantId,
        MS365_MCP_BROKER_SCOPES: ctx.scopes.join(' '),
        MS365_MCP_BROKER_RESOURCE: ctx.resource,
      },
    });

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk: Buffer) => {
      stdout += chunk.toString('utf8');
    });
    child.stderr.on('data', (chunk: Buffer) => {
      stderr += chunk.toString('utf8');
    });
    child.on('error', (err) => {
      reject(new Error(`Failed to spawn broker command: ${err.message}`));
    });
    child.on('close', (code) => {
      if (code !== 0) {
        const detail = stderr.trim() || stdout.trim() || '(no output)';
        reject(new Error(`Broker command exited with code ${code}: ${detail}`));
        return;
      }
      try {
        const parsed = parseBrokerOutput(stdout);
        resolve(parsed);
      } catch (err) {
        reject(new Error(`Failed to parse broker output: ${(err as Error).message}`));
      }
    });
  });
}

function parseBrokerOutput(raw: string): ExternalBrokerToken {
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error('broker produced empty stdout');
  }
  const parsed = JSON.parse(trimmed) as Record<string, unknown>;

  const token = parsed.access_token ?? parsed.token;
  if (typeof token !== 'string' || token.length === 0) {
    throw new Error('broker JSON missing access_token / token');
  }

  return {
    accessToken: token,
    expiresOn: parseBrokerExpiry(parsed.expires_on ?? parsed.expirationTimestamp),
  };
}

function parseBrokerExpiry(raw: unknown): Date | null {
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    // Heuristic: 13-digit values are epoch milliseconds, 10-digit are seconds.
    const ms = raw > 1e12 ? raw : raw * 1000;
    return new Date(ms);
  }
  if (typeof raw === 'string' && raw.length > 0) {
    const asNumber = Number(raw);
    if (Number.isFinite(asNumber)) {
      return parseBrokerExpiry(asNumber);
    }
    const parsedDate = new Date(raw);
    if (!Number.isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  }
  return null;
}
