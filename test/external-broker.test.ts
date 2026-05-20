import { describe, expect, it } from 'vitest';
import { acquireTokenFromExternalBroker } from '../src/external-broker.js';

const baseCtx = {
  clientId: 'client-123',
  tenantId: 'tenant-abc',
  scopes: ['User.Read'],
  resource: 'https://graph.microsoft.com',
};

describe('acquireTokenFromExternalBroker', () => {
  it('parses {access_token, expires_on} JSON', async () => {
    const futureEpochSec = Math.floor(Date.now() / 1000) + 3600;
    const command = `printf '%s' '${JSON.stringify({
      access_token: 'eyJTOKEN',
      expires_on: futureEpochSec,
    })}'`;

    const result = await acquireTokenFromExternalBroker(command, baseCtx);

    expect(result.accessToken).toBe('eyJTOKEN');
    expect(result.expiresOn?.getTime()).toBe(futureEpochSec * 1000);
  });

  it('parses azureauth-shaped {token, expirationTimestamp} JSON', async () => {
    const futureEpochSec = Math.floor(Date.now() / 1000) + 7200;
    const command = `printf '%s' '${JSON.stringify({
      user: 'u@example.com',
      displayName: 'User',
      token: 'eyJAZUREAUTH',
      expirationTimestamp: futureEpochSec,
    })}'`;

    const result = await acquireTokenFromExternalBroker(command, baseCtx);

    expect(result.accessToken).toBe('eyJAZUREAUTH');
    expect(result.expiresOn?.getTime()).toBe(futureEpochSec * 1000);
  });

  it('expands {clientId}, {tenantId}, {scopes}, {resource} placeholders', async () => {
    const command =
      `printf '%s' '{"access_token":"placeholders-` +
      `client={clientId};tenant={tenantId};scopes={scopes};resource={resource}"}'`;

    const result = await acquireTokenFromExternalBroker(command, {
      clientId: 'C',
      tenantId: 'T',
      scopes: ['Mail.Read', 'Files.Read'],
      resource: 'https://graph.microsoft.com',
    });

    expect(result.accessToken).toBe(
      'placeholders-client=C;tenant=T;scopes=Mail.Read Files.Read;resource=https://graph.microsoft.com'
    );
  });

  it('exports MS365_MCP_BROKER_* env vars to the child process', async () => {
    const command =
      'printf \'{"access_token":"env-%s-%s-%s-%s"}\' ' +
      '"$MS365_MCP_BROKER_CLIENT_ID" "$MS365_MCP_BROKER_TENANT_ID" ' +
      '"$MS365_MCP_BROKER_SCOPES" "$MS365_MCP_BROKER_RESOURCE"';

    const result = await acquireTokenFromExternalBroker(command, {
      clientId: 'envC',
      tenantId: 'envT',
      scopes: ['a', 'b'],
      resource: 'https://example/',
    });

    expect(result.accessToken).toBe('env-envC-envT-a b-https://example/');
  });

  it('accepts string expiry that is a numeric epoch', async () => {
    const futureEpochSec = Math.floor(Date.now() / 1000) + 60;
    const command = `printf '%s' '${JSON.stringify({
      access_token: 't',
      expires_on: String(futureEpochSec),
    })}'`;

    const result = await acquireTokenFromExternalBroker(command, baseCtx);

    expect(result.expiresOn?.getTime()).toBe(futureEpochSec * 1000);
  });

  it('returns expiresOn=null when broker omits expiry', async () => {
    const command = `printf '%s' '${JSON.stringify({ access_token: 'no-exp' })}'`;

    const result = await acquireTokenFromExternalBroker(command, baseCtx);

    expect(result.accessToken).toBe('no-exp');
    expect(result.expiresOn).toBeNull();
  });

  it('rejects when the broker exits non-zero, surfacing stderr', async () => {
    const command = 'printf "broker boom" >&2; exit 7';

    await expect(acquireTokenFromExternalBroker(command, baseCtx)).rejects.toThrow(
      /exited with code 7.*broker boom/s
    );
  });

  it('rejects when stdout is not valid JSON', async () => {
    const command = "printf 'not json {';";

    await expect(acquireTokenFromExternalBroker(command, baseCtx)).rejects.toThrow(
      /Failed to parse broker output/
    );
  });

  it('rejects when JSON has no token field', async () => {
    const command = `printf '%s' '${JSON.stringify({ other: 'value' })}'`;

    await expect(acquireTokenFromExternalBroker(command, baseCtx)).rejects.toThrow(
      /missing access_token \/ token/
    );
  });

  it('rejects when broker produces empty stdout', async () => {
    const command = 'true';

    await expect(acquireTokenFromExternalBroker(command, baseCtx)).rejects.toThrow(/empty stdout/);
  });
});
