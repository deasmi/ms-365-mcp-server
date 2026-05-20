import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { clearSecretsCache } from '../src/secrets.js';

const hoisted = vi.hoisted(() => ({
  acquireTokenInteractiveMock: vi.fn(),
  acquireTokenByDeviceCodeMock: vi.fn(),
  acquireTokenSilentMock: vi.fn(),
  acquireFromBrokerMock: vi.fn(),
  loggerInfo: vi.fn(),
  loggerWarn: vi.fn(),
  loggerError: vi.fn(),
}));

vi.mock('@azure/msal-node', () => ({
  PublicClientApplication: vi.fn().mockImplementation(() => ({
    acquireTokenInteractive: hoisted.acquireTokenInteractiveMock,
    acquireTokenByDeviceCode: hoisted.acquireTokenByDeviceCodeMock,
    acquireTokenSilent: hoisted.acquireTokenSilentMock,
    getTokenCache: () => ({
      serialize: () => 'serialized',
      deserialize: () => {},
      getAllAccounts: async () => [],
      removeAccount: async () => {},
    }),
  })),
}));

vi.mock('../src/external-broker.js', () => ({
  acquireTokenFromExternalBroker: hoisted.acquireFromBrokerMock,
}));

vi.mock('../src/logger.js', () => ({
  default: { info: hoisted.loggerInfo, warn: hoisted.loggerWarn, error: hoisted.loggerError },
}));

vi.mock('open', () => ({ default: vi.fn().mockResolvedValue(undefined) }));

async function loadAuthManager() {
  const mod = await import('../src/auth.js');
  return mod.default;
}

function makeStorage() {
  return {
    description: 'mock',
    failClosed: false as const,
    load: vi.fn().mockResolvedValue(undefined),
    save: vi.fn().mockResolvedValue(undefined),
    delete: vi.fn().mockResolvedValue(undefined),
  };
}

describe('AuthManager with --broker-cmd', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    clearSecretsCache();
  });

  afterEach(() => {
    clearSecretsCache();
  });

  it('routes acquireTokenByDeviceCode to the external broker, skipping MSAL device code', async () => {
    hoisted.acquireFromBrokerMock.mockResolvedValue({
      accessToken: 'broker-bearer',
      expiresOn: new Date(Date.now() + 60_000),
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      brokerCmd: 'azureauth.exe aad --client {clientId} --output json',
    });

    const token = await auth.acquireTokenByDeviceCode();

    expect(token).toBe('broker-bearer');
    expect(hoisted.acquireFromBrokerMock).toHaveBeenCalledTimes(1);
    expect(hoisted.acquireTokenByDeviceCodeMock).not.toHaveBeenCalled();

    const call = hoisted.acquireFromBrokerMock.mock.calls[0];
    expect(call[0]).toBe('azureauth.exe aad --client {clientId} --output json');
    expect(call[1]).toMatchObject({
      scopes: ['User.Read'],
      resource: expect.stringContaining('graph.microsoft.com'),
    });
  });

  it('routes acquireTokenInteractive to the external broker too', async () => {
    hoisted.acquireFromBrokerMock.mockResolvedValue({
      accessToken: 'broker-bearer-2',
      expiresOn: null,
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      brokerCmd: 'my-broker',
    });

    const token = await auth.acquireTokenInteractive();

    expect(token).toBe('broker-bearer-2');
    expect(hoisted.acquireFromBrokerMock).toHaveBeenCalledTimes(1);
    expect(hoisted.acquireTokenInteractiveMock).not.toHaveBeenCalled();
  });

  it('getToken short-circuits to the broker when no MSAL account is cached', async () => {
    hoisted.acquireFromBrokerMock.mockResolvedValue({
      accessToken: 'fresh-broker-token',
      expiresOn: new Date(Date.now() + 60_000),
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      brokerCmd: 'my-broker',
    });

    const token = await auth.getToken();

    expect(token).toBe('fresh-broker-token');
    expect(hoisted.acquireFromBrokerMock).toHaveBeenCalledTimes(1);
    expect(hoisted.acquireTokenSilentMock).not.toHaveBeenCalled();
  });

  it('getToken returns the cached broker token until it expires', async () => {
    hoisted.acquireFromBrokerMock.mockResolvedValueOnce({
      accessToken: 'first-token',
      expiresOn: new Date(Date.now() + 60_000),
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      brokerCmd: 'my-broker',
    });

    expect(await auth.getToken()).toBe('first-token');
    expect(await auth.getToken()).toBe('first-token');
    expect(hoisted.acquireFromBrokerMock).toHaveBeenCalledTimes(1);
  });

  it('does not invoke the broker when brokerCmd is unset', async () => {
    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
    });

    await expect(auth.getToken()).rejects.toThrow(/No valid token found/);
    expect(hoisted.acquireFromBrokerMock).not.toHaveBeenCalled();
  });
});
