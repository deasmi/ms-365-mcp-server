import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { clearSecretsCache } from '../src/secrets.js';

interface CapturedMsalConfig {
  auth: { clientId: string; authority: string };
  broker?: { nativeBrokerPlugin: unknown };
}

const hoisted = vi.hoisted(() => ({
  state: { lastMsalConfig: null as CapturedMsalConfig | null },
  acquireTokenInteractiveMock: vi.fn(),
  acquireTokenByDeviceCodeMock: vi.fn(),
  acquireTokenSilentMock: vi.fn(),
  NativeBrokerPluginCtor: vi.fn(),
  loggerInfo: vi.fn(),
  loggerWarn: vi.fn(),
  loggerError: vi.fn(),
}));

const {
  acquireTokenInteractiveMock,
  acquireTokenByDeviceCodeMock,
  NativeBrokerPluginCtor,
  loggerWarn,
} = hoisted;

vi.mock('@azure/msal-node', () => ({
  PublicClientApplication: vi.fn().mockImplementation((config: CapturedMsalConfig) => {
    hoisted.state.lastMsalConfig = config;
    return {
      acquireTokenInteractive: hoisted.acquireTokenInteractiveMock,
      acquireTokenByDeviceCode: hoisted.acquireTokenByDeviceCodeMock,
      acquireTokenSilent: hoisted.acquireTokenSilentMock,
      getTokenCache: () => ({
        serialize: () => 'serialized',
        deserialize: () => {},
        getAllAccounts: async () => [],
        removeAccount: async () => {},
      }),
    };
  }),
}));

vi.mock('@azure/msal-node-extensions', () => ({
  NativeBrokerPlugin: hoisted.NativeBrokerPluginCtor,
}));

vi.mock('../src/logger.js', () => ({
  default: { info: hoisted.loggerInfo, warn: hoisted.loggerWarn, error: hoisted.loggerError },
}));

vi.mock('open', () => ({
  default: vi.fn().mockResolvedValue(undefined),
}));

const originalPlatform = process.platform;

function setPlatform(value: string): void {
  Object.defineProperty(process, 'platform', { value, configurable: true });
}

async function loadAuthManager() {
  // Importing dynamically per-test guarantees we re-read process.platform-affected
  // state and start with a fresh module-level cache where applicable.
  const mod = await import('../src/auth.js');
  return mod.default;
}

describe('AuthManager WAM broker support', () => {
  beforeEach(() => {
    hoisted.state.lastMsalConfig = null;
    vi.clearAllMocks();
    clearSecretsCache();
  });

  afterEach(() => {
    setPlatform(originalPlatform);
    clearSecretsCache();
  });

  it('attaches the NativeBrokerPlugin to the MSAL config on Windows when useBroker is true', async () => {
    setPlatform('win32');

    const AuthManager = await loadAuthManager();
    await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      useBroker: true,
    });

    expect(NativeBrokerPluginCtor).toHaveBeenCalledTimes(1);
    expect(hoisted.state.lastMsalConfig?.broker).toBeDefined();
    expect(hoisted.state.lastMsalConfig?.broker?.nativeBrokerPlugin).toBeInstanceOf(
      NativeBrokerPluginCtor
    );
  });

  it('logs a warning and skips broker setup on non-Windows platforms', async () => {
    setPlatform('linux');

    const AuthManager = await loadAuthManager();
    await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      useBroker: true,
    });

    expect(NativeBrokerPluginCtor).not.toHaveBeenCalled();
    expect(hoisted.state.lastMsalConfig?.broker).toBeUndefined();
    expect(loggerWarn).toHaveBeenCalledWith(
      expect.stringContaining(
        '--use-broker requested but the WAM broker is only supported on Windows'
      )
    );
  });

  it('does not configure the broker when useBroker is false (default)', async () => {
    setPlatform('win32');

    const AuthManager = await loadAuthManager();
    await AuthManager.create(['User.Read'], undefined, { storage: makeStorage() });

    expect(NativeBrokerPluginCtor).not.toHaveBeenCalled();
    expect(hoisted.state.lastMsalConfig?.broker).toBeUndefined();
  });

  it('routes device code requests through the brokered interactive flow on Windows', async () => {
    setPlatform('win32');

    acquireTokenInteractiveMock.mockResolvedValue({
      accessToken: 'broker-token',
      expiresOn: new Date(Date.now() + 60_000),
      scopes: ['User.Read'],
      account: { homeAccountId: 'home-id', username: 'u@example.com' },
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
      useBroker: true,
    });

    const token = await auth.acquireTokenByDeviceCode();

    expect(token).toBe('broker-token');
    expect(acquireTokenInteractiveMock).toHaveBeenCalledTimes(1);
    expect(acquireTokenByDeviceCodeMock).not.toHaveBeenCalled();

    const interactiveCallArgs = acquireTokenInteractiveMock.mock.calls[0][0];
    expect(interactiveCallArgs.windowHandle).toBeInstanceOf(Buffer);
  });

  it('keeps the original device code path when broker is disabled', async () => {
    setPlatform('win32');

    acquireTokenByDeviceCodeMock.mockResolvedValue({
      accessToken: 'device-token',
      expiresOn: new Date(Date.now() + 60_000),
      scopes: ['User.Read'],
      account: { homeAccountId: 'home-id', username: 'u@example.com' },
    });

    const AuthManager = await loadAuthManager();
    const auth = await AuthManager.create(['User.Read'], undefined, {
      storage: makeStorage(),
    });

    const token = await auth.acquireTokenByDeviceCode();

    expect(token).toBe('device-token');
    expect(acquireTokenByDeviceCodeMock).toHaveBeenCalledTimes(1);
    expect(acquireTokenInteractiveMock).not.toHaveBeenCalled();
  });
});

function makeStorage() {
  return {
    description: 'mock',
    failClosed: false as const,
    load: vi.fn().mockResolvedValue(undefined),
    save: vi.fn().mockResolvedValue(undefined),
    delete: vi.fn().mockResolvedValue(undefined),
  };
}
