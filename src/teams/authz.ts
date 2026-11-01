import { CliError } from '../errors.ts';
import { createTeamsClient, type Endpoints } from './client.ts';

const AUTHZ_URL = 'https://teams.microsoft.com/api/authsvc/v1.0/authz';

export interface Authorization {
  skypeToken: string;
  expiresAt: Date;
  /** e.g. `amer`. */
  region: string;
  endpoints: Endpoints;
}

interface AuthzResponse {
  tokens?: { skypeToken?: string; expiresIn?: number };
  region?: string;
  regionGtms?: { chatService?: string; urlp?: string; urlpV2?: string };
}

/** Keeps an https URL-preview base from authz and drops a trailing slash. Anything else is unusable. */
function urlpBase(value: string | undefined): string | null {
  if (!value) return null;
  try {
    const url = new URL(value);
    if (url.protocol !== 'https:') return null;
    return `${url.origin}${url.pathname}`.replace(/\/+$/, '');
  } catch {
    return null;
  }
}

/** Exchanges a `spaces` access token for the skypetoken and the account's regional endpoints. */
export async function authorize(spacesToken: string, { fetch, now }: { fetch: typeof globalThis.fetch; now: () => Date }): Promise<Authorization> {
  const client = createTeamsClient({
    fetch,
    session: {
      headers: async () => ({ authorization: `Bearer ${spacesToken}` }),
      endpoints: () => Promise.reject(new Error('authz has no endpoints yet')),
    },
  });
  const body = await client.json<AuthzResponse | null>('csa', AUTHZ_URL, { method: 'POST' }).catch((error: unknown) => {
    if (error instanceof CliError && (error.details as { status?: number } | null)?.status === 403) {
      throw new CliError(
        'auth_required',
        'Signed in to Azure AD, but this account has no access to Teams (no Teams license, a personal teams.live.com account, or a tenant policy). Sign in with a work account that uses Teams.',
        { status: 403 },
        { cause: error },
      );
    }
    throw error;
  });
  const skypeToken = body?.tokens?.skypeToken;
  const region = body?.region;
  const chatService = body?.regionGtms?.chatService;
  if (!skypeToken || !region || !chatService) throw new CliError('remote', 'Teams authz returned an unexpected response.');
  return {
    skypeToken,
    expiresAt: new Date(now().getTime() + Number(body.tokens?.expiresIn ?? 3600) * 1000),
    region,
    endpoints: {
      chatService: chatService.replace(/\/+$/, ''),
      csa: `https://teams.microsoft.com/api/csa/${region}`,
      // urlpV2 is the regional async gateway (`https://us-prod.asyncgw.teams.microsoft.com/urlp`). The chat host is not.
      urlp: urlpBase(body.regionGtms?.urlpV2) ?? urlpBase(body.regionGtms?.urlp),
    },
  };
}
