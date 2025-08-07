import { createMiddleware } from "hono/factory";
import { HTTPException } from "hono/http-exception";

/* ---------- Bearer-token middleware ---------- */

export const quickBooksBearerTokenAuthMiddleware = createMiddleware<{
  Bindings: Env;
}>(async (c, next) => {
  const auth = c.req.header("Authorization") ?? "";
  if (!auth.startsWith("Bearer ")) {
    throw new HTTPException(401, {
      message: "Missing or invalid access token",
    });
  }

  const accessToken = auth.slice(7);
  const refreshToken = c.req.header("X-QuickBooks-Refresh-Token") ?? "";

  // @ts-expect-error  – Cloudflare Workers executionCtx props
  c.executionCtx.props = { accessToken, refreshToken };
  await next();
});

/* ---------- Helpers ---------- */

export const QUICKBOOKS_DEFAULT_SCOPES = ["com.intuit.quickbooks.accounting"];

export function getQuickBooksAuthEndpoint(
  endpoint: "authorize" | "token"
): string {
  return endpoint === "authorize"
    ? "https://appcenter.intuit.com/connect/oauth2"
    : "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
}

type TokenResponse = {
  access_token: string;
  token_type: string;
  expires_in: number;
  refresh_token?: string;
  scope?: string;
};

function form(params: Record<string, string | undefined>): URLSearchParams {
  const body = new URLSearchParams();
  Object.entries(params).forEach(([k, v]) => {
    if (v) body.append(k, v);
  });
  return body;
}

/* ---------- Code-grant exchange ---------- */

export async function exchangeCodeForToken(
  code: string,
  redirectUri: string,
  clientId: string,
  clientSecret: string,
  codeVerifier?: string,
  scopes: string[] = QUICKBOOKS_DEFAULT_SCOPES
): Promise<TokenResponse> {
  const body = form({
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: codeVerifier,
    scope: scopes.join(" "),
  });

  const basicAuth = btoa(`${clientId}:${clientSecret}`);

  const res = await fetch(getQuickBooksAuthEndpoint("token"), {
    method: "POST",
    headers: {
      Authorization: `Basic ${basicAuth}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  if (!res.ok) throw new Error(`Token exchange failed: ${await res.text()}`);
  return res.json() as Promise<TokenResponse>;
}

/* ---------- Refresh-token flow ---------- */

export async function refreshAccessToken(
  refreshToken: string,
  clientId: string,
  clientSecret: string,
  scopes: string[] = QUICKBOOKS_DEFAULT_SCOPES
): Promise<TokenResponse> {
  const body = form({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: scopes.join(" "),
  });

  const basicAuth = btoa(`${clientId}:${clientSecret}`);

  const res = await fetch(getQuickBooksAuthEndpoint("token"), {
    method: "POST",
    headers: {
      Authorization: `Basic ${basicAuth}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  if (!res.ok) throw new Error(`Refresh failed: ${await res.text()}`);
  return res.json() as Promise<TokenResponse>;
}
