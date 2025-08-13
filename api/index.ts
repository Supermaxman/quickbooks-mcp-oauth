import { QuickBooksMCP } from "./QuickBooksMCP.ts";
import {
  quickBooksBearerTokenAuthMiddleware,
  getQuickBooksAuthEndpoint,
  exchangeCodeForToken,
  refreshAccessToken,
} from "./lib/quickbooks-auth.ts";
import { cors } from "hono/cors";
import { Hono } from "hono";

// Export the QuickBooksMCP class so the Worker runtime can find it
export { QuickBooksMCP };

// Store registered clients in memory (in production, use a database)
interface RegisteredClient {
  client_id: string;
  client_name: string;
  redirect_uris: string[];
  grant_types: string[];
  response_types: string[];
  scope?: string;
  token_endpoint_auth_method: string;
  created_at: number;
}
const registeredClients = new Map<string, RegisteredClient>();

export default new Hono<{ Bindings: Env }>()
  .use(cors())

  // OAuth Authorization Server Discovery
  .get("/.well-known/oauth-authorization-server", async (c) => {
    const url = new URL(c.req.url);
    return c.json({
      issuer: url.origin,
      authorization_endpoint: `${url.origin}/authorize`,
      token_endpoint: `${url.origin}/token`,
      registration_endpoint: `${url.origin}/register`,
      response_types_supported: ["code"],
      response_modes_supported: ["query"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      token_endpoint_auth_methods_supported: ["none"],
      code_challenge_methods_supported: ["S256", "plain"],
      scopes_supported: ["com.intuit.quickbooks.accounting"],
    });
  })

  // Dynamic Client Registration endpoint
  .post("/register", async (c) => {
    const body = await c.req.json();

    // Generate a client ID
    const clientId = crypto.randomUUID();

    // Store the client registration
    registeredClients.set(clientId, {
      client_id: clientId,
      client_name: body.client_name || "MCP Client",
      redirect_uris: body.redirect_uris || [],
      grant_types: body.grant_types || ["authorization_code", "refresh_token"],
      response_types: body.response_types || ["code"],
      scope: body.scope,
      token_endpoint_auth_method: "none",
      created_at: Date.now(),
    });

    // Return the client registration response
    return c.json(
      {
        client_id: clientId,
        client_name: body.client_name || "MCP Client",
        redirect_uris: body.redirect_uris || [],
        grant_types: body.grant_types || [
          "authorization_code",
          "refresh_token",
        ],
        response_types: body.response_types || ["code"],
        scope: body.scope,
        token_endpoint_auth_method: "none",
      },
      201
    );
  })

  // Authorization endpoint - redirects to QuickBooks
  .get("/authorize", async (c) => {
    const url = new URL(c.req.url);
    const quickbooksAuthUrl = new URL(getQuickBooksAuthEndpoint("authorize"));

    // Copy all query parameters except client_id
    url.searchParams.forEach((value, key) => {
      if (key !== "client_id") {
        quickbooksAuthUrl.searchParams.set(key, value);
      }
    });

    // Use our QuickBooks app's client_id
    quickbooksAuthUrl.searchParams.set("client_id", c.env.QUICKBOOKS_CLIENT_ID);

    const m = quickbooksAuthUrl.searchParams.get("code_challenge_method"); // 'S256' | 'plain' | null
    const cc = quickbooksAuthUrl.searchParams.get("code_challenge");
    const ru = quickbooksAuthUrl.searchParams.get("redirect_uri");
    console.log("PKCE authorize:", {
      method: m,
      code_challenge: cc?.slice(0, 8) + "…",
      redirect_uri: ru,
    });

    // Redirect to QuickBooks's authorization page
    return c.redirect(quickbooksAuthUrl.toString());
  })

  // Token exchange endpoint
  .post("/token", async (c) => {
    const body = await c.req.parseBody();
    const cv = body.code_verifier as string | undefined;
    console.log("PKCE token:", {
      code_verifier_len: cv?.length,
      redirect_uri: body.redirect_uri,
    });

    try {
      if (body.grant_type === "authorization_code") {
        const result = await exchangeCodeForToken(
          body.code as string,
          body.redirect_uri as string,
          c.env.QUICKBOOKS_CLIENT_ID,
          c.env.QUICKBOOKS_CLIENT_SECRET,
          body.code_verifier as string | undefined,
          (body.scope as string | undefined) ||
            (typeof body.scope === "string"
              ? (body.scope as string)
              : undefined)
        );
        return c.json(result);
      } else if (body.grant_type === "refresh_token") {
        const result = await refreshAccessToken(
          body.refresh_token as string,
          c.env.QUICKBOOKS_CLIENT_ID,
          c.env.QUICKBOOKS_CLIENT_SECRET
        );
        return c.json(result);
      }
    } catch (err) {
      // Pass through OAuth errors from Microsoft
      const e = err as unknown as {
        status?: number;
        body?: unknown;
        name?: string;
      };
      if (e && e.name === "OAuthHttpError") {
        const allowedStatuses = [
          400, 401, 403, 404, 405, 409, 410, 415, 422, 429, 500, 502, 503, 504,
        ] as const;
        const statusCandidate = (e.status as number) || 400;
        const status = (allowedStatuses as readonly number[]).includes(
          statusCandidate
        )
          ? (statusCandidate as
              | 400
              | 401
              | 403
              | 404
              | 405
              | 409
              | 410
              | 415
              | 422
              | 429
              | 500
              | 502
              | 503
              | 504)
          : (400 as const);
        return c.json(e.body ?? { error: "invalid_request" }, { status });
      }
      throw err;
    }

    return c.json({ error: "unsupported_grant_type" }, 400);
  })

  // QuickBooks MCP endpoints
  .use("/sse/*", quickBooksBearerTokenAuthMiddleware)
  .route(
    "/sse",
    new Hono().mount(
      "/",
      QuickBooksMCP.serveSSE("/sse", { binding: "QUICKBOOKS_MCP_OBJECT" }).fetch
    )
  )

  .use("/mcp", quickBooksBearerTokenAuthMiddleware)
  .route(
    "/mcp",
    new Hono().mount(
      "/",
      QuickBooksMCP.serve("/mcp", { binding: "QUICKBOOKS_MCP_OBJECT" }).fetch
    )
  )

  // Health check endpoint
  .get("/", (c) => c.text("QuickBooks MCP Server is running"));
