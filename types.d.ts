// Environment variables and bindings
interface Env {
  QUICKBOOKS_CLIENT_ID: string;
  QUICKBOOKS_CLIENT_SECRET: string;
  QUICKBOOKS_ACCOUNT_ID: string;
  QUICKBOOKS_MCP_OBJECT: DurableObjectNamespace;
}

export type Todo = {
  id: string;
  text: string;
  completed: boolean;
};

// Context from the auth process, extracted from the Stytch auth token JWT
// and provided to the MCP Server as this.props
type AuthenticationContext = {
  claims: {
    iss: string;
    scope: string;
    sub: string;
    aud: string[];
    client_id: string;
    exp: number;
    iat: number;
    nbf: number;
    jti: string;
  };
  accessToken: string;
};

// Context from the QuickBooks OAuth process
export type QuickBooksAuthContext = {
  accessToken: string;
  refreshToken: string;
  expiresIn?: number;
  tokenType?: string;
  scope?: string;
};
