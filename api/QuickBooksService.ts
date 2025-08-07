export class QuickBooksService {
  private env: Env;
  private accessToken: string;
  private refreshToken: string;

  constructor(env: Env, accessToken: string, refreshToken: string) {
    this.env = env;
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
  }

  private async makeRequest<T>(
    url: string,
    options: RequestInit = {}
  ): Promise<T> {
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });

    if (res.status === 401) {
      await this.refreshAccessToken();
      return this.makeRequest<T>(url, options);
    }

    if (!res.ok) {
      throw new Error(`QuickBooks API error: ${res.status} ${res.statusText}`);
    }

    return (await res.json()) as T;
  }

  private async refreshAccessToken(): Promise<void> {
    const body = new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: this.refreshToken,
    });

    const basicAuth = btoa(
      `${this.env.QUICKBOOKS_CLIENT_ID}:${this.env.QUICKBOOKS_CLIENT_SECRET}`
    );

    const res = await fetch(
      "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer",
      {
        method: "POST",
        headers: {
          Authorization: `Basic ${basicAuth}`,
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body,
      }
    );

    if (!res.ok) {
      throw new Error(`Failed to refresh access token: ${await res.text()}`);
    }

    const { access_token, refresh_token } = (await res.json()) as {
      access_token: string;
      refresh_token?: string;
    };

    this.accessToken = access_token;
    if (refresh_token) {
      this.refreshToken = refresh_token;
    }
  }

  // Example QuickBooks API call: fetch company info
  async getCompanyInfo(): Promise<unknown> {
    const url = `https://quickbooks.api.intuit.com/v3/company/${this.env.QUICKBOOKS_ACCOUNT_ID}/companyinfo/${this.env.QUICKBOOKS_ACCOUNT_ID}`;
    return this.makeRequest<unknown>(url, { method: "GET" });
  }
}
