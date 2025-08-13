import {
  QuickBooksQueryResponse,
  QuickBooksTable,
  TABLE_DEFAULT_SORT,
} from "./lib/quickbooks-types";

export class QuickBooksService {
  private env: Env;
  private accessToken: string;
  private baseUrl: string;

  constructor(env: Env, accessToken: string) {
    this.env = env;
    this.accessToken = accessToken;
    this.baseUrl = `https://quickbooks.api.intuit.com/v3/company/${this.env.QUICKBOOKS_ACCOUNT_ID}`;
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
        Accept: "application/json",
        ...options.headers,
      },
    });

    if (!res.ok) {
      throw new Error(`QuickBooks API error: ${res.status} ${res.statusText}`);
    }

    return (await res.json()) as T;
  }

  // TODO add filtering support
  // See https://developer.intuit.com/app/developer/qbo/docs/learn/explore-the-quickbooks-online-api/data-queries
  private async getQueryData<TRow, K extends QuickBooksTable>(
    table: K,
    offset: number = 0,
    maxResults: number = 10
  ): Promise<QuickBooksQueryResponse<TRow, K>> {
    const sort = TABLE_DEFAULT_SORT[table];
    const params = new URLSearchParams({
      query: `SELECT * FROM ${table} ORDERBY ${sort} STARTPOSITION ${offset} MAXRESULTS ${maxResults}`,
    });

    const url = `${this.baseUrl}/query?${params.toString()}`;
    return this.makeRequest<QuickBooksQueryResponse<TRow, K>>(url, {
      method: "GET",
    });
  }

  // Example QuickBooks API call: fetch company info
  async getCompanyInfo(): Promise<unknown> {
    const url = `${this.baseUrl}/companyinfo/${this.env.QUICKBOOKS_ACCOUNT_ID}`;
    return this.makeRequest<unknown>(url, { method: "GET" });
  }

  async getInvoices(
    page: number = 0,
    pageSize: number = 10
  ): Promise<unknown[]> {
    const resp = await this.getQueryData<unknown, "Invoice">(
      "Invoice",
      page,
      pageSize
    );
    return resp.QueryResponse.Invoice;
  }

  async getCustomers(
    page: number = 0,
    pageSize: number = 10
  ): Promise<unknown[]> {
    const resp = await this.getQueryData<unknown, "Customer">(
      "Customer",
      page,
      pageSize
    );
    return resp.QueryResponse.Customer;
  }
}
