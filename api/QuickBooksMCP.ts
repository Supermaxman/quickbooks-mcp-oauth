import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { QuickBooksService } from "./QuickBooksService.ts";
import { QuickBooksAuthContext } from "../types";
import {
  UpdateInvoiceSchema,
  CopyInvoiceSchema,
  CreateInvoiceSchema,
} from "./lib/quickbooks-types";

/**
 * The `QuickBooksMCP` class exposes the QuickBooks API via the Model Context Protocol
 * for consumption by API Agents
 */
export class QuickBooksMCP extends McpAgent<
  Env,
  unknown,
  QuickBooksAuthContext
> {
  async init() {}

  get quickBooksService() {
    return new QuickBooksService(this.env, this.props.accessToken);
  }

  formatResponse = (
    data: unknown
  ): {
    content: Array<{ type: "text"; text: string }>;
  } => {
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(data, null, 2),
        },
      ],
    };
  };

  get server() {
    const server = new McpServer(
      {
        name: "QuickBooks Service",
        description: "QuickBooks MCP Server for QuickBooks",
        version: "1.0.0",
      },
      {
        instructions:
          "This MCP server is for the QuickBooks API. It can be used to get the user's QuickBooks data, create invoices, and more.",
      }
    );

    server.tool(
      "getCompanyInfo",
      "Get QuickBooks company information",
      async () => {
        const info = await this.quickBooksService.getCompanyInfo();
        return this.formatResponse(info);
      }
    );

    server.tool(
      "getInvoices",
      "Get QuickBooks invoices, sorted by due date descending. Returns a list of invoices.",
      {
        page: z
          .number()
          .min(0)
          .default(0)
          .describe("The page number to get (0-indexed, default 0)"),
        pageSize: z
          .number()
          .min(1)
          .max(20)
          .default(10)
          .describe("The number of invoices to get per page (default 10)"),
      },
      async ({ page, pageSize }) => {
        const invoices = await this.quickBooksService.getInvoices(
          page,
          pageSize
        );
        return this.formatResponse(invoices);
      }
    );

    server.tool(
      "getCustomers",
      "Get QuickBooks customers, sorted by name ascending. Returns a list of customers.",
      {
        page: z
          .number()
          .min(0)
          .default(0)
          .describe("The page number to get (0-indexed, default 0)"),
        pageSize: z
          .number()
          .min(1)
          .max(20)
          .default(10)
          .describe("The number of customers to get per page (default 10)"),
      },
      async ({ page, pageSize }) => {
        const customers = await this.quickBooksService.getCustomers(
          page,
          pageSize
        );
        return this.formatResponse(customers);
      }
    );

    server.tool(
      "updateInvoice",
      "Update an existing invoice by Id using sparse update (only provided fields will change). Returns the updated invoice.",
      UpdateInvoiceSchema.shape,
      async (args) => {
        const parsed = UpdateInvoiceSchema.parse(args);
        const updated = await this.quickBooksService.updateInvoice(parsed);
        return this.formatResponse(updated);
      }
    );

    server.tool(
      "copyInvoice",
      "Create a new invoice by copying fields from an existing invoice Id. Returns the created invoice.",
      CopyInvoiceSchema.shape,
      async (args) => {
        const { Id } = CopyInvoiceSchema.parse(args);
        const created = await this.quickBooksService.copyInvoice(Id);
        return this.formatResponse(created);
      }
    );

    server.tool(
      "createInvoice",
      "Create a new invoice from scratch. Returns the created invoice.",
      CreateInvoiceSchema.shape,
      async (args) => {
        const parsed = CreateInvoiceSchema.parse(args);
        const created = await this.quickBooksService.createInvoice(parsed);
        return this.formatResponse(created);
      }
    );

    return server;
  }
}
