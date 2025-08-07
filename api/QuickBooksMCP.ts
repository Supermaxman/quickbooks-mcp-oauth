import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { QuickBooksService } from "./QuickBooksService.ts";
import { QuickBooksAuthContext } from "../types";

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
    return new QuickBooksService(
      this.env,
      this.props.accessToken,
      this.props.refreshToken
    );
  }

  formatResponse = (
    description: string,
    data: unknown
  ): {
    content: Array<{ type: "text"; text: string }>;
  } => {
    return {
      content: [
        {
          type: "text",
          text: `Success! ${description}\n\nResult:\n${JSON.stringify(
            data,
            null,
            2
          )}`,
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
        return this.formatResponse("Company info retrieved", info);
      }
    );

    return server;
  }
}
