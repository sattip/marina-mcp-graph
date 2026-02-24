import { ConfidentialClientApplication } from "@azure/msal-node";
import type { GraphConfig } from "./types.js";

export class GraphClient {
  private app: ConfidentialClientApplication;
  private config: GraphConfig;
  private tokenCache: { token: string; expiresAt: number } | null = null;

  constructor(config: GraphConfig) {
    this.config = config;
    this.app = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    });
  }

  private async getToken(): Promise<string> {
    if (this.tokenCache && Date.now() < this.tokenCache.expiresAt - 60_000) {
      return this.tokenCache.token;
    }

    const result = await this.app.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });

    if (!result?.accessToken) {
      throw new Error("Failed to acquire Graph API token");
    }

    this.tokenCache = {
      token: result.accessToken,
      expiresAt: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600_000,
    };

    return result.accessToken;
  }

  async request(method: string, path: string, body?: unknown): Promise<unknown> {
    const token = await this.getToken();
    const url = `https://graph.microsoft.com/v1.0${path}`;

    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph API ${method} ${path} failed (${res.status}): ${text}`);
    }

    if (res.status === 204) return { success: true };

    return res.json();
  }

  get defaultFromEmail(): string {
    return this.config.defaultFromEmail;
  }
}
