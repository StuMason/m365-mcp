export interface TokenData {
  access_token: string;
  refresh_token: string;
  expires_at: string; // ISO 8601
  scopes: string;
}

export interface AuthConfig {
  clientId: string;
  clientSecret?: string;
  tenantId: string;
}
