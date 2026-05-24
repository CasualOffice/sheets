/** Mirror of the server-side AdminConfig shape. Kept in sync by hand
 *  for the typescript types — the server is the source of truth, this
 *  is only for the UI. */

export type WebhookEvent =
  | 'room.created'
  | 'room.dropped'
  | 'file.uploaded'
  | 'file.saved'
  | 'file.deleted'
  | 'user.joined'
  | 'user.left'
  | 'admin.login'
  | 'admin.login_failed';

export const ALL_WEBHOOK_EVENTS: WebhookEvent[] = [
  'room.created',
  'room.dropped',
  'file.uploaded',
  'file.saved',
  'file.deleted',
  'user.joined',
  'user.left',
  'admin.login',
  'admin.login_failed',
];

export interface WebhookSubscription {
  name: string;
  url: string;
  events: WebhookEvent[];
  secret: string;
  enabled: boolean;
}

export interface AdminConfig {
  branding: {
    appName: string;
    accentColor: string;
    logoUrl: string | null;
  };
  basePath: string;
  storage: {
    backend: 'memory' | 'local' | 's3' | 'postgres';
    local: { path: string };
    s3: {
      endpoint: string;
      region: string;
      bucket: string;
      accessKey: string;
      secretKey: string;
      forcePathStyle: boolean;
      keyPrefix: string;
    };
    postgres: { url: string };
  };
  networking: {
    publicOrigin: string;
    corsOrigins: string;
    trustProxy: string;
    hstsMaxAge: number;
  };
  limits: {
    maxRooms: number;
    maxFileSizeMb: number;
    roomTtlMin: number;
    maxUsersPerRoom: number;
  };
  auth: {
    oidc: {
      enabled: boolean;
      issuer: string;
      clientId: string;
      clientSecret: string;
      redirectUri: string;
    };
    saml: {
      enabled: boolean;
      idpMetadataUrl: string;
      spEntityId: string;
    };
    jwt: {
      enabled: boolean;
      issuer: string;
      defaultTtlSeconds: number;
    };
  };
  webhooks: WebhookSubscription[];
}
