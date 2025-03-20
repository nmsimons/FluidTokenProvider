import {
  app,
  type HttpRequest,
  type HttpResponseInit,
  type InvocationContext,
} from "@azure/functions";
import type { IUser } from "@fluidframework/azure-client";
import { generateToken } from "@fluidframework/server-services-client";
import { ScopeType } from "@fluidframework/protocol-definitions/internal";
import { SecretClient } from "@azure/keyvault-secrets";
import {
  DefaultAzureCredential,
  getBearerTokenProvider,
} from "@azure/identity";

async function getAFRTenantKey(): Promise<string | undefined> {
  if (process.env.AFR_API_KEY) {
    return process.env.AFR_API_KEY;
  }
  const keyVaultUrl = process.env.KEY_VAULT_URL;
  if (!keyVaultUrl) throw new Error("Env KEY_VAULT_URL is empty");
  const credential = new DefaultAzureCredential();
  const client = new SecretClient(keyVaultUrl, credential);
  const tenantKey = await client.getSecret("afrTenantKey");
  return tenantKey.value;
}

async function getAFRTenantId(): Promise<string | undefined> {
  if (process.env.AFR_TENANT_ID) {
    return process.env.AFR_TENANT_ID;
  }
  const keyVaultUrl = process.env.KEY_VAULT_URL;
  if (!keyVaultUrl) throw new Error("Env KEY_VAULT_URL is empty");
  const credential = new DefaultAzureCredential();
  const client = new SecretClient(keyVaultUrl, credential);
  const tenantId = await client.getSecret("tenantId");
  return tenantId.value;
}

export async function AzureFluidRelayTokenProvider(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log(`Http function processed request for url "${req.url}"`);

  const tenantKey = await getAFRTenantKey();

  // The example uses body fallbacks, but the example TokenProvider does not provide a body.
  const body = (await req.json().catch(() => undefined)) as
    | Record<string, string>
    | undefined;

  const tenantId = req.query.get("tenantId") ?? body?.tenantId;
  const documentId = req.query.get("documentId") ?? body?.documentId;
  const userId = req.query.get("userId") ?? body?.userId;
  // The example includes userName, but that is not strictly accepted for generateToken
  const userName = req.query.get("userName") ?? body?.userName;

  if (!tenantId) {
    return {
      status: 400,
      body: "Missing tenantId in the request.",      
    };
  }

  if (!tenantKey) {
    return {
      status: 404,
      body: `No key found for the provided tenantId: ${tenantId}`,      
    };
  }

  const user: IUser & { name: string } = {
    id: userId ?? "test-user",
    name: userName ?? `Test User ${Math.floor(Math.random() * 1000)}`,
  };

  // Will generate the token and returned by an ITokenProvider implementation to use with the AzureClient.
  const token = generateToken(
    tenantId,
    documentId ?? "",
    tenantKey,
    [ScopeType.DocRead, ScopeType.DocWrite, ScopeType.SummaryWrite],
    user
  );

  return {
    status: 200,
    body: token,    
  };
}

export async function AzureOpenAiTokenProvider(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log(`Http function processed request for url "${req.url}"`);

  const credential = new DefaultAzureCredential();
  const tokenProvider = getBearerTokenProvider(
    credential,
    "https://cognitiveservices.azure.com/.default"
  );
  const token = await tokenProvider();

  return {
    status: 200,
    body: token,    
  };
}

app.http("getAfrToken", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: AzureFluidRelayTokenProvider,
});

app.http("getOpenAiToken", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: AzureOpenAiTokenProvider,
});
