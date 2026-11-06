import { createHash, createHmac, timingSafeEqual } from "node:crypto";

export const SESSION_COOKIE = "unicash_session";
export const SESSION_MAX_AGE_SECONDS = 60 * 60 * 8;

export interface AuthConfiguration {
  username: string;
  password: string;
  secret: string;
}

export function getAuthConfiguration(): AuthConfiguration {
  const isProduction = process.env.NODE_ENV === "production";
  const username =
    process.env.UNICASH_USERNAME ?? (isProduction ? "" : "Unicash");
  const password =
    process.env.UNICASH_PASSWORD ?? (isProduction ? "" : "password");

  if (!username || !password) {
    throw new Error(
      "Set UNICASH_USERNAME and UNICASH_PASSWORD in the server environment."
    );
  }

  const secret = createHash("sha256")
    .update(
      JSON.stringify([
        "central-payment-reconciliation-session",
        username,
        password,
      ])
    )
    .digest("hex");

  return {
    username,
    password,
    secret,
  };
}

function safeEqual(left: string, right: string): boolean {
  const leftHash = createHash("sha256").update(left).digest();
  const rightHash = createHash("sha256").update(right).digest();
  return timingSafeEqual(leftHash, rightHash);
}

export function credentialsAreValid(
  username: string,
  password: string,
  config: AuthConfiguration
): boolean {
  return (
    safeEqual(username, config.username) &&
    safeEqual(password, config.password)
  );
}

export function createSessionToken(
  username: string,
  secret: string
): string {
  const payload = Buffer.from(
    JSON.stringify({
      username,
      expiresAt: Date.now() + SESSION_MAX_AGE_SECONDS * 1000,
    })
  ).toString("base64url");
  const signature = createHmac("sha256", secret)
    .update(payload)
    .digest("base64url");

  return `${payload}.${signature}`;
}

export function verifySessionToken(
  token: string | undefined,
  config: AuthConfiguration
): boolean {
  if (!token) return false;

  const [payload, signature, extra] = token.split(".");
  if (!payload || !signature || extra !== undefined) return false;

  const expectedSignature = createHmac("sha256", config.secret)
    .update(payload)
    .digest("base64url");

  if (!safeEqual(signature, expectedSignature)) return false;

  try {
    const session: unknown = JSON.parse(
      Buffer.from(payload, "base64url").toString("utf8")
    );

    if (
      typeof session !== "object" ||
      session === null ||
      !("username" in session) ||
      !("expiresAt" in session) ||
      typeof session.username !== "string" ||
      typeof session.expiresAt !== "number"
    ) {
      return false;
    }

    return (
      safeEqual(session.username, config.username) &&
      session.expiresAt > Date.now()
    );
  } catch {
    return false;
  }
}
