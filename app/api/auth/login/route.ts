import { NextRequest, NextResponse } from "next/server";
import {
  createSessionToken,
  credentialsAreValid,
  getAuthConfiguration,
  SESSION_COOKIE,
  SESSION_MAX_AGE_SECONDS,
} from "@/lib/auth";
import type { AuthConfiguration } from "@/lib/auth";

export async function POST(request: NextRequest) {
  let config: AuthConfiguration;
  try {
    config = getAuthConfiguration();
  } catch (error) {
    console.error("Authentication configuration error:", error);
    return NextResponse.json(
      { error: "Login is not configured on the server." },
      { status: 503 }
    );
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json(
      { error: "Enter a username and password." },
      { status: 400 }
    );
  }

  if (
    typeof body !== "object" ||
    body === null ||
    !("username" in body) ||
    !("password" in body) ||
    typeof body.username !== "string" ||
    typeof body.password !== "string"
  ) {
    return NextResponse.json(
      { error: "Enter a username and password." },
      { status: 400 }
    );
  }

  if (!credentialsAreValid(body.username, body.password, config)) {
    return NextResponse.json(
      { error: "The username or password is incorrect." },
      { status: 401 }
    );
  }

  const response = NextResponse.json({ success: true });
  response.cookies.set(
    SESSION_COOKIE,
    createSessionToken(config.username, config.secret),
    {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      path: "/",
      maxAge: SESSION_MAX_AGE_SECONDS,
    }
  );

  return response;
}
