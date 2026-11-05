import { NextRequest, NextResponse } from "next/server";
import {
  getAuthConfiguration,
  SESSION_COOKIE,
  verifySessionToken,
} from "@/lib/auth";

export function proxy(request: NextRequest) {
  const pathname = request.nextUrl.pathname;

  if (
    pathname === "/api/auth/login" ||
    pathname === "/api/auth/logout"
  ) {
    return NextResponse.next();
  }

  let authenticated: boolean;
  try {
    const config = getAuthConfiguration();
    authenticated = verifySessionToken(
      request.cookies.get(SESSION_COOKIE)?.value,
      config
    );
  } catch (error) {
    console.error("Authentication configuration error:", error);
    return pathname.startsWith("/api/")
      ? NextResponse.json(
          { error: "Authentication is not configured on the server." },
          { status: 503 }
        )
      : new NextResponse("Authentication is not configured on the server.", {
          status: 503,
        });
  }

  if (pathname === "/login") {
    return authenticated
      ? NextResponse.redirect(new URL("/", request.url))
      : NextResponse.next();
  }

  if (authenticated) return NextResponse.next();

  return pathname.startsWith("/api/")
    ? NextResponse.json({ error: "Authentication required." }, { status: 401 })
    : NextResponse.redirect(new URL("/login", request.url));
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
