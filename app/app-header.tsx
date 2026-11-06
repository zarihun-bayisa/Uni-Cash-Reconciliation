"use client";

import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { useState } from "react";

export default function AppHeader() {
  const pathname = usePathname();
  const router = useRouter();
  const [loggingOut, setLoggingOut] = useState(false);
  const [error, setError] = useState("");

  async function handleLogout() {
    setLoggingOut(true);
    setError("");

    try {
      const response = await fetch("/api/auth/logout", { method: "POST" });
      if (!response.ok) {
        throw new Error("Could not sign out. Please try again.");
      }

      router.replace("/login");
      router.refresh();
    } catch (logoutError) {
      setError(
        logoutError instanceof Error
          ? logoutError.message
          : "Could not sign out. Please try again."
      );
      setLoggingOut(false);
    }
  }

  const navClass = (href: string) =>
    `rounded-lg px-4 py-2 text-sm font-semibold transition ${
      pathname === href
        ? "bg-blue-700 text-white shadow-sm"
        : "text-slate-300 hover:bg-slate-800 hover:text-white"
    }`;

  return (
    <aside className="w-full shrink-0 bg-slate-950 text-white md:sticky md:top-0 md:flex md:h-screen md:w-72 md:flex-col">
      <div className="flex items-center gap-3 border-b border-slate-800 px-6 py-5">
        <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-700 font-bold text-white shadow-sm">
          UC
        </div>
        <div>
          <h1 className="text-lg font-bold tracking-tight">
            Uni-Cash
          </h1>
          <p className="mt-1 text-xs text-slate-400">
            Payment Reconciliation
          </p>
        </div>
      </div>

      <nav aria-label="Main navigation" className="flex gap-2 overflow-x-auto p-4 md:flex-1 md:flex-col">
        <p className="hidden px-4 pb-2 pt-3 text-xs font-semibold uppercase tracking-widest text-slate-500 md:block">
          Workspace
        </p>
        <Link
          href="/"
          aria-current={pathname === "/" ? "page" : undefined}
          className={navClass("/")}
        >
          Dashboard
        </Link>
        <Link
          href="/bill-difference"
          aria-current={pathname === "/bill-difference" ? "page" : undefined}
          className={navClass("/bill-difference")}
        >
          Water Bill Difference
        </Link>
        <Link
          href="/school-bill-difference"
          aria-current={
            pathname === "/school-bill-difference" ? "page" : undefined
          }
          className={navClass("/school-bill-difference")}
        >
          School Bill Difference
        </Link>
      </nav>

      <div className="flex items-center justify-between gap-3 border-t border-slate-800 px-5 py-4">
        <span className="hidden items-center gap-2 text-xs font-medium text-emerald-300 md:inline-flex">
          <span className="h-2 w-2 rounded-full bg-emerald-400" />
          Secure session
        </span>
        <button
          type="button"
          onClick={handleLogout}
          disabled={loggingOut}
          className="rounded-lg border border-slate-700 px-4 py-2 text-sm font-semibold text-slate-200 transition hover:border-red-400 hover:bg-red-950 hover:text-white disabled:cursor-not-allowed disabled:opacity-60"
        >
          {loggingOut ? "Signing out..." : "Sign out"}
        </button>
      </div>
      {error && (
        <p role="alert" className="px-5 pb-4 text-sm text-red-300">
          {error}
        </p>
      )}
    </aside>
  );
}
