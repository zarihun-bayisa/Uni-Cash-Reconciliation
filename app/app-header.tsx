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
    `group flex shrink-0 items-center gap-3 rounded-xl px-4 py-3 text-sm font-semibold transition duration-200 ${
      pathname === href
        ? "bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-lg shadow-blue-950/30"
        : "text-slate-300 hover:bg-white/10 hover:text-white"
    }`;

  return (
    <aside className="w-full shrink-0 border-b border-slate-800 bg-gradient-to-b from-slate-950 via-slate-950 to-slate-900 text-white shadow-xl shadow-slate-950/10 md:sticky md:top-0 md:flex md:h-screen md:w-72 md:flex-col md:border-b-0 md:border-r">
      <div className="flex items-center gap-3 border-b border-white/10 px-6 py-5 md:py-7">
        <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-400 via-blue-600 to-indigo-700 text-sm font-black tracking-wide text-white shadow-lg shadow-blue-950/50 ring-1 ring-white/15">
          UC
        </div>
        <div>
          <h1 className="text-lg font-bold tracking-tight text-white">
            Uni-Cash
          </h1>
          <p className="mt-1 text-[11px] font-medium uppercase tracking-[0.14em] text-slate-400">
            Payment Reconciliation
          </p>
        </div>
      </div>

      <nav aria-label="Main navigation" className="flex gap-2 overflow-x-auto p-3 md:flex-1 md:flex-col md:gap-1.5 md:p-4">
        <p className="hidden px-4 pb-3 pt-5 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-500 md:block">
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
        <Link
          href="/student-import"
          aria-current={pathname === "/student-import" ? "page" : undefined}
          className={navClass("/student-import")}
        >
          Student Import
        </Link>
        <Link
          href="/reading-checker"
          aria-current={pathname === "/reading-checker" ? "page" : undefined}
          className={navClass("/reading-checker")}
        >
          Reading Checker
        </Link>
      </nav>

      <div className="flex items-center justify-between gap-3 border-t border-white/10 px-5 py-4 md:bg-black/10">
        <span className="inline-flex items-center gap-2 text-xs font-medium text-emerald-300">
          <span className="relative flex h-2 w-2">
            <span className="absolute inline-flex h-full w-full animate-ping rounded-full bg-emerald-400 opacity-40" />
            <span className="relative inline-flex h-2 w-2 rounded-full bg-emerald-400" />
          </span>
          Secure session
        </span>
        <button
          type="button"
          onClick={handleLogout}
          disabled={loggingOut}
          className="rounded-xl border border-white/10 bg-white/5 px-4 py-2 text-sm font-semibold text-slate-200 transition hover:border-red-400/40 hover:bg-red-500/10 hover:text-white disabled:cursor-not-allowed disabled:opacity-60"
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
