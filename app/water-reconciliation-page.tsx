
"use client";

import { useState } from "react";
import Link from "next/link";
import AppHeader from "./app-header";
import {
  useReconciliation,
  type ReconciliationResult,
} from "./reconciliation-context";

export default function WaterReconciliationPage({
  heading = "Dashboard",
}: {
  heading?: string;
}) {
  const [centralFile, setCentralFile] = useState<File | null>(null);
  const [paidFile, setPaidFile] = useState<File | null>(null);

  const [loading, setLoading] = useState(false);
  const [downloadingDuplicates, setDownloadingDuplicates] = useState(false);
  const [downloadingMatches, setDownloadingMatches] = useState(false);

  const [error, setError] = useState("");
  const { result, setResult } = useReconciliation();

  // Compare Central and Paid Excel files
  async function handleCompare() {
    if (!centralFile || !paidFile) {
      setError("Please select both Excel files.");
      return;
    }

    setLoading(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();

      formData.append("centralFile", centralFile);
      formData.append("paidFile", paidFile);

      const response = await fetch("/api/reconcile", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "Reconciliation failed.");
      }

      setResult(data as ReconciliationResult);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Something went wrong during reconciliation."
      );
    } finally {
      setLoading(false);
    }
  }

  async function handleDownloadDuplicates() {
    if (!result || result.duplicateRows.length === 0) {
      setError("There are no duplicate rows to download.");
      return;
    }

    setDownloadingDuplicates(true);
    setError("");

    try {
      const response = await fetch("/api/duplicate-report", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          duplicateRows: result.duplicateRows,
        }),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || "Failed to generate duplicate report.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = url;
      link.download = "uni-cash-duplicate-rows.xlsx";

      document.body.appendChild(link);
      link.click();
      link.remove();

      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Could not download the duplicate report."
      );
    } finally {
      setDownloadingDuplicates(false);
    }
  }

  async function handleDownloadMatches() {
    if (!result || result.matchedPayments.length === 0) {
      setError("There are no matched payments to download.");
      return;
    }

    setDownloadingMatches(true);
    setError("");

    try {
      const response = await fetch("/api/match-report", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          matchedPayments: result.matchedPayments,
        }),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || "Failed to generate matched report.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = url;
      link.download = "uni-cash-matched-payments.xlsx";

      document.body.appendChild(link);
      link.click();
      link.remove();

      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Could not download the matched payments report."
      );
    } finally {
      setDownloadingMatches(false);
    }
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 md:flex">

      <AppHeader />

      <main className="min-w-0 flex-1">
      <section className="mx-auto max-w-7xl px-6 py-10">

        {/* PAGE TITLE */}
        <div className="mb-8">
          <p className="text-sm font-semibold uppercase tracking-[0.16em] text-blue-700">
          {heading}
          </p>
          <h2 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">
            Payment Reconciliation
          </h2>

          <p className="mt-2 max-w-2xl text-slate-500">
            Upload Central and Paid Excel files to identify payment differences.
          </p>
        </div>

        {/* FILE UPLOAD SECTION */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">

          {/* CENTRAL FILE */}
          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm transition hover:border-blue-200 hover:shadow-md">

            <h3 className="text-lg font-semibold mb-2">
              Central File
            </h3>

            <p className="text-sm text-slate-500 mb-5">
              Upload the Central payment export.
            </p>

            <label className="border-2 border-dashed border-blue-300 rounded-xl p-8 flex flex-col items-center justify-center cursor-pointer hover:bg-blue-50 transition">

              <span className="text-4xl mb-3">📊</span>

              <span className="font-medium text-blue-800">
                Select Central Excel File
              </span>

              <span className="text-xs text-slate-400 mt-2">
                XLSX or XLS format
              </span>

              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  setCentralFile(e.target.files?.[0] || null);
                  setResult(null);
                  setError("");
                }}
              />

            </label>

            {centralFile && (
              <div className="mt-4 text-sm text-green-700 break-all">
                ✓ {centralFile.name}
              </div>
            )}

          </div>

          {/* PAID FILE */}
          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm transition hover:border-emerald-200 hover:shadow-md">

            <h3 className="text-lg font-semibold mb-2">
              Paid File
            </h3>

            <p className="text-sm text-slate-500 mb-5">
              Upload the export from your billing system.
            </p>

            <label className="border-2 border-dashed border-green-300 rounded-xl p-8 flex flex-col items-center justify-center cursor-pointer hover:bg-green-50 transition">

              <span className="text-4xl mb-3">📁</span>

              <span className="font-medium text-green-800">
                Select Uni-Cash File
              </span>

              <span className="text-xs text-slate-400 mt-2">
                XLSX or XLS format
              </span>

              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  setPaidFile(e.target.files?.[0] || null);
                  setResult(null);
                  setError("");
                }}
              />

            </label>

            {paidFile && (
              <div className="mt-4 text-sm text-green-700 break-all">
                ✓ {paidFile.name}
              </div>
            )}

          </div>

        </div>

        {/* COMPARE BUTTON */}
        <div className="mt-8 text-center">

          <button
            onClick={handleCompare}
            disabled={
              !centralFile ||
              !paidFile ||
              loading ||
              downloadingDuplicates ||
              downloadingMatches
            }
            className="rounded-xl bg-blue-700 px-10 py-4 font-semibold text-white shadow-lg shadow-blue-700/20 transition hover:-translate-y-0.5 hover:bg-blue-800 disabled:cursor-not-allowed disabled:bg-slate-400 disabled:shadow-none"
          >
            {loading ? "Comparing Payments..." : "Compare Payments"}
          </button>

          <p className="text-sm text-slate-500 mt-3">
            Select both Excel files before comparing.
          </p>

        </div>

        {/* ERROR MESSAGE */}
        {error && (
          <div className="mt-6 bg-red-50 border border-red-200 text-red-700 p-4 rounded-lg">
            {error}
          </div>
        )}

        {/* RESULTS */}
        {result && (
          <>

            {/* SUMMARY CARDS */}
            <div className="mt-10 grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-4">

              {/* TOTAL */}
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                <p className="text-sm font-medium text-slate-500">
                  Total Central Payments
                </p>

                <h3 className="text-3xl font-bold mt-2">
                  {result.totalPayments.toLocaleString()}
                </h3>
              </div>

              {/* DUPLICATES REMOVED */}
              <div className="rounded-2xl border border-amber-100 bg-white p-6 shadow-sm">
                <p className="text-sm font-medium text-slate-500">
                  Duplicate Rows Removed
                </p>

                <h3 className="text-3xl font-bold text-amber-600 mt-2">
                  {result.duplicateCount.toLocaleString()}
                </h3>
                <p className="mt-2 text-xs text-slate-500">
                  Rows are excluded only when Customer ID, amount, and merchant
                  reference match. Different references count as separate
                  payments.
                </p>
                <button
                  type="button"
                  onClick={handleDownloadDuplicates}
                  disabled={
                    result.duplicateRows.length === 0 ||
                    downloadingDuplicates ||
                    loading
                  }
                  className="mt-4 rounded-lg bg-amber-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-amber-700 disabled:cursor-not-allowed disabled:bg-slate-400"
                >
                  {downloadingDuplicates
                    ? "Generating Report..."
                    : "⬇ Download Duplicate Excel"}
                </button>
              </div>

              {/* MATCHED */}
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                <p className="text-sm font-medium text-slate-500">
                  Matched Payments
                </p>

                <h3 className="text-3xl font-bold text-green-600 mt-2">
                  {result.matchedCount.toLocaleString()}
                </h3>
                <button
                  type="button"
                  onClick={handleDownloadMatches}
                  disabled={
                    result.matchedPayments.length === 0 ||
                    downloadingMatches ||
                    downloadingDuplicates ||
                    loading
                  }
                  className="mt-4 rounded-lg bg-green-700 px-4 py-2 text-sm font-semibold text-white transition hover:bg-green-800 disabled:cursor-not-allowed disabled:bg-slate-400"
                >
                  {downloadingMatches
                    ? "Generating Report..."
                    : "⬇ Download Matched Excel"}
                </button>
              </div>

              {/* DIFFERENCES */}
              <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
                <p className="text-sm font-medium text-slate-500">
                  Differences
                </p>

                <h3 className="text-3xl font-bold text-red-600 mt-2">
                  {result.differenceCount.toLocaleString()}
                </h3>
                <Link
                  href="/bill-difference"
                  className="mt-4 inline-flex rounded-lg bg-red-700 px-4 py-2 text-sm font-semibold text-white transition hover:bg-red-800"
                >
                  View Bill Difference
                </Link>
              </div>

            </div>

          </>
        )}

      </section>

      {/* FOOTER */}
      <footer className="border-t border-slate-200 bg-white mt-12">
        <div className="max-w-7xl mx-auto px-6 py-5 text-center text-sm text-slate-500">
          Central Payment Reconciliation System
        </div>
      </footer>

      </main>
    </div>
  );
}
