
"use client";

import { useState } from "react";
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
  const [centralFiles, setCentralFiles] = useState<File[]>([]);
  const [paidFiles, setPaidFiles] = useState<File[]>([]);

  const [loading, setLoading] = useState(false);
  const [downloadingDuplicates, setDownloadingDuplicates] = useState(false);
  const [downloadingMatches, setDownloadingMatches] = useState(false);
  const [downloadingDifferences, setDownloadingDifferences] = useState(false);
  const [showDifferences, setShowDifferences] = useState(false);

  const [error, setError] = useState("");
  const { result, setResult } = useReconciliation();

  // Compare Central and Paid Excel files
  async function handleCompare() {
    if (centralFiles.length === 0 || paidFiles.length === 0) {
      setError("Please select both Excel files.");
      return;
    }

    setLoading(true);
    setError("");
    setResult(null);
    setShowDifferences(false);

    try {
      const formData = new FormData();

      centralFiles.forEach((file) => formData.append("centralFile", file));
      paidFiles.forEach((file) => formData.append("paidFile", file));

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

  async function handleDownloadDifferences() {
    if (!result || result.differences.length === 0) {
      setError("There are no bill differences to download.");
      return;
    }

    setDownloadingDifferences(true);
    setError("");

    try {
      const response = await fetch("/api/water-report", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          differences: result.differences,
        }),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || "Failed to generate the difference report.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "unicash-bill-difference.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the bill difference report."
      );
    } finally {
      setDownloadingDifferences(false);
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50/70 text-slate-900 md:flex">

      <AppHeader />

      <main className="min-w-0 flex-1">
      <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8 sm:py-10">

        {/* PAGE TITLE */}
        <div className="relative mb-8 overflow-hidden rounded-[2rem] bg-gradient-to-br from-slate-950 via-blue-950 to-indigo-900 px-6 py-8 text-white shadow-2xl shadow-blue-950/15 sm:px-9 sm:py-10">
          <div aria-hidden="true" className="absolute -right-10 -top-28 h-80 w-80 rounded-full border-[40px] border-white/[0.04]" />
          <div aria-hidden="true" className="absolute -bottom-36 right-40 h-64 w-64 rounded-full bg-blue-400/20 blur-3xl" />
          <div className="relative">
            <p className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.18em] text-blue-100">
              <span className="h-1.5 w-1.5 rounded-full bg-cyan-300" />
              {heading}
            </p>
            <h2 className="mt-4 text-3xl font-bold tracking-tight sm:text-4xl">
              Payment Reconciliation
            </h2>
            <p className="mt-3 max-w-2xl text-sm leading-6 text-blue-100 sm:text-base">
              Upload Central and Paid Excel files to identify payment differences.
            </p>
          </div>
        </div>

        {/* FILE UPLOAD SECTION */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">

          {/* CENTRAL FILE */}
          <div className="group rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-lg shadow-slate-200/40 transition duration-200 hover:-translate-y-1 hover:border-blue-200 hover:shadow-xl hover:shadow-blue-900/10 sm:p-7">

            <div className="mb-5 flex items-center gap-4">
              <span className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-50 text-2xl ring-1 ring-blue-100">📊</span>
              <div>
                <h3 className="text-lg font-bold tracking-tight text-slate-900">
                  Central File
                </h3>
                <p className="mt-1 text-xs font-medium uppercase tracking-wider text-blue-700">Source payments</p>
              </div>
            </div>

            <p className="text-sm text-slate-500 mb-5">
              Upload the Central payment export.
            </p>

            <label className="flex min-h-40 cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-blue-200 bg-blue-50/50 p-7 text-center transition hover:border-blue-400 hover:bg-blue-50">

              <span className="font-medium text-blue-800">
                Select Central Excel Files
              </span>

              <span className="text-xs text-slate-400 mt-2">
                XLSX or XLS format
              </span>

              <input
                type="file"
                accept=".xlsx,.xls"
                multiple
                className="hidden"
                onChange={(e) => {
                  setCentralFiles(Array.from(e.target.files ?? []));
                  setResult(null);
                  setShowDifferences(false);
                  setError("");
                }}
              />

            </label>

            {centralFiles.length > 0 && (
              <div className="mt-4 flex items-center gap-2 rounded-xl bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-800 break-all ring-1 ring-emerald-100">
                <span aria-hidden="true">✓</span>
                <span>{centralFiles.map((file) => file.name).join(", ")}</span>
              </div>
            )}

          </div>

          {/* PAID FILE */}
          <div className="group rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-lg shadow-slate-200/40 transition duration-200 hover:-translate-y-1 hover:border-emerald-200 hover:shadow-xl hover:shadow-emerald-900/10 sm:p-7">

            <div className="mb-5 flex items-center gap-4">
              <span className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-50 text-2xl ring-1 ring-emerald-100">📁</span>
              <div>
                <h3 className="text-lg font-bold tracking-tight text-slate-900">
                  Paid File
                </h3>
                <p className="mt-1 text-xs font-medium uppercase tracking-wider text-emerald-700">Billing export</p>
              </div>
            </div>

            <p className="text-sm text-slate-500 mb-5">
              Upload the export from your billing system.
            </p>

            <label className="flex min-h-40 cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-emerald-200 bg-emerald-50/50 p-7 text-center transition hover:border-emerald-400 hover:bg-emerald-50">

              <span className="font-medium text-green-800">
                Select Uni-Cash Files
              </span>

              <span className="text-xs text-slate-400 mt-2">
                XLSX or XLS format
              </span>

              <input
                type="file"
                accept=".xlsx,.xls"
                multiple
                className="hidden"
                onChange={(e) => {
                  setPaidFiles(Array.from(e.target.files ?? []));
                  setResult(null);
                  setShowDifferences(false);
                  setError("");
                }}
              />

            </label>

            {paidFiles.length > 0 && (
              <div className="mt-4 flex items-center gap-2 rounded-xl bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-800 break-all ring-1 ring-emerald-100">
                <span aria-hidden="true">✓</span>
                <span>{paidFiles.map((file) => file.name).join(", ")}</span>
              </div>
            )}

          </div>

        </div>

        {/* COMPARE BUTTON */}
        <div className="mt-8 text-center">

          <button
            onClick={handleCompare}
            disabled={
              centralFiles.length === 0 ||
              paidFiles.length === 0 ||
              loading ||
              downloadingDuplicates ||
              downloadingMatches ||
              downloadingDifferences
            }
            className="rounded-2xl bg-gradient-to-r from-blue-700 via-blue-700 to-indigo-700 px-10 py-4 font-semibold text-white shadow-xl shadow-blue-800/25 transition hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-blue-800/30 disabled:cursor-not-allowed disabled:from-slate-400 disabled:to-slate-400 disabled:shadow-none"
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
              <div className="rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-lg shadow-slate-200/40">
                <p className="text-sm font-medium text-slate-500">
                  Total Central Payments
                </p>

                <h3 className="text-3xl font-bold mt-2">
                  {result.totalPayments.toLocaleString()}
                </h3>
              </div>

              {/* DUPLICATES REMOVED */}
              <div className="rounded-3xl border border-amber-100 bg-gradient-to-br from-white to-amber-50/60 p-6 shadow-lg shadow-amber-900/5">
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
              <div className="rounded-3xl border border-emerald-100 bg-gradient-to-br from-white to-emerald-50/60 p-6 shadow-lg shadow-emerald-900/5">
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
              <div className="rounded-3xl border border-rose-100 bg-gradient-to-br from-white to-rose-50/60 p-6 shadow-lg shadow-rose-900/5">
                <p className="text-sm font-medium text-slate-500">
                  Differences
                </p>

                <h3 className="text-3xl font-bold text-red-600 mt-2">
                  {result.differenceCount.toLocaleString()}
                </h3>
                <button
                  type="button"
                  onClick={() => setShowDifferences((visible) => !visible)}
                  className="mt-4 inline-flex rounded-lg bg-red-700 px-4 py-2 text-sm font-semibold text-white transition hover:bg-red-800"
                >
                  {showDifferences ? "Hide Bill Difference" : "View Bill Difference"}
                </button>
              </div>

            </div>

            {showDifferences && (
              <div className="mt-8 overflow-hidden rounded-3xl border border-slate-200/80 bg-white/95 shadow-xl shadow-slate-200/50">
                <div className="flex flex-col gap-4 border-b border-slate-100 bg-gradient-to-r from-white to-rose-50/60 p-6 sm:flex-row sm:items-center sm:justify-between sm:p-7">
                  <div>
                    <h3 className="text-xl font-bold">Bill Differences</h3>
                    <p className="mt-1 text-sm text-slate-500">
                      Showing {result.differences.length.toLocaleString()} records.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={handleDownloadDifferences}
                    disabled={
                      result.differences.length === 0 ||
                      downloadingDifferences ||
                      downloadingDuplicates ||
                      downloadingMatches ||
                      loading
                    }
                    className="shrink-0 rounded-lg bg-red-700 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-red-800 disabled:cursor-not-allowed disabled:bg-slate-400"
                  >
                    {downloadingDifferences
                      ? "Generating Report..."
                      : "Download Difference Excel"}
                  </button>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-100/80 text-xs uppercase tracking-wide text-slate-600">
                      <tr>
                        <th className="whitespace-nowrap p-4">Customer ID</th>
                        <th className="whitespace-nowrap p-4">Central Amount</th>
                        <th className="whitespace-nowrap p-4">
                          Central Merchant Reference
                        </th>
                        <th className="whitespace-nowrap p-4">Paid Bill ID</th>
                        <th className="whitespace-nowrap p-4">Paid Amount</th>
                        <th className="whitespace-nowrap p-4">Payment Status</th>
                        <th className="whitespace-nowrap p-4">Occurrences</th>
                        <th className="min-w-64 p-4">Difference Note</th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.differences.map((item, index) => (
                        <tr
                          key={`${item.customerId}-${index}`}
                          className="border-t border-slate-100 odd:bg-white even:bg-slate-50/50 hover:bg-rose-50/70"
                        >
                          <td className="whitespace-nowrap p-4 font-medium">
                            {item.customerId || "Missing"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.centralAmount.toLocaleString()}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.centralMerchantReference || "-"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.paidBillId || "Not found"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.paidAmount !== null
                              ? item.paidAmount.toLocaleString()
                              : "Not found"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.status || "-"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.occurrences.toLocaleString()}
                          </td>
                          <td className="min-w-64 p-4 text-red-700">
                            {item.matchNote}
                          </td>
                        </tr>
                      ))}
                      {result.differences.length === 0 && (
                        <tr>
                          <td
                            colSpan={8}
                            className="p-10 text-center font-medium text-green-700"
                          >
                            All payments are matched.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

          </>
        )}

      </section>

      {/* FOOTER */}
      <footer className="mt-12 border-t border-slate-200/70 bg-white/70 backdrop-blur">
        <div className="max-w-7xl mx-auto px-6 py-5 text-center text-sm text-slate-500">
          Central Payment Reconciliation System
        </div>
      </footer>

      </main>
    </div>
  );
}
