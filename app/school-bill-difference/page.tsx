"use client";

import { useState } from "react";
import AppHeader from "../app-header";
import {
  useReconciliation,
  type ReconciliationResult,
} from "../reconciliation-context";

export default function SchoolBillDifferencePage() {
  const [centralFiles, setCentralFiles] = useState<File[]>([]);
  const [paidFiles, setPaidFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [downloading, setDownloading] = useState<
    "differences" | "duplicates" | "unduplicated" | "matches" | null
  >(null);
  const [error, setError] = useState("");
  const { schoolResult, setSchoolResult } = useReconciliation();

  async function handleCompare() {
    if (centralFiles.length === 0 || paidFiles.length === 0) {
      setError("Please select the school Central and Paid files.");
      return;
    }

    setLoading(true);
    setError("");
    setSchoolResult(null);

    try {
      const formData = new FormData();
      centralFiles.forEach((file) => formData.append("centralFile", file));
      paidFiles.forEach((file) => formData.append("paidFile", file));

      const response = await fetch("/api/school-reconcile", {
        method: "POST",
        body: formData,
      });
      const data: unknown = await response.json();

      if (!response.ok) {
        const message =
          typeof data === "object" &&
          data !== null &&
          "error" in data &&
          typeof data.error === "string"
            ? data.error
            : "School reconciliation failed.";
        throw new Error(message);
      }

      setSchoolResult(data as ReconciliationResult);
    } catch (compareError) {
      setError(
        compareError instanceof Error
          ? compareError.message
          : "Something went wrong during school reconciliation."
      );
    } finally {
      setLoading(false);
    }
  }

  async function handleDownload(
    report: "differences" | "duplicates" | "unduplicated" | "matches"
  ) {
    if (!schoolResult) {
      setError("Compare the school files before downloading a report.");
      return;
    }

    const reportOptions = {
      differences: {
        endpoint: "/api/report",
        payload: { differences: schoolResult.differences },
        filename: "school-bill-difference.xlsx",
        empty: schoolResult.differences.length === 0,
        emptyMessage: "There are no school bill differences to download.",
        failureMessage: "Failed to generate the school difference report.",
      },
      duplicates: {
        endpoint: "/api/duplicate-report",
        payload: { duplicateRows: schoolResult.duplicateRows },
        filename: "school-duplicate-rows.xlsx",
        empty: schoolResult.duplicateRows.length === 0,
        emptyMessage: "There are no duplicate school rows to download.",
        failureMessage: "Failed to generate the school duplicate report.",
      },
      unduplicated: {
        endpoint: "/api/school-unduplicated-report",
        payload: { rows: schoolResult.unduplicatedRows ?? [] },
        filename: "school-unduplicated-rows.xlsx",
        empty: (schoolResult.unduplicatedRows ?? []).length === 0,
        emptyMessage: "There are no unduplicated school rows to download.",
        failureMessage: "Failed to generate the unduplicated school report.",
      },
      matches: {
        endpoint: "/api/match-report",
        payload: { matchedPayments: schoolResult.matchedPayments },
        filename: "school-matched-payments.xlsx",
        empty: schoolResult.matchedPayments.length === 0,
        emptyMessage: "There are no matched school payments to download.",
        failureMessage: "Failed to generate the school matched report.",
      },
    } as const;
    const selectedReport = reportOptions[report];

    if (selectedReport.empty) {
      setError(selectedReport.emptyMessage);
      return;
    }

    setDownloading(report);
    setError("");

    try {
      const response = await fetch(selectedReport.endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(selectedReport.payload),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || selectedReport.failureMessage);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = selectedReport.filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the school bill difference report."
      );
    } finally {
      setDownloading(null);
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50/70 text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8 sm:py-10">
          <div className="mb-8 flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
            <div className="relative w-full overflow-hidden rounded-[2rem] bg-gradient-to-br from-slate-950 via-blue-950 to-indigo-900 px-6 py-8 text-white shadow-2xl shadow-blue-950/15 sm:px-9 sm:py-10">
              <div aria-hidden="true" className="absolute -right-10 -top-28 h-80 w-80 rounded-full border-[40px] border-white/[0.04]" />
              <div className="relative">
                <p className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.18em] text-blue-100">
                  <span className="h-1.5 w-1.5 rounded-full bg-cyan-300" />
                  School reconciliation
                </p>
                <h2 className="mt-4 text-3xl font-bold tracking-tight sm:text-4xl">
                  School Bill Difference
                </h2>
                <p className="mt-3 max-w-2xl text-sm leading-6 text-blue-100 sm:text-base">
                  Compare school payments and review bill differences.
                </p>
              </div>
            </div>
          </div>

          <div className="grid gap-6 md:grid-cols-2">
            <div className="rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-lg shadow-slate-200/40 transition duration-200 hover:-translate-y-1 hover:border-blue-200 hover:shadow-xl sm:p-7">
              <div className="mb-5 flex items-center gap-4">
                <span className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-50 text-2xl ring-1 ring-blue-100">📊</span>
                <div>
                  <h3 className="text-lg font-bold tracking-tight text-slate-900">School Central File</h3>
                  <p className="mt-1 text-xs font-medium uppercase tracking-wider text-blue-700">Source payments</p>
                </div>
              </div>
              <p className="mb-5 text-sm text-slate-500">
                Upload the school Central payment export (XLSX or XLS).
              </p>
              <label className="flex min-h-40 cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-blue-200 bg-blue-50/50 p-7 text-center transition hover:border-blue-400 hover:bg-blue-50">
                <span className="font-medium text-blue-800">
                  Select School Central Files
                </span>
                <span className="mt-2 text-xs text-slate-400">
                  XLSX or XLS format
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    setCentralFiles(Array.from(event.target.files ?? []));
                    setSchoolResult(null);
                    setError("");
                  }}
                />
              </label>
              {centralFiles.length > 0 && (
                <div className="mt-4 flex items-center gap-2 break-all rounded-xl bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-800 ring-1 ring-emerald-100">
                  <span aria-hidden="true">✓</span>
                  <span>{centralFiles.map((file) => file.name).join(", ")}</span>
                </div>
              )}
            </div>

            <div className="rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-lg shadow-slate-200/40 transition duration-200 hover:-translate-y-1 hover:border-emerald-200 hover:shadow-xl sm:p-7">
              <div className="mb-5 flex items-center gap-4">
                <span className="flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-50 text-2xl ring-1 ring-emerald-100">📁</span>
                <div>
                  <h3 className="text-lg font-bold tracking-tight text-slate-900">School Paid File</h3>
                  <p className="mt-1 text-xs font-medium uppercase tracking-wider text-emerald-700">Billing export</p>
                </div>
              </div>
              <p className="mb-5 text-sm text-slate-500">
                Upload the school billing export (CSV, XLSX, or XLS).
              </p>
              <label className="flex min-h-40 cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-emerald-200 bg-emerald-50/50 p-7 text-center transition hover:border-emerald-400 hover:bg-emerald-50">
                <span className="font-medium text-emerald-800">
                  Select School Paid Files
                </span>
                <span className="mt-2 text-xs text-slate-400">
                  CSV, XLSX, or XLS format
                </span>
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    setPaidFiles(Array.from(event.target.files ?? []));
                    setSchoolResult(null);
                    setError("");
                  }}
                />
              </label>
              {paidFiles.length > 0 && (
                <div className="mt-4 flex items-center gap-2 break-all rounded-xl bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-800 ring-1 ring-emerald-100">
                  <span aria-hidden="true">✓</span>
                  <span>{paidFiles.map((file) => file.name).join(", ")}</span>
                </div>
              )}
            </div>
          </div>

          <div className="mt-8 text-center">
            <button
              type="button"
              onClick={handleCompare}
              disabled={
                centralFiles.length === 0 ||
                paidFiles.length === 0 ||
                loading ||
                downloading !== null
              }
              className="rounded-2xl bg-gradient-to-r from-blue-700 via-blue-700 to-indigo-700 px-10 py-4 font-semibold text-white shadow-xl shadow-blue-800/25 transition hover:-translate-y-0.5 hover:shadow-2xl disabled:cursor-not-allowed disabled:from-slate-400 disabled:to-slate-400 disabled:shadow-none"
            >
              {loading ? "Comparing School Payments..." : "Compare School Bills"}
            </button>
          </div>

          {error && (
            <div
              role="alert"
              className="mt-6 rounded-lg border border-red-200 bg-red-50 p-4 text-red-700"
            >
              {error}
            </div>
          )}

          {schoolResult && (
            <>
              <div className="mt-10 grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-5">
                <SummaryCard
                  label="Central Payments"
                  value={schoolResult.totalPayments}
                />
                <SummaryCard
                  label="Duplicates Removed"
                  value={schoolResult.duplicateCount}
                  color="text-amber-600"
                >
                  <DownloadButton
                    label="Download Duplicate Excel"
                    loading={downloading === "duplicates"}
                    disabled={
                      schoolResult.duplicateRows.length === 0 ||
                      downloading !== null ||
                      loading
                    }
                    onClick={() => handleDownload("duplicates")}
                  />
                </SummaryCard>
                <SummaryCard
                  label="Unduplicated Central Rows"
                  value={(schoolResult.unduplicatedRows ?? []).length}
                  color="text-blue-700"
                >
                  <DownloadButton
                    label="Download Unduplicated Excel"
                    loading={downloading === "unduplicated"}
                    disabled={
                      (schoolResult.unduplicatedRows ?? []).length === 0 ||
                      downloading !== null ||
                      loading
                    }
                    onClick={() => handleDownload("unduplicated")}
                  />
                </SummaryCard>
                <SummaryCard
                  label="Matched Payments"
                  value={schoolResult.matchedCount}
                  color="text-green-600"
                >
                  <DownloadButton
                    label="Download Matched Excel"
                    loading={downloading === "matches"}
                    disabled={
                      schoolResult.matchedPayments.length === 0 ||
                      downloading !== null ||
                      loading
                    }
                    onClick={() => handleDownload("matches")}
                  />
                </SummaryCard>
                <SummaryCard
                  label="Differences"
                  value={schoolResult.differenceCount}
                  color="text-red-600"
                >
                  <DownloadButton
                    label="Download Difference Excel"
                    loading={downloading === "differences"}
                    disabled={
                      schoolResult.differences.length === 0 ||
                      downloading !== null ||
                      loading
                    }
                    onClick={() => handleDownload("differences")}
                  />
                </SummaryCard>
              </div>

              <div className="mt-10 overflow-hidden rounded-3xl border border-slate-200/80 bg-white/95 shadow-xl shadow-slate-200/50">
                <div className="border-b border-slate-100 bg-gradient-to-r from-white to-blue-50/60 p-6 sm:p-7">
                  <h3 className="text-xl font-bold">School Bill Differences</h3>
                  <p className="mt-1 text-sm text-slate-500">
                    Showing{" "}
                    {schoolResult.differences.length.toLocaleString()} records.
                  </p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-100/80 text-xs uppercase tracking-wide text-slate-600">
                      <tr>
                        <th className="whitespace-nowrap p-4">Student ID</th>
                        <th className="whitespace-nowrap p-4">
                          Central Merchant Reference
                        </th>
                        <th className="whitespace-nowrap p-4">
                          School Paid Merchant Reference
                        </th>
                        <th className="whitespace-nowrap p-4">Difference Note</th>
                      </tr>
                    </thead>
                    <tbody>
                      {schoolResult.differences.map((item, index) => (
                        <tr
                          key={`${item.customerId}-${index}`}
                          className="border-t border-slate-100 odd:bg-white even:bg-slate-50/50 hover:bg-blue-50/70"
                        >
                          <td className="whitespace-nowrap p-4 font-medium">
                            {item.customerId || "Missing"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.centralMerchantReference || "-"}
                          </td>
                          <td className="whitespace-nowrap p-4">
                            {item.paidMerchantReference || "Not found"}
                          </td>
                          <td className="min-w-64 p-4 text-red-700">
                            {item.matchNote}
                          </td>
                        </tr>
                      ))}
                      {schoolResult.differences.length === 0 && (
                        <tr>
                          <td
                            colSpan={4}
                            className="p-10 text-center font-medium text-green-700"
                          >
                            ✓ All school payments are matched!
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </section>
      </main>
    </div>
  );
}

function SummaryCard({
  label,
  value,
  color = "text-slate-900",
  children,
}: {
  label: string;
  value: number;
  color?: string;
  children?: React.ReactNode;
}) {
  return (
    <div className="rounded-3xl border border-slate-200/80 bg-gradient-to-br from-white to-blue-50/40 p-6 shadow-lg shadow-slate-200/40 transition duration-200 hover:-translate-y-0.5 hover:shadow-xl">
      <p className="text-sm font-medium text-slate-500">{label}</p>
      <h3 className={`mt-2 text-3xl font-bold ${color}`}>
        {value.toLocaleString()}
      </h3>
      {children}
    </div>
  );
}

function DownloadButton({
  label,
  loading,
  disabled,
  onClick,
}: {
  label: string;
  loading: boolean;
  disabled: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="mt-4 rounded-xl bg-gradient-to-r from-blue-700 to-indigo-700 px-4 py-2.5 text-sm font-semibold text-white shadow-md shadow-blue-800/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:from-slate-400 disabled:to-slate-400 disabled:hover:translate-y-0"
    >
      {loading ? "Generating Report..." : `⬇ ${label}`}
    </button>
  );
}
