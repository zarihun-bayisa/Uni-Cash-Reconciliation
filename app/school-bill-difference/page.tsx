"use client";

import { useState } from "react";
import AppHeader from "../app-header";
import {
  useReconciliation,
  type ReconciliationResult,
} from "../reconciliation-context";

export default function SchoolBillDifferencePage() {
  const [centralFile, setCentralFile] = useState<File | null>(null);
  const [paidFile, setPaidFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloading, setDownloading] = useState<
    "differences" | "duplicates" | "matches" | null
  >(null);
  const [error, setError] = useState("");
  const { schoolResult, setSchoolResult } = useReconciliation();

  async function handleCompare() {
    if (!centralFile || !paidFile) {
      setError("Please select the school Central and Paid files.");
      return;
    }

    setLoading(true);
    setError("");
    setSchoolResult(null);

    try {
      const formData = new FormData();
      formData.append("centralFile", centralFile);
      formData.append("paidFile", paidFile);

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
    report: "differences" | "duplicates" | "matches"
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
    <div className="min-h-screen bg-slate-50 text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-7xl px-6 py-10">
          <div className="mb-8 flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
            <div>
              <p className="text-sm font-semibold uppercase tracking-[0.16em] text-blue-700">
                School Reconciliation
              </p>
              <h2 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">
                School Bill Difference
              </h2>
              <p className="mt-2 max-w-2xl text-slate-500">
                Compare school payment records by Student ID, Amount, and
                Merchant Reference. Amounts are compared using the whole-number
                portion only; digits after the decimal point are ignored.
              </p>
            </div>
          </div>

          <div className="grid gap-6 md:grid-cols-2">
            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <h3 className="mb-2 text-lg font-semibold">School Central File</h3>
              <p className="mb-5 text-sm text-slate-500">
                Upload the school Central payment export (XLSX or XLS).
              </p>
              <label className="flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed border-blue-300 p-8 transition hover:bg-blue-50">
                <span className="mb-3 text-4xl">📊</span>
                <span className="font-medium text-blue-800">
                  Select School Central File
                </span>
                <span className="mt-2 text-xs text-slate-400">
                  XLSX or XLS format
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(event) => {
                    setCentralFile(event.target.files?.[0] ?? null);
                    setSchoolResult(null);
                    setError("");
                  }}
                />
              </label>
              {centralFile && (
                <p className="mt-4 break-all text-sm text-green-700">
                  ✓ {centralFile.name}
                </p>
              )}
            </div>

            <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
              <h3 className="mb-2 text-lg font-semibold">School Paid File</h3>
              <p className="mb-5 text-sm text-slate-500">
                Upload the school billing export (CSV, XLSX, or XLS).
              </p>
              <label className="flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed border-emerald-300 p-8 transition hover:bg-emerald-50">
                <span className="mb-3 text-4xl">📁</span>
                <span className="font-medium text-emerald-800">
                  Select School Paid File
                </span>
                <span className="mt-2 text-xs text-slate-400">
                  CSV, XLSX, or XLS format
                </span>
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  className="hidden"
                  onChange={(event) => {
                    setPaidFile(event.target.files?.[0] ?? null);
                    setSchoolResult(null);
                    setError("");
                  }}
                />
              </label>
              {paidFile && (
                <p className="mt-4 break-all text-sm text-green-700">
                  ✓ {paidFile.name}
                </p>
              )}
            </div>
          </div>

          <div className="mt-8 text-center">
            <button
              type="button"
              onClick={handleCompare}
              disabled={
                !centralFile || !paidFile || loading || downloading !== null
              }
              className="rounded-xl bg-blue-700 px-10 py-4 font-semibold text-white shadow-lg shadow-blue-700/20 transition hover:bg-blue-800 disabled:cursor-not-allowed disabled:bg-slate-400 disabled:shadow-none"
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
              <div className="mt-10 grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-4">
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

              <div className="mt-10 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
                <div className="border-b border-slate-200 p-6">
                  <h3 className="text-xl font-bold">School Bill Differences</h3>
                  <p className="mt-1 text-sm text-slate-500">
                    Showing{" "}
                    {schoolResult.differences.length.toLocaleString()} records.
                  </p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 text-slate-600">
                      <tr>
                        <th className="whitespace-nowrap p-4">Student ID</th>
                        <th className="whitespace-nowrap p-4">Central Amount</th>
                        <th className="whitespace-nowrap p-4">
                          Central Merchant Reference
                        </th>
                        <th className="whitespace-nowrap p-4">Paid Reference</th>
                        <th className="whitespace-nowrap p-4">Paid Amount</th>
                        <th className="whitespace-nowrap p-4">Payment Status</th>
                        <th className="whitespace-nowrap p-4">Difference Note</th>
                      </tr>
                    </thead>
                    <tbody>
                      {schoolResult.differences.map((item, index) => (
                        <tr
                          key={`${item.customerId}-${index}`}
                          className="border-t border-slate-200 hover:bg-slate-50"
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
                          <td className="min-w-64 p-4 text-red-700">
                            {item.matchNote}
                          </td>
                        </tr>
                      ))}
                      {schoolResult.differences.length === 0 && (
                        <tr>
                          <td
                            colSpan={7}
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
    <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
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
      className="mt-4 rounded-lg bg-blue-700 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-800 disabled:cursor-not-allowed disabled:bg-slate-400"
    >
      {loading ? "Generating Report..." : `⬇ ${label}`}
    </button>
  );
}
