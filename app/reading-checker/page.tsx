"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import AppHeader from "../app-header";

const REQUIRED_HEADERS = [
  "Id",
  "ReadingTimePeriod",
  "CustomerName",
  "ContractNumber",
  "CurrentReading",
  "PreviousReading",
  "Latitude",
  "Longitude",
  "MeterSerial",
  "Wereda",
  "HouseNumber",
  "PhoneNumber",
  "Kebele",
  "SubCity",
] as const;

type ReadingIssue = {
  sheet: string;
  rowNumber: number;
  reason: string;
  values: unknown[];
};

type ReadingSummary = {
  checkedCount: number;
  correctCount: number;
  issues: ReadingIssue[];
};

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function cellText(value: unknown): string {
  return String(value ?? "").trim();
}

function parseReading(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const text = cellText(value).replace(/,/g, "");
  if (!text) return null;

  const reading = Number(text);
  return Number.isFinite(reading) ? reading : null;
}

function validateWorkbook(buffer: ArrayBuffer): ReadingSummary {
  const workbook = XLSX.read(buffer, { type: "array" });
  if (workbook.SheetNames.length === 0) {
    throw new Error("The selected workbook does not contain any worksheets.");
  }

  const formatErrors: string[] = [];
  const issues: ReadingIssue[] = [];
  let checkedCount = 0;
  let correctCount = 0;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Could not read the "${sheetName}" worksheet.`);
    }

    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: true,
    });
    const headers = rows[0] ?? [];
    const normalizedHeaders = headers.map(normalizeHeader);
    const missingHeaders = REQUIRED_HEADERS.filter(
      (header) => !normalizedHeaders.includes(normalizeHeader(header))
    );
    const extraHeaders = headers.filter(
      (header) =>
        cellText(header) &&
        !REQUIRED_HEADERS.some(
          (expected) => normalizeHeader(expected) === normalizeHeader(header)
        )
    );
    const hasExpectedOrder =
      normalizedHeaders.length === REQUIRED_HEADERS.length &&
      REQUIRED_HEADERS.every(
        (header, index) =>
          normalizedHeaders[index] === normalizeHeader(header)
      );

    if (missingHeaders.length > 0 || extraHeaders.length > 0 || !hasExpectedOrder) {
      const reasons = [
        missingHeaders.length > 0
          ? `missing columns: ${missingHeaders.join(", ")}`
          : "",
        extraHeaders.length > 0
          ? `unexpected columns: ${extraHeaders.map(cellText).join(", ")}`
          : "",
        missingHeaders.length === 0 &&
        extraHeaders.length === 0 &&
        !hasExpectedOrder
          ? "columns are not in the required order"
          : "",
      ].filter(Boolean);
      formatErrors.push(`${sheetName}: ${reasons.join("; ")}`);
      continue;
    }

    rows.slice(1).forEach((row, index) => {
      if (!row.some((value) => cellText(value))) return;
      checkedCount += 1;

      const currentReading = parseReading(row[4]);
      const previousReading = parseReading(row[5]);
      const reasons: string[] = [];

      if (currentReading === null) {
        reasons.push("CurrentReading is missing or is not a number.");
      }
      if (previousReading === null) {
        reasons.push("PreviousReading is missing or is not a number.");
      }
      if (
        currentReading !== null &&
        previousReading !== null &&
        currentReading < previousReading
      ) {
        reasons.push(
          `CurrentReading (${cellText(row[4])}) is less than PreviousReading (${cellText(row[5])}).`
        );
      }

      if (reasons.length > 0) {
        issues.push({
          sheet: sheetName,
          rowNumber: index + 2,
          reason: reasons.join(" "),
          values: row,
        });
      } else {
        correctCount += 1;
      }
    });
  }

  if (formatErrors.length > 0) {
    throw new Error(
      `The workbook format does not match the required reading template:\n${formatErrors.join("\n")}`
    );
  }
  if (checkedCount === 0) {
    throw new Error("No reading rows were found below the column headings.");
  }

  return { checkedCount, correctCount, issues };
}

function copyToClipboard(rows: unknown[][]): Promise<void> {
  const text = rows
    .map((row) =>
      row
        .map((value) => cellText(value).replace(/[\t\r\n]/g, " "))
        .join("\t")
    )
    .join("\n");
  return navigator.clipboard.writeText(text);
}

export default function ReadingCheckerPage() {
  const [file, setFile] = useState<File | null>(null);
  const [summary, setSummary] = useState<ReadingSummary | null>(null);
  const [error, setError] = useState("");
  const [copyMessage, setCopyMessage] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleCheck() {
    if (!file) {
      setError("Choose an Excel workbook first.");
      return;
    }

    setLoading(true);
    setError("");
    setCopyMessage("");
    setSummary(null);

    try {
      setSummary(validateWorkbook(await file.arrayBuffer()));
    } catch (checkError) {
      setError(
        checkError instanceof Error
          ? checkError.message
          : "Could not check the selected workbook."
      );
    } finally {
      setLoading(false);
    }
  }

  async function handleCopy(rows: unknown[][], message: string) {
    setError("");
    setCopyMessage("");
    try {
      await copyToClipboard(rows);
      setCopyMessage(message);
    } catch {
      setError("Could not copy rows. Allow clipboard access and try again.");
    }
  }

  function handleDownload() {
    if (!summary || summary.issues.length === 0) {
      setError("There are no incorrect rows to download.");
      return;
    }

    try {
      const workbook = XLSX.utils.book_new();
      const rowsBySheet = new Map<string, ReadingIssue[]>();
      for (const issue of summary.issues) {
        const sheetIssues = rowsBySheet.get(issue.sheet) ?? [];
        sheetIssues.push(issue);
        rowsBySheet.set(issue.sheet, sheetIssues);
      }

      for (const [sheetName, issues] of rowsBySheet) {
        const worksheet = XLSX.utils.aoa_to_sheet([
          [...REQUIRED_HEADERS],
          ...issues.map((issue) => issue.values),
        ]);
        XLSX.utils.book_append_sheet(
          workbook,
          worksheet,
          sheetName.slice(0, 31)
        );
      }

      XLSX.writeFile(workbook, "incorrect-reading-rows.xlsx");
      setError("");
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the incorrect reading rows."
      );
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50/70 text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8 sm:py-10">
          <header className="mb-8 rounded-[2rem] bg-gradient-to-br from-slate-950 via-blue-950 to-indigo-900 px-6 py-8 text-white shadow-2xl shadow-blue-950/15 sm:px-9 sm:py-10">
            <p className="text-xs font-bold uppercase tracking-[0.16em] text-blue-200">
              Reading validation
            </p>
            <h2 className="mt-3 text-3xl font-bold tracking-tight sm:text-4xl">
              Reading Checker
            </h2>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-blue-100 sm:text-base">
              Check the workbook columns, then flag readings where the current
              value is lower than the previous value.
            </p>
          </header>

          <section className="rounded-3xl border border-slate-200/80 bg-white/95 p-5 shadow-xl shadow-slate-200/50 sm:p-7">
            <h3 className="text-lg font-bold tracking-tight text-slate-900">
              Upload a reading workbook
            </h3>
            <p className="mt-1 text-sm leading-6 text-slate-500">
              Each worksheet must use the 14 columns in the supplied template,
              in the same order. Readings are checked row by row; equal or
              higher current readings are accepted.
            </p>
            <div className="mt-5 flex flex-col gap-4 sm:flex-row sm:items-end">
              <label className="min-w-0 flex-1 text-sm font-medium text-slate-700">
                Excel file (.xlsx or .xls)
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 p-3 text-sm shadow-inner shadow-slate-100 file:mr-4 file:cursor-pointer file:rounded-lg file:border-0 file:bg-blue-700 file:px-3 file:py-2 file:font-semibold file:text-white hover:border-blue-300"
                  onChange={(event) => {
                    setFile(event.target.files?.[0] ?? null);
                    setSummary(null);
                    setError("");
                    setCopyMessage("");
                  }}
                />
              </label>
              <button
                type="button"
                onClick={handleCheck}
                disabled={!file || loading}
                className="rounded-xl bg-gradient-to-r from-blue-700 to-indigo-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-blue-700/20 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {loading ? "Checking..." : "Check workbook"}
              </button>
            </div>
            {error && (
              <p
                role="alert"
                className="mt-4 whitespace-pre-line rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700"
              >
                {error}
              </p>
            )}
          </section>

          {summary && (
            <section
              aria-live="polite"
              className="mt-6 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7"
            >
              <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h3 className="text-lg font-bold text-slate-900">
                    Reading check results
                  </h3>
                  <p className="mt-1 text-sm text-slate-600">
                    {summary.correctCount} of {summary.checkedCount} rows are
                    correct. {summary.issues.length} row
                    {summary.issues.length === 1 ? "" : "s"} need review.
                  </p>
                </div>
                {summary.issues.length > 0 && (
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        void handleCopy(
                          [
                            [...REQUIRED_HEADERS],
                            ...summary.issues.map((issue) => issue.values),
                          ],
                          "Incorrect rows copied."
                        )
                      }
                      className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-2.5 text-sm font-semibold text-blue-800 hover:bg-blue-100"
                    >
                      Copy incorrect rows
                    </button>
                    <button
                      type="button"
                      onClick={handleDownload}
                      className="rounded-xl bg-blue-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-800"
                    >
                      Download incorrect rows
                    </button>
                  </div>
                )}
              </div>

              {copyMessage && (
                <p role="status" className="mt-3 text-sm font-medium text-emerald-700">
                  {copyMessage}
                </p>
              )}

              {summary.issues.length === 0 ? (
                <p className="mt-5 rounded-xl bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-800">
                  All reading rows are correct.
                </p>
              ) : (
                <div className="mt-5 overflow-x-auto rounded-xl border border-slate-200">
                  <table className="min-w-full border-collapse text-left text-sm">
                    <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
                      <tr>
                        <th className="whitespace-nowrap px-3 py-3">Sheet</th>
                        <th className="whitespace-nowrap px-3 py-3">Excel row</th>
                        <th className="whitespace-nowrap px-3 py-3">Reason</th>
                        {REQUIRED_HEADERS.map((header) => (
                          <th key={header} className="whitespace-nowrap px-3 py-3">
                            {header}
                          </th>
                        ))}
                        <th className="whitespace-nowrap px-3 py-3">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {summary.issues.map((issue) => (
                        <tr
                          key={`${issue.sheet}-${issue.rowNumber}`}
                          className="align-top text-slate-700"
                        >
                          <td className="whitespace-nowrap px-3 py-3">{issue.sheet}</td>
                          <td className="px-3 py-3">{issue.rowNumber}</td>
                          <td className="min-w-64 px-3 py-3 text-red-700">
                            {issue.reason}
                          </td>
                          {REQUIRED_HEADERS.map((header, index) => (
                            <td key={header} className="whitespace-nowrap px-3 py-3">
                              {cellText(issue.values[index])}
                            </td>
                          ))}
                          <td className="px-3 py-3">
                            <button
                              type="button"
                              onClick={() =>
                                void handleCopy(
                                  [issue.values],
                                  `Row ${issue.rowNumber} copied.`
                                )
                              }
                              className="whitespace-nowrap rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
                            >
                              Copy row
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          )}
        </section>
      </main>
    </div>
  );
}
