"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import AppHeader from "../app-header";

const RENT_FIELDS = [
  "customerCode",
  "name",
  "houseCode",
  "houseNumber",
  "amount",
  "contractDate",
  "phoneNumber",
  "city",
  "propertyType",
  "kebele",
] as const;

type RentField = (typeof RENT_FIELDS)[number];
const REQUIRED_RENT_FIELDS: RentField[] = [
  "customerCode",
  "name",
  "houseCode",
  "houseNumber",
  "amount",
  "contractDate",
];
type RentValues = Record<RentField, string>;
type SourceColumns = {
  customerCode: number | null;
  englishName: number | null;
  amharicName: number | null;
  fullName: number | null;
  firstName: number | null;
  secondName: number | null;
  lastName: number | null;
  houseCode: number | null;
  houseNumber: number | null;
  amount: number | null;
  contractDate: number | null;
  phoneNumber: number | null;
  kebele: number | null;
};

type RentRecord = {
  values: RentValues;
  source: string;
  issues: string[];
};

type OutputTemplate = {
  headers: unknown[];
  headerRow: number;
  sheetName: string;
};

const FIELD_LABELS: Record<RentField, string> = {
  customerCode: "Customer Code",
  name: "Customer Name",
  houseCode: "House Code",
  houseNumber: "House Number",
  amount: "Payment Amount",
  contractDate: "Contract Date",
  phoneNumber: "Phone Number",
  city: "City",
  propertyType: "R/C",
  kebele: "Kebele",
};

const FIELD_ALIASES: Record<RentField, string[]> = {
  customerCode: [
    "customer code",
    "customer id",
    "customer number",
    "customer no",
    "client code",
    "tenant code",
    "tenant id",
    "account code",
    "account number",
    "code",
    "id",
  ],
  name: [
    "customer name",
    "customer english name",
    "name eng",
    "name english",
    "tenant name",
    "renter name",
    "english name",
    "english",
    "name in english",
    "full name",
    "name",
    "first name",
    "ferst name",
    "given name",
  ],
  houseCode: [
    "house code",
    "property code",
    "building code",
    "unit code",
    "house id",
  ],
  houseNumber: [
    "house number",
    "house no",
    "house num",
    "house #",
    "shop number",
    "shop no",
    "room number",
    "room no",
    "unit number",
    "unit no",
    "property number",
    "apartment number",
  ],
  amount: [
    "payment amount",
    "paid amount",
    "amount paid",
    "rent amount",
    "rental amount",
    "monthly rent",
    "payment",
    "payement",
    "payemnt",
    "amount",
    "rent",
  ],
  contractDate: [
    "contract date",
    "date of contract",
    "agreement date",
    "lease date",
    "contract start date",
    "date",
  ],
  phoneNumber: [
    "phone",
    "phone number",
    "phone no",
    "phonenum",
    "telephone",
    "mobile",
    "mobile number",
  ],
  city: ["city", "city name"],
  propertyType: [
    "r/c",
    "r or c",
    "rc",
    "residential/commercial",
    "property type",
  ],
  kebele: ["kebele", "kebele name"],
};

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function cellText(value: unknown): string {
  return String(value ?? "").trim();
}

function sourceCellText(row: unknown[], column: number | null): string {
  return column === null ? "" : cellText(row[column]);
}

function incrementCode(startValue: string, increment: number): string {
  const value = startValue.trim();
  const match = /^(.*?)(\d+)$/.exec(value);
  if (!match) return value;

  const nextNumber = String(Number(match[2]) + increment).padStart(
    match[2].length,
    "0"
  );
  return `${match[1]}${nextNumber}`;
}

function findSourceColumns(header: unknown[]): SourceColumns {
  const columns: SourceColumns = {
    customerCode: null,
    englishName: null,
    amharicName: null,
    fullName: null,
    firstName: null,
    secondName: null,
    lastName: null,
    houseCode: null,
    houseNumber: null,
    amount: null,
    contractDate: null,
    phoneNumber: null,
    kebele: null,
  };
  const seen = new Set<number>();

  header.forEach((value, index) => {
    const name = normalizeHeader(value);
    if (!name) return;

    if (
      columns.houseCode === null &&
      FIELD_ALIASES.houseCode.some((alias) => normalizeHeader(alias) === name)
    ) {
      columns.houseCode = index;
      seen.add(index);
    }
    if (
      columns.houseNumber === null &&
      FIELD_ALIASES.houseNumber.some(
        (alias) => normalizeHeader(alias) === name
      )
    ) {
      columns.houseNumber = index;
      seen.add(index);
    }
    if (
      columns.amount === null &&
      FIELD_ALIASES.amount.some((alias) => normalizeHeader(alias) === name)
    ) {
      columns.amount = index;
      seen.add(index);
    }
    if (
      columns.contractDate === null &&
      FIELD_ALIASES.contractDate.some(
        (alias) => normalizeHeader(alias) === name
      )
    ) {
      columns.contractDate = index;
      seen.add(index);
    }
    if (
      columns.phoneNumber === null &&
      FIELD_ALIASES.phoneNumber.some(
        (alias) => normalizeHeader(alias) === name
      )
    ) {
      columns.phoneNumber = index;
      seen.add(index);
    }
    if (
      columns.kebele === null &&
      FIELD_ALIASES.kebele.some((alias) => normalizeHeader(alias) === name)
    ) {
      columns.kebele = index;
      seen.add(index);
    }
    if (
      columns.customerCode === null &&
      FIELD_ALIASES.customerCode.some(
        (alias) => normalizeHeader(alias) === name
      ) &&
      !seen.has(index)
    ) {
      columns.customerCode = index;
      seen.add(index);
    }

    if (
      name === "nameeng" ||
      name === "englishname" ||
      name === "nameinenglish" ||
      name === "customerenglishname"
    ) {
      columns.englishName ??= index;
    } else if (
      name === "nameamh" ||
      name === "amharicname" ||
      name === "nameinamharic" ||
      name === "customeramharicname"
    ) {
      columns.amharicName ??= index;
    } else if (name === "secondname" || name === "middlename" || name === "fathername") {
      columns.secondName ??= index;
    } else if (name === "lastname" || name === "surname" || name === "familyname") {
      columns.lastName ??= index;
    } else if (
      name === "firstname" ||
      name === "ferstname" ||
      name === "givenname"
    ) {
      columns.firstName ??= index;
    } else if (
      columns.fullName === null &&
      FIELD_ALIASES.name.some((alias) => normalizeHeader(alias) === name)
    ) {
      columns.fullName = index;
    }
  });

  return columns;
}

function findHeaderRow(
  rows: unknown[][]
): { index: number; columns: SourceColumns } | null {
  for (let index = 0; index < Math.min(rows.length, 25); index += 1) {
    const columns = findSourceColumns(rows[index] ?? []);
    const hasName =
      columns.englishName !== null ||
      columns.amharicName !== null ||
      columns.fullName !== null ||
      columns.firstName !== null ||
      columns.secondName !== null ||
      columns.lastName !== null;
    const recognizedFields = [
      columns.customerCode,
      columns.houseCode,
      columns.houseNumber,
      columns.amount,
      columns.contractDate,
      columns.phoneNumber,
      columns.kebele,
    ].filter((column) => column !== null).length;
    if (hasName && recognizedFields > 0) return { index, columns };
  }
  return null;
}

function readName(row: unknown[], columns: SourceColumns): string {
  for (const column of [
    columns.englishName,
    columns.fullName,
    columns.amharicName,
  ]) {
    if (column === null) continue;
    const fullName = cellText(row[column]);
    if (fullName) return fullName;
  }

  return [
    columns.firstName,
    columns.secondName,
    columns.lastName,
  ]
    .map((column) => (column === null ? "" : cellText(row[column])))
    .filter(Boolean)
    .join(" ");
}

function findTargetColumn(headers: unknown[], field: RentField): number {
  const aliases = FIELD_ALIASES[field].map(normalizeHeader);
  return headers.findIndex((header) => aliases.includes(normalizeHeader(header)));
}

function findOutputTemplate(
  workbook: XLSX.WorkBook,
  preferredSheet?: string
): OutputTemplate {
  const sheetName =
    preferredSheet && workbook.SheetNames.includes(preferredSheet)
      ? preferredSheet
      : workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("The selected format workbook has no worksheets.");
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Could not read the "${sheetName}" worksheet.`);

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  for (let index = 0; index < Math.min(rows.length, 25); index += 1) {
    const headers = rows[index] ?? [];
    const mappedFields = RENT_FIELDS.filter(
      (field) => findTargetColumn(headers, field) >= 0
    );
    if (
      mappedFields.includes("name") &&
      mappedFields.includes("customerCode") &&
      mappedFields.includes("houseCode") &&
      mappedFields.includes("houseNumber") &&
      mappedFields.includes("amount") &&
      mappedFields.includes("contractDate")
    ) {
      return {
        headers,
        headerRow: index,
        sheetName,
      };
    }
  }

  throw new Error(
    "The format workbook must have columns for customer code, customer name, house code, house number, payment amount, and contract date."
  );
}

async function loadOutputTemplate(): Promise<ArrayBuffer> {
  const response = await fetch("/UniRentNewDataFormat.xlsx");
  if (!response.ok) {
    throw new Error("Could not load the UniRent output format workbook.");
  }
  return response.arrayBuffer();
}

function makeRecord(
  row: unknown[],
  columns: SourceColumns,
  source: string,
  generatedCode: string,
  defaultAmount: string,
  defaultContractDate: string,
  city: string,
  propertyType: string,
  defaultKebele: string,
  requiredFields: RentField[]
): RentRecord | null {
  const name = readName(row, columns);
  const customerCode = sourceCellText(row, columns.customerCode);
  const houseCode = sourceCellText(row, columns.houseCode);
  const houseNumber = sourceCellText(row, columns.houseNumber);
  const amountFromFile = sourceCellText(row, columns.amount);
  const contractDateFromFile = sourceCellText(row, columns.contractDate);
  const phoneNumberFromFile = sourceCellText(row, columns.phoneNumber);
  const kebeleFromFile = sourceCellText(row, columns.kebele);

  if (
    ![
      name,
      customerCode,
      houseCode,
      houseNumber,
      amountFromFile,
      contractDateFromFile,
    ].some(Boolean)
  ) {
    return null;
  }

  const resolvedCode = customerCode || generatedCode.trim();
  const values: RentValues = {
    customerCode: resolvedCode,
    name,
    houseCode: resolvedCode,
    houseNumber: resolvedCode,
    amount: amountFromFile || defaultAmount.trim(),
    contractDate: contractDateFromFile || defaultContractDate.trim(),
    phoneNumber: phoneNumberFromFile || "09",
    city: city.trim(),
    propertyType,
    kebele: kebeleFromFile || defaultKebele.trim(),
  };
  const issues = requiredFields.filter((field) => !values[field]).map(
    (field) => `${FIELD_LABELS[field]} is missing`
  );
  return { values, source, issues };
}

function updateRecordValue(
  records: RentRecord[],
  index: number,
  field: RentField,
  value: string,
  requiredFields: RentField[]
): RentRecord[] {
  return records.map((record, recordIndex) => {
    if (recordIndex !== index) return record;
    const values =
      field === "customerCode" ||
      field === "houseCode" ||
      field === "houseNumber"
        ? {
            ...record.values,
            customerCode: value,
            houseCode: value,
            houseNumber: value,
          }
        : { ...record.values, [field]: value };
    return {
      ...record,
      values,
      issues: requiredFields.filter((candidate) => !values[candidate]).map(
        (candidate) => `${FIELD_LABELS[candidate]} is missing`
      ),
    };
  });
}

export default function RentImportPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [defaultCustomerCode, setDefaultCustomerCode] = useState("");
  const [defaultAmount, setDefaultAmount] = useState("");
  const [defaultContractDate, setDefaultContractDate] = useState("");
  const [city, setCity] = useState("");
  const [propertyType, setPropertyType] = useState("");
  const [defaultKebele, setDefaultKebele] = useState("");
  const [records, setRecords] = useState<RentRecord[]>([]);
  const [template, setTemplate] = useState<OutputTemplate | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleConvert() {
    if (files.length === 0) {
      setError("Choose at least one rent workbook first.");
      return;
    }
    if (
      defaultAmount.trim() &&
      (!Number.isFinite(Number(defaultAmount)) || Number(defaultAmount) < 0)
    ) {
      setError("Enter a valid non-negative default payment amount.");
      return;
    }

    setLoading(true);
    setError("");
    setRecords([]);
    setTemplate(null);

    try {
      if (
        defaultCustomerCode.trim() &&
        !/\d+$/.test(defaultCustomerCode.trim())
      ) {
        throw new Error(
          "The starting customer code must end in digits, for example AA/001."
        );
      }

      const outputTemplate = findOutputTemplate(
        XLSX.read(await loadOutputTemplate(), { type: "array" })
      );
      const requiredFields = REQUIRED_RENT_FIELDS.filter(
        (field) => findTargetColumn(outputTemplate.headers, field) >= 0
      );
      const convertedRecords: RentRecord[] = [];
      let customerCodeSequence = 0;

      for (const file of files) {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        if (workbook.SheetNames.length === 0) {
          throw new Error(`${file.name} does not contain any worksheets.`);
        }

        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) {
            throw new Error(`Could not read "${sheetName}" in ${file.name}.`);
          }
          const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
            header: 1,
            defval: "",
            raw: false,
          });
          const header = findHeaderRow(rows);
          if (!header) continue;

          rows.slice(header.index + 1).forEach((row, rowIndex) => {
            const hasCustomerCode = sourceCellText(
              row,
              header.columns.customerCode
            );
            const record = makeRecord(
              row,
              header.columns,
              `${file.name} / ${sheetName} / row ${header.index + rowIndex + 2}`,
              incrementCode(defaultCustomerCode, customerCodeSequence),
              defaultAmount,
              defaultContractDate,
              city,
              propertyType,
              defaultKebele,
              requiredFields
            );
            if (record) {
              if (!hasCustomerCode && defaultCustomerCode.trim()) {
                customerCodeSequence += 1;
              }
              convertedRecords.push(record);
            }
          });
        }
      }

      if (convertedRecords.length === 0) {
        throw new Error(
          "No rent records were found. Include a name column and at least one other recognized rent field."
        );
      }

      setRecords(convertedRecords);
      setTemplate(outputTemplate);
    } catch (convertError) {
      setError(
        convertError instanceof Error
          ? convertError.message
          : "Could not read the selected rent workbook."
      );
    } finally {
      setLoading(false);
    }
  }

  async function handleDownload() {
    if (records.length === 0) {
      setError("Convert rent records before downloading.");
      return;
    }
    if (!template) {
      setError("Convert rent records before downloading.");
      return;
    }
    try {
      const headers = template.headers;
      const outputWorkbook = XLSX.read(await loadOutputTemplate(), {
        type: "array",
        cellStyles: true,
      });
      const worksheet = outputWorkbook.Sheets[template.sheetName];
      if (!worksheet) {
        throw new Error(
          `Could not read the "${template.sheetName}" output worksheet.`
        );
      }

      const existingRange = worksheet["!ref"]
        ? XLSX.utils.decode_range(worksheet["!ref"])
        : {
            s: { r: 0, c: 0 },
            e: { r: template.headerRow, c: headers.length - 1 },
          };
      for (
        let rowIndex = template.headerRow + 1;
        rowIndex <= existingRange.e.r;
        rowIndex += 1
      ) {
        for (
          let columnIndex = existingRange.s.c;
          columnIndex <= existingRange.e.c;
          columnIndex += 1
        ) {
          const address = XLSX.utils.encode_cell({
            r: rowIndex,
            c: columnIndex,
          });
          const cell = worksheet[address];
          if (cell) {
            delete cell.v;
            delete cell.w;
            delete cell.t;
            delete cell.f;
          }
        }
      }

      records.forEach((record, recordIndex) => {
        for (const field of RENT_FIELDS) {
          const column = findTargetColumn(headers, field);
          if (column < 0) continue;
          const address = XLSX.utils.encode_cell({
            r: template.headerRow + recordIndex + 1,
            c: column,
          });
          const previousCell = worksheet[address];
          const value = record.values[field];
          const number =
            field === "amount" ? Number(value.replace(/,/g, "")) : NaN;
          worksheet[address] = {
            ...previousCell,
            t: Number.isFinite(number) ? "n" : "s",
            v: Number.isFinite(number) ? number : value,
          };
        }
      });
      existingRange.e.r = Math.max(
        existingRange.e.r,
        template.headerRow + records.length
      );
      existingRange.e.c = Math.max(existingRange.e.c, headers.length - 1);
      worksheet["!ref"] = XLSX.utils.encode_range(existingRange);
      XLSX.writeFile(outputWorkbook, "converted-rent-data.xlsx");
      setError("");
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the converted rent workbook."
      );
    }
  }
  
  const missingCount = records.filter((record) => record.issues.length > 0).length;
  const tableHeaders =
    template?.headers ?? RENT_FIELDS.map((field) => FIELD_LABELS[field]);
  const requiredFields = template
    ? REQUIRED_RENT_FIELDS.filter(
        (field) => findTargetColumn(template.headers, field) >= 0
      )
    : [];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-violet-50/60 text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8 sm:py-10">
          <header className="mb-8 rounded-[2rem] bg-gradient-to-br from-slate-950 via-indigo-950 to-violet-900 px-6 py-8 text-white shadow-2xl shadow-indigo-950/15 sm:px-9 sm:py-10">
            <p className="text-xs font-bold uppercase tracking-[0.16em] text-violet-200">
              Rent records
            </p>
            <h2 className="mt-3 text-3xl font-bold tracking-tight sm:text-4xl">
              Rent Format
            </h2>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-violet-100 sm:text-base">
              Convert rent spreadsheets into the UniRent import layout. Names
              from multiple columns are combined into one name field; source
              values are kept when present.
            </p>
          </header>

          <section className="rounded-3xl border border-slate-200/80 bg-white/95 p-5 shadow-xl shadow-slate-200/50 sm:p-7">
            <h3 className="text-lg font-bold tracking-tight text-slate-900">
              Upload rent workbooks
            </h3>
            <p className="mt-1 text-sm leading-6 text-slate-500">
              Missing customer codes, house details, payment amounts, and
              contract dates can be entered in the preview before export.
            </p>
            <div className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-4 lg:items-end">
              <label className="text-sm font-medium text-slate-700">
                Rent Excel files (.xlsx or .xls)
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 p-3 text-sm shadow-inner shadow-slate-100 file:mr-3 file:cursor-pointer file:rounded-lg file:border-0 file:bg-violet-700 file:px-3 file:py-2 file:font-semibold file:text-white hover:border-violet-300"
                  onChange={(event) => {
                    setFiles(Array.from(event.target.files ?? []));
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                />
              </label>
              <label className="text-sm font-medium text-slate-700">
                Default customer code
                <input
                  type="text"
                  value={defaultCustomerCode}
                  placeholder="AA/001"
                  onChange={(event) => {
                    setDefaultCustomerCode(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-label="Default customer code"
                />
                <span className="mt-1 block text-xs font-normal text-slate-500">
                  Customer code, house code, and house number use the same
                  value. Generated codes increment when the source code is blank.
                </span>
              </label>
              <label className="text-sm font-medium text-slate-700">
                Default payment amount
                <input
                  type="number"
                  min="0"
                  step="any"
                  value={defaultAmount}
                  onChange={(event) => {
                    setDefaultAmount(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-describedby="rent-amount-help"
                />
                <span id="rent-amount-help" className="mt-1 block text-xs font-normal text-slate-500">
                  Used only where a source row has no amount.
                </span>
              </label>
              <label className="text-sm font-medium text-slate-700">
                Default contract date
                <input
                  type="text"
                  value={defaultContractDate}
                  onChange={(event) => {
                    setDefaultContractDate(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  placeholder="YYYY-MM-DD"
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-describedby="rent-date-help"
                />
                <span id="rent-date-help" className="mt-1 block text-xs font-normal text-slate-500">
                  Used only where a source row has no contract date.
                </span>
              </label>
              <label className="text-sm font-medium text-slate-700">
                City
                <input
                  type="text"
                  value={city}
                  onChange={(event) => {
                    setCity(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-label="City name"
                />
              </label>
              <label className="text-sm font-medium text-slate-700">
                R/C
                <select
                  value={propertyType}
                  onChange={(event) => {
                    setPropertyType(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-label="Residential or commercial"
                >
                  <option value="">Select R or C</option>
                  <option value="R">R</option>
                  <option value="C">C</option>
                </select>
              </label>
              <label className="text-sm font-medium text-slate-700">
                Default kebele
                <input
                  type="text"
                  value={defaultKebele}
                  onChange={(event) => {
                    setDefaultKebele(event.target.value);
                    setRecords([]);
                    setTemplate(null);
                    setError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-violet-400 focus:bg-white focus:outline-none"
                  aria-label="Default kebele"
                />
                <span className="mt-1 block text-xs font-normal text-slate-500">
                  A kebele value from the rent workbook takes priority.
                </span>
              </label>
              <button
                type="button"
                onClick={handleConvert}
                disabled={files.length === 0 || loading}
                className="rounded-xl bg-gradient-to-r from-violet-700 to-indigo-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-violet-700/20 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50 sm:col-span-2 lg:col-span-4 lg:justify-self-start"
              >
                {loading ? "Converting..." : "Convert and preview"}
              </button>
              <p className="text-sm text-slate-500 sm:col-span-2 lg:col-span-4">
                {files.length === 0
                  ? "No rent workbooks selected."
                  : `${files.length} file${files.length === 1 ? "" : "s"} selected: ${files.map((file) => file.name).join(", ")}`}
              </p>
            </div>
            {error && (
              <p role="alert" className="mt-4 rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700">
                {error}
              </p>
            )}
          </section>

          {records.length > 0 && (
            <section className="mt-6 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
              <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
                <div>
                  <h3 className="text-lg font-bold text-slate-900">
                    Review rent records
                  </h3>
                  <p className="mt-1 text-sm text-slate-600">
                    {records.length} records · {missingCount} need attention.
                    Output follows UniRentNewDataFormat.xlsx.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => void handleDownload()}
                  className="shrink-0 rounded-xl bg-gradient-to-r from-emerald-700 to-teal-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-emerald-700/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Download Rent Excel ({records.length})
                </button>
              </div>
              {missingCount > 0 && (
                <p className="mt-4 rounded-xl bg-amber-50 px-4 py-3 text-sm text-amber-800">
                  {missingCount} record{missingCount === 1 ? " has" : "s have"} blank required fields. Blank values will remain blank in
                  the downloaded workbook; you can still edit the highlighted
                  cells before exporting.
                </p>
              )}
              <div className="mt-5 max-h-[36rem] overflow-auto rounded-2xl border border-slate-200">
                <table className="min-w-full divide-y divide-slate-200 text-left text-sm">
                  <thead className="sticky top-0 bg-slate-100 text-xs uppercase tracking-wide text-slate-600">
                    <tr>
                      {tableHeaders.map((header, index) => (
                        <th key={`${cellText(header)}-${index}`} className="whitespace-nowrap px-4 py-3 font-semibold">
                          {cellText(header)}
                        </th>
                      ))}
                      <th className="whitespace-nowrap px-4 py-3 font-semibold">Source</th>
                      <th className="min-w-56 px-4 py-3 font-semibold">Review notes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {records.map((record, recordIndex) => (
                      <tr
                        key={`${record.source}-${recordIndex}`}
                        className={record.issues.length ? "bg-amber-50/60" : "bg-white"}
                      >
                        {tableHeaders.map((_, columnIndex) => {
                          const field = RENT_FIELDS.find(
                            (candidate) =>
                              findTargetColumn(tableHeaders, candidate) ===
                              columnIndex
                          );
                          return (
                            <td key={columnIndex} className="min-w-40 px-3 py-2">
                              {field ? (
                                <>
                                  {field === "propertyType" ? (
                                    <select
                                      aria-label={`${FIELD_LABELS[field]} for record ${recordIndex + 1}`}
                                      value={record.values[field]}
                                      onChange={(event) =>
                                        setRecords((current) =>
                                          updateRecordValue(
                                            current,
                                            recordIndex,
                                            field,
                                            event.target.value,
                                            requiredFields
                                          )
                                        )
                                      }
                                      className={`w-full rounded-lg border px-3 py-2 text-sm focus:outline-none ${
                                        !record.values[field]
                                          ? "border-amber-300 bg-amber-50 focus:border-amber-500"
                                          : "border-slate-200 bg-white focus:border-violet-400"
                                      }`}
                                    >
                                      <option value="">Select R or C</option>
                                      <option value="R">R</option>
                                      <option value="C">C</option>
                                    </select>
                                  ) : (
                                    <input
                                      type="text"
                                      aria-label={`${FIELD_LABELS[field]} for record ${recordIndex + 1}`}
                                      value={record.values[field]}
                                      onChange={(event) =>
                                        setRecords((current) =>
                                          updateRecordValue(
                                            current,
                                            recordIndex,
                                            field,
                                            event.target.value,
                                            requiredFields
                                          )
                                        )
                                      }
                                      className={`w-full rounded-lg border px-3 py-2 text-sm focus:outline-none ${
                                        !record.values[field]
                                          ? "border-amber-300 bg-amber-50 focus:border-amber-500"
                                          : "border-slate-200 bg-white focus:border-violet-400"
                                      }`}
                                    />
                                  )}
                                </>
                              ) : (
                                <span className="whitespace-nowrap text-slate-400">—</span>
                              )}
                            </td>
                          );
                        })}
                        <td className="whitespace-nowrap px-4 py-3 text-slate-600">
                          {record.source}
                        </td>
                        <td className="min-w-56 px-4 py-3 text-amber-800">
                          {record.issues.join(" · ") || "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="mt-3 text-xs leading-5 text-slate-500">
                Names from first, second, and last-name source columns are
                combined into one customer name. Text is preserved as entered,
                including Amharic names, in the name column.
              </p>
            </section>
          )}
        </section>
      </main>
    </div>
  );
}
