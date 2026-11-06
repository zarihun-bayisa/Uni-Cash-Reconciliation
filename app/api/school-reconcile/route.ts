import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";

type ExcelRow = Record<string, unknown>;

type PaymentDifference = {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidBillId: string;
  paidAmount: number | null;
  status: string;
  occurrences: number;
  matchNote: string;
};

type MatchedPayment = {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidBillId: string;
  paidAmount: number;
  status: string;
};

function normalize(value: unknown): string {
  return String(value ?? "").trim().toUpperCase();
}

function normalizeHeader(value: unknown): string {
  return normalize(value).replace(/[^A-Z0-9]/g, "");
}

function getField(row: ExcelRow, names: string[]): unknown {
  const keys = Object.keys(row);

  for (const name of names) {
    const key = keys.find(
      (candidate) => normalizeHeader(candidate) === normalizeHeader(name)
    );
    if (key !== undefined) return row[key];
  }

  return undefined;
}

function parseAmount(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;

  const amount = Number(String(value).replace(/,/g, "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(amount) ? amount : null;
}

function amountsMatch(left: number, right: number): boolean {
  return Math.trunc(left) === Math.trunc(right);
}

function readRows(buffer: Buffer): ExcelRow[] {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

  if (!firstSheet) {
    throw new Error("An uploaded file does not contain a worksheet.");
  }

  return XLSX.utils.sheet_to_json<ExcelRow>(firstSheet, {
    defval: "",
    raw: false,
  });
}

function paymentDifference(
  customerId: string,
  centralAmount: number,
  merchantReference: string,
  occurrences: number,
  matchNote: string,
  candidates: ExcelRow[] = []
): PaymentDifference {
  const billIds = [
    ...new Set(
      candidates
        .map((row) => normalize(getField(row, ["reference", "Bill ID"])))
        .filter(Boolean)
    ),
  ];
  const amounts = [
    ...new Set(
      candidates
        .map((row) => parseAmount(getField(row, ["totalAmount", "Amount"])))
        .filter((amount): amount is number => amount !== null)
    ),
  ];
  const statuses = [
    ...new Set(
      candidates
        .map((row) =>
          normalize(getField(row, ["paymentStatus", "Status"]))
        )
        .filter(Boolean)
    ),
  ];

  return {
    customerId,
    centralAmount,
    centralMerchantReference: merchantReference,
    paidBillId: billIds.length === 1 ? billIds[0] : billIds.join(", "),
    paidAmount: amounts.length === 1 ? amounts[0] : null,
    status: statuses.join(", "),
    occurrences,
    matchNote,
  };
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const centralFile = formData.get("centralFile");
    const paidFile = formData.get("paidFile");

    if (!(centralFile instanceof File) || !(paidFile instanceof File)) {
      return NextResponse.json(
        { error: "Please upload both school Central and Paid files." },
        { status: 400 }
      );
    }

    const centralExtension = centralFile.name.toLowerCase();
    const paidExtension = paidFile.name.toLowerCase();
    if (!/\.(xlsx|xls)$/.test(centralExtension)) {
      return NextResponse.json(
        { error: "The school Central file must be an XLSX or XLS workbook." },
        { status: 400 }
      );
    }
    if (!/\.(csv|xlsx|xls)$/.test(paidExtension)) {
      return NextResponse.json(
        { error: "The school Paid file must be a CSV, XLSX, or XLS workbook." },
        { status: 400 }
      );
    }

    const centralRows = readRows(
      Buffer.from(await centralFile.arrayBuffer())
    );
    const paidRows = readRows(Buffer.from(await paidFile.arrayBuffer()));

    const uniqueRows: ExcelRow[] = [];
    const duplicateRows: ExcelRow[] = [];
    const seenPayments = new Set<string>();

    for (const row of centralRows) {
      const customerId = normalize(
        getField(row, ["Customer Id", "Customer ID"])
      );
      const amount = parseAmount(getField(row, ["Amount"]));
      const merchantReference = normalize(
        getField(row, ["Merchant Reference"])
      );

      if (customerId && amount !== null) {
        const key = JSON.stringify([customerId, amount, merchantReference]);
        if (seenPayments.has(key)) {
          duplicateRows.push(row);
          continue;
        }
        seenPayments.add(key);
      }

      uniqueRows.push(row);
    }

    const paidByStudent = new Map<string, ExcelRow[]>();
    for (const row of paidRows) {
      const studentId = normalize(
        getField(row, ["studentId", "Student ID", "Customer Id"])
      );
      if (!studentId) continue;

      const records = paidByStudent.get(studentId) ?? [];
      records.push(row);
      paidByStudent.set(studentId, records);
    }

    const occurrences = new Map<string, number>();
    for (const row of uniqueRows) {
      const customerId = normalize(
        getField(row, ["Customer Id", "Customer ID"])
      );
      if (customerId) {
        occurrences.set(customerId, (occurrences.get(customerId) ?? 0) + 1);
      }
    }

    const matchedPayments: MatchedPayment[] = [];
    const differences: PaymentDifference[] = [];

    for (const row of uniqueRows) {
      const customerId = normalize(
        getField(row, ["Customer Id", "Customer ID"])
      );
      const centralAmount = parseAmount(getField(row, ["Amount"]));
      const merchantReference = normalize(
        getField(row, ["Merchant Reference"])
      );
      const rowOccurrences = occurrences.get(customerId) ?? 0;

      if (!customerId || centralAmount === null || !merchantReference) {
        differences.push(
          paymentDifference(
            customerId,
            centralAmount ?? 0,
            merchantReference,
            rowOccurrences,
            "Incomplete Central payment information"
          )
        );
        continue;
      }

      const candidates = paidByStudent.get(customerId) ?? [];
      if (candidates.length === 0) {
        differences.push(
          paymentDifference(
            customerId,
            centralAmount,
            merchantReference,
            rowOccurrences,
            "Student ID not found in School Paid file"
          )
        );
        continue;
      }

      const exactMatch = candidates.find((candidate) => {
        const paidReference = normalize(
          getField(candidate, ["merchantReference", "Merchant Reference"])
        );
        const paidAmount = parseAmount(
          getField(candidate, ["totalAmount", "Total Amount"])
        );
        return (
          paidReference === merchantReference &&
          paidAmount !== null &&
          amountsMatch(paidAmount, centralAmount)
        );
      });

      if (exactMatch) {
        const paidAmount = parseAmount(
          getField(exactMatch, ["totalAmount", "Total Amount"])
        );
        if (paidAmount !== null) {
          matchedPayments.push({
            customerId,
            centralAmount,
            centralMerchantReference: merchantReference,
            paidBillId: normalize(
              getField(exactMatch, ["reference", "Bill ID"])
            ),
            paidAmount,
            status: normalize(
              getField(exactMatch, ["paymentStatus", "Status"])
            ),
          });
          continue;
        }
      }

      let note = "Payment details do not match";
      if (candidates.length > 1) {
        note = "Multiple Paid records found; review required";
      } else {
        const candidate = candidates[0];
        const paidReference = normalize(
          getField(candidate, ["merchantReference", "Merchant Reference"])
        );
        const paidAmount = parseAmount(
          getField(candidate, ["totalAmount", "Total Amount"])
        );

        if (
          paidReference !== merchantReference &&
          (paidAmount === null || !amountsMatch(paidAmount, centralAmount))
        ) {
          note = "Merchant Reference and Amount do not match";
        } else if (paidReference !== merchantReference) {
          note = "Merchant Reference does not match";
        } else {
          note = "Central Amount and Paid Amount do not match";
        }
      }

      differences.push(
        paymentDifference(
          customerId,
          centralAmount,
          merchantReference,
          rowOccurrences,
          note,
          candidates
        )
      );
    }

    return NextResponse.json({
      success: true,
      totalPayments: uniqueRows.length,
      duplicateCount: duplicateRows.length,
      duplicateRows,
      matchedCount: matchedPayments.length,
      matchedPayments,
      differenceCount: differences.length,
      differences,
    });
  } catch (error) {
    console.error("School reconciliation error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "An unexpected school reconciliation error occurred.",
      },
      { status: 500 }
    );
  }
}
