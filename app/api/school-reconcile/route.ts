import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";

type ExcelRow = Record<string, unknown>;

type PaymentDifference = {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidMerchantReference: string;
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
  paidMerchantReference: string;
  paidBillId: string;
  paidAmount: number | null;
  status: string;
};

function normalize(value: unknown): string {
  return String(value ?? "").trim().toUpperCase();
}

function preserveText(value: unknown): string {
  return String(value ?? "").trim();
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
        .map((row) =>
          normalize(getField(row, ["reference", "Bill ID", "Bills", "Bill"]))
        )
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
    paidMerchantReference: [
      ...new Set(
        candidates
          .map((row) =>
            preserveText(getField(row, ["merchantReference", "Merchant Reference"]))
          )
          .filter(Boolean)
      ),
    ].join(", "),
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
    const centralEntries = formData.getAll("centralFile");
    const paidEntries = formData.getAll("paidFile");
    const centralFiles = centralEntries.filter(
      (entry): entry is File => entry instanceof File
    );
    const paidFiles = paidEntries.filter(
      (entry): entry is File => entry instanceof File
    );

    if (
      centralFiles.length === 0 ||
      paidFiles.length === 0 ||
      centralFiles.length !== centralEntries.length ||
      paidFiles.length !== paidEntries.length
    ) {
      return NextResponse.json(
        { error: "Please upload both school Central and Paid files." },
        { status: 400 }
      );
    }

    if (centralFiles.some((file) => !/\.(xlsx|xls)$/i.test(file.name))) {
      return NextResponse.json(
        { error: "All school Central files must be XLSX or XLS workbooks." },
        { status: 400 }
      );
    }
    if (paidFiles.some((file) => !/\.(csv|xlsx|xls)$/i.test(file.name))) {
      return NextResponse.json(
        { error: "All school Paid files must be CSV, XLSX, or XLS files." },
        { status: 400 }
      );
    }

    const centralRows = (
      await Promise.all(
        centralFiles.map(async (file) =>
          readRows(Buffer.from(await file.arrayBuffer()))
        )
      )
    ).flat();
    const paidRows = (
      await Promise.all(
        paidFiles.map(async (file) =>
          readRows(Buffer.from(await file.arrayBuffer()))
        )
      )
    ).flat();

    const uniqueRows: ExcelRow[] = [];
    const duplicateRows: ExcelRow[] = [];
    const seenMerchantReferences = new Set<string>();

    for (const row of centralRows) {
      const merchantReference = normalize(
        getField(row, ["Merchant Reference"])
      );

      if (merchantReference) {
        if (seenMerchantReferences.has(merchantReference)) {
          duplicateRows.push(row);
          continue;
        }
        seenMerchantReferences.add(merchantReference);
      }

      uniqueRows.push(row);
    }

    const paidByMerchantReference = new Map<string, ExcelRow[]>();
    for (const row of paidRows) {
      const merchantReference = normalize(
        getField(row, ["merchantReference", "Merchant Reference"])
      );
      if (!merchantReference) continue;

      const records = paidByMerchantReference.get(merchantReference) ?? [];
      records.push(row);
      paidByMerchantReference.set(merchantReference, records);
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
      const merchantReference = preserveText(
        getField(row, ["Merchant Reference"])
      );
      const rowOccurrences = occurrences.get(customerId) ?? 0;

      if (!merchantReference) {
        differences.push(
          paymentDifference(
            customerId,
            centralAmount ?? 0,
            merchantReference,
            rowOccurrences,
            "Merchant Reference is missing"
          )
        );
        continue;
      }

      const candidates =
        paidByMerchantReference.get(normalize(merchantReference)) ?? [];
      if (candidates.length === 0) {
        differences.push(
          paymentDifference(
            customerId,
            centralAmount ?? 0,
            merchantReference,
            rowOccurrences,
            "Merchant Reference not found in School Paid file"
          )
        );
        continue;
      }

      const matchedRow = candidates[0];
      matchedPayments.push({
        customerId,
        centralAmount: centralAmount ?? 0,
        centralMerchantReference: merchantReference,
        paidMerchantReference: preserveText(
          getField(matchedRow, ["merchantReference", "Merchant Reference"])
        ),
        paidBillId: normalize(
          getField(matchedRow, ["reference", "Bill ID", "Bills", "Bill"])
        ),
        paidAmount: parseAmount(
          getField(matchedRow, ["totalAmount", "Total Amount"])
        ),
        status: normalize(
          getField(matchedRow, ["paymentStatus", "Status"])
        ),
      });
    }

    return NextResponse.json({
      success: true,
      totalPayments: uniqueRows.length,
      duplicateCount: duplicateRows.length,
      duplicateRows,
      unduplicatedRows: uniqueRows,
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
