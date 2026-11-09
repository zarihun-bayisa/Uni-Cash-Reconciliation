
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

function customerIdText(value: unknown): string {
  return String(value ?? "");
}

function normalizeHeader(value: unknown): string {
  return normalize(value).replace(/[^A-Z0-9]/g, "");
}

function getField(
  row: ExcelRow,
  possibleNames: string[]
): unknown {
  const keys = Object.keys(row);

  for (const name of possibleNames) {
    const foundKey = keys.find(
      (key) => normalizeHeader(key) === normalizeHeader(name)
    );

    if (foundKey !== undefined) {
      return row[foundKey];
    }
  }

  return undefined;
}

function parseAmount(value: unknown): number | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const cleaned = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "");
  const amount = Number(cleaned);

  return Number.isFinite(amount) ? amount : null;
}

function readExcel(buffer: Buffer): ExcelRow[] {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

  if (!firstSheet) {
    throw new Error("The Excel file does not contain a worksheet.");
  }

  return XLSX.utils.sheet_to_json<ExcelRow>(firstSheet, {
    defval: "",
    raw: false,
  });
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
        { error: "Please upload both Central and Paid Excel files." },
        { status: 400 }
      );
    }

    const allowedExtensions = [".xlsx", ".xls"];

    for (const file of [...centralFiles, ...paidFiles]) {
      const validExtension = allowedExtensions.some((extension) =>
        file.name.toLowerCase().endsWith(extension)
      );

      if (!validExtension) {
        return NextResponse.json(
          { error: `Unsupported file format: ${file.name}` },
          { status: 400 }
        );
      }
    }

    const centralRows = (
      await Promise.all(
        centralFiles.map(async (file) =>
          readExcel(Buffer.from(await file.arrayBuffer()))
        )
      )
    ).flat();
    const paidRows = (
      await Promise.all(
        paidFiles.map(async (file) =>
          readExcel(Buffer.from(await file.arrayBuffer()))
        )
      )
    ).flat();

    const uniqueCentralRows: ExcelRow[] = [];
    const duplicateRows: ExcelRow[] = [];
    const seenCentralPayments = new Set<string>();
    let duplicateCount = 0;

    for (const row of centralRows) {
      const customerId = customerIdText(
        getField(row, ["Customer Id", "Customer ID"])
      );
      const amount = parseAmount(getField(row, ["Amount"]));
      const merchantReference = normalize(
        getField(row, ["Merchant Reference"])
      );

      if (customerId.trim() && amount !== null) {
        const paymentKey = JSON.stringify([
          customerId,
          amount,
          merchantReference,
        ]);

        if (seenCentralPayments.has(paymentKey)) {
          duplicateCount++;
          duplicateRows.push(row);
          continue;
        }

        seenCentralPayments.add(paymentKey);
      }

      uniqueCentralRows.push(row);
    }

    // Index Paid records using Contract Number.
    const paidByCustomer = new Map<string, ExcelRow[]>();

    for (const row of paidRows) {
      const customerId = customerIdText(
        getField(row, ["Contract Number"])
      );

      if (!customerId.trim()) continue;

      const existing = paidByCustomer.get(customerId) ?? [];
      existing.push(row);
      paidByCustomer.set(customerId, existing);
    }

    // Count occurrences of each Customer ID in the Central file.
    const occurrences = new Map<string, number>();

    for (const row of uniqueCentralRows) {
      const customerId = customerIdText(
        getField(row, ["Customer Id", "Customer ID"])
      );

      if (customerId.trim()) {
        occurrences.set(
          customerId,
          (occurrences.get(customerId) ?? 0) + 1
        );
      }
    }

    let matchedCount = 0;
    const matchedPayments: MatchedPayment[] = [];
    const differences: PaymentDifference[] = [];

    for (const centralRow of uniqueCentralRows) {
      const customerId = customerIdText(
        getField(centralRow, ["Customer Id", "Customer ID"])
      );

      const centralAmount = parseAmount(
        getField(centralRow, ["Amount"])
      );

      const merchantReference = normalize(
        getField(centralRow, ["Merchant Reference"])
      );

      if (!customerId.trim() || centralAmount === null || !merchantReference) {
        differences.push({
          customerId,
          centralAmount: centralAmount ?? 0,
          centralMerchantReference: merchantReference,
          paidBillId: "",
          paidAmount: null,
          status: "",
          occurrences: occurrences.get(customerId) ?? 0,
          matchNote: "Incomplete Central payment information",
        });
        continue;
      }

      const candidates = paidByCustomer.get(customerId) ?? [];

      if (candidates.length === 0) {
        differences.push({
          customerId,
          centralAmount,
          centralMerchantReference: merchantReference,
          paidBillId: "",
          paidAmount: null,
          status: "",
          occurrences: occurrences.get(customerId) ?? 0,
          matchNote: "Customer ID not found in Paid file",
        });
        continue;
      }

      // All three fields must match.
      const exactMatch = candidates.find((paidRow) => {
        const paidBillId = normalize(
          getField(paidRow, ["Bill Id", "Bill ID"])
        );

        const paidAmount = parseAmount(
          getField(paidRow, ["Total Amount"])
        );

        return (
          paidBillId === merchantReference &&
          paidAmount !== null &&
          paidAmount === centralAmount
        );
      });

      if (exactMatch) {
        const paidAmount = parseAmount(
          getField(exactMatch, ["Total Amount"])
        );

        if (paidAmount !== null) {
          matchedCount++;
          matchedPayments.push({
            customerId,
            centralAmount,
            centralMerchantReference: merchantReference,
            paidBillId: normalize(
              getField(exactMatch, ["Bill Id", "Bill ID"])
            ),
            paidAmount,
            status: normalize(getField(exactMatch, ["Status"])),
          });
          continue;
        }
      }

      // Do not arbitrarily choose a record if several candidates exist.
      const candidateBillIds = [
        ...new Set(
          candidates.map((row) =>
            normalize(getField(row, ["Bill Id", "Bill ID"]))
          ).filter(Boolean)
        ),
      ];

      const candidateAmounts = [
        ...new Set(
          candidates.map((row) =>
            parseAmount(getField(row, ["Total Amount"]))
          ).filter((amount): amount is number => amount !== null)
        ),
      ];

      const candidateStatuses = [
        ...new Set(
          candidates.map((row) =>
            normalize(getField(row, ["Status"]))
          ).filter(Boolean)
        ),
      ];

      let matchNote = "Payment details do not match";

      if (candidates.length > 1) {
        matchNote = "Multiple Paid records found; review required";
      } else {
        const paidRow = candidates[0];

        const paidBillId = normalize(
          getField(paidRow, ["Bill Id", "Bill ID"])
        );

        const paidAmount = parseAmount(
          getField(paidRow, ["Total Amount"])
        );

        if (paidBillId !== merchantReference && paidAmount !== centralAmount) {
          matchNote = "Merchant Reference and Amount do not match";
        } else if (paidBillId !== merchantReference) {
          matchNote = "Merchant Reference and Bill ID do not match";
        } else {
          matchNote = "Central Amount and Paid Amount do not match";
        }
      }

      differences.push({
        customerId,
        centralAmount,
        centralMerchantReference: merchantReference,
        paidBillId:
          candidateBillIds.length === 1
            ? candidateBillIds[0]
            : candidateBillIds.join(", "),
        paidAmount:
          candidateAmounts.length === 1
            ? candidateAmounts[0]
            : null,
        status: candidateStatuses.join(", "),
        occurrences: occurrences.get(customerId) ?? 0,
        matchNote,
      });
    }

    return NextResponse.json({
      success: true,
      totalPayments: uniqueCentralRows.length,
      duplicateCount,
      duplicateRows,
      matchedCount,
      matchedPayments,
      differenceCount: differences.length,
      differences,
    });
  } catch (error) {
    console.error("Reconciliation error:", error);

    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "An unexpected error occurred.",
      },
      { status: 500 }
    );
  }
}
