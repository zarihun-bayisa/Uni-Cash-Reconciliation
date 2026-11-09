import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

type MatchedPayment = {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidMerchantReference?: string;
  paidBillId: string;
  paidAmount: number;
  status: string;
};

function isMatchedPayment(value: unknown): value is MatchedPayment {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }

  const payment = value as Record<string, unknown>;

  return (
    typeof payment.customerId === "string" &&
    typeof payment.centralAmount === "number" &&
    Number.isFinite(payment.centralAmount) &&
    typeof payment.centralMerchantReference === "string" &&
    (payment.paidMerchantReference === undefined ||
      typeof payment.paidMerchantReference === "string") &&
    typeof payment.paidBillId === "string" &&
    (payment.paidAmount === null ||
      (typeof payment.paidAmount === "number" &&
        Number.isFinite(payment.paidAmount))) &&
    typeof payment.status === "string"
  );
}

export async function POST(request: NextRequest) {
  try {
    const data: unknown = await request.json();
    const matchedPayments =
      typeof data === "object" &&
      data !== null &&
      "matchedPayments" in data
        ? data.matchedPayments
        : undefined;

    if (
      !Array.isArray(matchedPayments) ||
      matchedPayments.length === 0 ||
      !matchedPayments.every(isMatchedPayment)
    ) {
      return NextResponse.json(
        { error: "Invalid or empty matched payment data." },
        { status: 400 }
      );
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Central Payment Reconciliation";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Matched Payments");
    const hasPaidMerchantReference = matchedPayments.some(
      (payment) => typeof payment.paidMerchantReference === "string"
    );
    worksheet.columns = hasPaidMerchantReference
      ? [
          { header: "Customer ID", key: "customerId", width: 20 },
          { header: "Central Merchant Reference", key: "centralMerchantReference", width: 32 },
          { header: "School Paid Merchant Reference", key: "paidMerchantReference", width: 32 },
        ]
      : [
      { header: "Customer ID", key: "customerId", width: 20 },
      { header: "Central Amount", key: "centralAmount", width: 18 },
      {
        header: "Central Merchant Reference",
        key: "centralMerchantReference",
        width: 30,
      },
      { header: "Paid Bill ID", key: "paidBillId", width: 22 },
      { header: "Paid Amount", key: "paidAmount", width: 18 },
      { header: "Paid Status", key: "status", width: 18 },
    ];
    worksheet.addRows(matchedPayments);

    const header = worksheet.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF15803D" },
    };
    header.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    header.height = 32;

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.alignment = { vertical: "middle", wrapText: true };
        row.eachCell((cell) => {
          cell.border = {
            bottom: {
              style: "thin",
              color: { argb: "FFD1D5DB" },
            },
          };
        });
      }
    });

    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.autoFilter = {
      from: "A1",
      to: hasPaidMerchantReference ? "C1" : "F1",
    };

    const buffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          'attachment; filename="uni-cash-matched-payments.xlsx"',
      },
    });
  } catch (error) {
    console.error("Matched report generation error:", error);

    return NextResponse.json(
      { error: "Could not generate the matched payments report." },
      { status: 500 }
    );
  }
}
