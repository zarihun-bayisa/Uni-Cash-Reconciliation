import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

type WaterDifference = {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidBillId: string;
  paidAmount: number | null;
  status: string;
};

function isWaterDifference(value: unknown): value is WaterDifference {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }

  const difference = value as Record<string, unknown>;

  return (
    typeof difference.customerId === "string" &&
    typeof difference.centralAmount === "number" &&
    Number.isFinite(difference.centralAmount) &&
    typeof difference.centralMerchantReference === "string" &&
    typeof difference.paidBillId === "string" &&
    (difference.paidAmount === null ||
      (typeof difference.paidAmount === "number" &&
        Number.isFinite(difference.paidAmount))) &&
    typeof difference.status === "string"
  );
}

export async function POST(request: NextRequest) {
  try {
    const data: unknown = await request.json();
    const differences =
      typeof data === "object" && data !== null && "differences" in data
        ? data.differences
        : undefined;

    if (
      !Array.isArray(differences) ||
      differences.length === 0 ||
      !differences.every(isWaterDifference)
    ) {
      return NextResponse.json(
        { error: "Invalid or empty water difference data." },
        { status: 400 }
      );
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Uni-Cash Payment Reconciliation";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Water Differences");
    worksheet.columns = [
      { header: "ContractNumber", key: "customerId", width: 20 },
      { header: "Central Bill Id", key: "centralMerchantReference", width: 22 },
      { header: "central Amount", key: "centralAmount", width: 18 },
      { header: "Bill Id - From bill", key: "paidBillId", width: 24 },
      { header: "Bill Amount", key: "paidAmount", width: 18 },
      { header: "Status", key: "status", width: 18 },
    ];

    worksheet.addRows(differences);

    const header = worksheet.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E3A8A" },
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
            bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
          };
        });
      }
    });

    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.autoFilter = { from: "A1", to: "F1" };

    const buffer = await workbook.xlsx.writeBuffer();
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          'attachment; filename="unicash-bill-difference.xlsx"',
      },
    });
  } catch (error) {
    console.error("Water difference report generation error:", error);
    return NextResponse.json(
      { error: "Could not generate the Uni-Cash water difference report." },
      { status: 500 }
    );
  }
}
