
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export async function POST(request: NextRequest) {
  try {
    const data = await request.json();
    const differences = data.differences;
    const isUniCashFormat = data.format === "unicash";

    if (!Array.isArray(differences)) {
      return NextResponse.json(
        { error: "Invalid difference data." },
        { status: 400 }
      );
    }

    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Central Payment Reconciliation";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Differences");

    if (isUniCashFormat) {
      worksheet.columns = [
        { header: "ContractNumber", key: "customerId", width: 20 },
        { header: "Central Bill Id", key: "centralMerchantReference", width: 22 },
        { header: "central Amount", key: "centralAmount", width: 18 },
        { header: "Bill Id - From bill", key: "paidBillId", width: 24 },
        { header: "Bill Amount", key: "paidAmount", width: 18 },
        { header: "Status", key: "status", width: 18 },
      ];
      worksheet.addRows(
        differences.map((difference) => ({
          customerId: difference.customerId,
          centralMerchantReference: difference.centralMerchantReference,
          centralAmount: difference.centralAmount,
          paidBillId: difference.paidBillId,
          paidAmount: difference.paidAmount,
          status: difference.status,
        }))
      );
    } else {
      const hasPaidMerchantReference = differences.some(
        (difference) => typeof difference.paidMerchantReference === "string"
      );
      if (hasPaidMerchantReference) {
        worksheet.columns = [
          { header: "Customer ID", key: "customerId", width: 20 },
          { header: "Central Merchant Reference", key: "centralMerchantReference", width: 32 },
          { header: "School Paid Merchant Reference", key: "paidMerchantReference", width: 32 },
          { header: "Match Note", key: "matchNote", width: 42 },
        ];
        worksheet.addRows(differences);
      } else {
        worksheet.columns = [
        { header: "Customer ID", key: "customerId", width: 20 },
        { header: "004496 Amount", key: "centralAmount", width: 18 },
        {
          header: "004496 Merchant Reference",
          key: "centralMerchantReference",
          width: 30,
        },
        { header: "Bill Id", key: "paidBillId", width: 22 },
        { header: "Paid Amount", key: "paidAmount", width: 18 },
        { header: "Status", key: "status", width: 16 },
        {
          header: "Occurrences in 004496",
          key: "occurrences",
          width: 22,
        },
        { header: "Match Note", key: "matchNote", width: 42 },
      ];
      worksheet.addRows(differences);
      }
    }

    // Format the heading row.
    const header = worksheet.getRow(1);
    header.font = {
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
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

    // Format the data rows.
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.alignment = {
          vertical: "middle",
          wrapText: true,
        };

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
      to: isUniCashFormat
        ? "F1"
        : differences.some(
              (difference) => typeof difference.paidMerchantReference === "string"
            )
          ? "D1"
          : "H1",
    };

    const buffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          `attachment; filename="${isUniCashFormat ? "unicash bill difference.xlsx" : "Feres_Meda_Difference.xlsx"}"`,
      },
    });
  } catch (error) {
    console.error("Report generation error:", error);

    return NextResponse.json(
      { error: "Could not generate the Excel report." },
      { status: 500 }
    );
  }
}
