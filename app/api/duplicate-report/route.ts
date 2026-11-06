import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

type DuplicateRow = Record<string, string | number | boolean | null>;

export async function POST(request: NextRequest) {
  try {
    const data: unknown = await request.json();
    const duplicateRows =
      typeof data === "object" &&
      data !== null &&
      "duplicateRows" in data
        ? data.duplicateRows
        : undefined;

    if (
      !Array.isArray(duplicateRows) ||
      duplicateRows.length === 0 ||
      !duplicateRows.every(
        (row): row is DuplicateRow =>
          typeof row === "object" &&
          row !== null &&
          !Array.isArray(row) &&
          Object.values(row).every(
            (value) =>
              value === null ||
              typeof value === "string" ||
              typeof value === "number" ||
              typeof value === "boolean"
          )
      )
    ) {
      return NextResponse.json(
        { error: "Invalid or empty duplicate data." },
        { status: 400 }
      );
    }

    const headers = [
      ...new Set(duplicateRows.flatMap((row) => Object.keys(row))),
    ];

    if (headers.length === 0) {
      return NextResponse.json(
        { error: "The duplicate rows do not contain any columns." },
        { status: 400 }
      );
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Central Payment Reconciliation";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Duplicate Rows");
    worksheet.columns = headers.map((header) => ({
      header,
      key: header,
      width: Math.min(Math.max(header.length + 4, 16), 40),
    }));
    worksheet.addRows(duplicateRows);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.autoFilter = {
      from: "A1",
      to: `${worksheet.getColumn(headers.length).letter}1`,
    };

    const buffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          'attachment; filename="uni-cash-duplicate-rows.xlsx"',
      },
    });
  } catch (error) {
    console.error("Duplicate report generation error:", error);

    return NextResponse.json(
      { error: "Could not generate the duplicate Excel report." },
      { status: 500 }
    );
  }
}
