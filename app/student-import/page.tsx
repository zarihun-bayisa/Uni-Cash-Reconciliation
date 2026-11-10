"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import AppHeader from "../app-header";

const OUTPUT_HEADERS = [
  "ID",
  "FERST NAME",
  "SECOND NAME",
  "LAST NAME",
  "GRADE",
  "SECTION",
  "PHONE",
] as const;

const EARLY_GRADE_OPTIONS = ["Nursery", "UKG", "KG-1", "KG-2", "KG-3"];
const GRADE_OPTIONS = [
  ...EARLY_GRADE_OPTIONS,
  ...Array.from({ length: 12 }, (_, index) => String(index + 1)),
];

type GradeFilter =
  | "all"
  | "nursery-kg3"
  | "grades-1-4"
  | "grades-5-8"
  | "grades-9-12"
  | "specific";

type Field =
  | "id"
  | "firstName"
  | "secondName"
  | "lastName"
  | "grade"
  | "section"
  | "phone";

type ColumnMap = Record<Field, number | null> & { fullName: number | null };

type StudentRecord = {
  id: string;
  sourceSheet: string;
  amount: string;
  firstName: string;
  secondName: string;
  lastName: string;
  grade: string;
  section: string;
  phone: string;
  generatedId: boolean;
  issues: string[];
};

type ImportedSheet = {
  name: string;
  recordCount: number;
  columns: ColumnMap;
  amountColumn: number | null;
};

type DetectedSheet = {
  name: string;
  sheetGrade: string;
  rows: unknown[][];
  header: { index: number; columns: ColumnMap };
  recordCount: number;
  amountColumn: number | null;
};

type IdState = {
  usedIds: Set<string>;
  fallbackGenerator: IdGenerator;
};

type IdGenerator = {
  prefix: string;
  suffix: string;
  width: number;
  nextNumber: number;
};

const HEADER_ALIASES: Record<Field, string[]> = {
  id: [
    "id",
    "student id",
    "student number",
    "student no",
    "student code",
    "code",
    "admission number",
    "admission no",
    "admission code",
    "registration number",
    "registration no",
    "registration code",
  ],
  firstName: ["first name", "ferst name", "given name", "forename", "name 1", "first"],
  secondName: [
    "second name",
    "middle name",
    "father name",
    "name 2",
    "father's name",
    "second",
  ],
  lastName: [
    "last name",
    "third name",
    "surname",
    "family name",
    "grandfather name",
    "name 3",
    "last",
  ],
  grade: [
    "grade",
    "grade level",
    "grade section",
    "grade and section",
    "class",
    "class level",
    "level",
  ],
  section: ["section", "stream", "division", "class section"],
  phone: [
    "phone",
    "phone number",
    "mobile",
    "mobile number",
    "telephone",
    "contact",
    "contact number",
    "cellphone",
  ],
};

const FULL_NAME_ALIASES = ["name", "full name", "student name", "student full name"];
const PAYMENT_HEADERS = [
  "amount",
  "paid amount",
  "payment amount",
  "fee",
  "fees",
  "payment",
];
const FIELDS = Object.keys(HEADER_ALIASES) as Field[];
const FIELD_LABELS: Record<Field, string> = {
  id: "ID",
  firstName: "First name",
  secondName: "Second name",
  lastName: "Last name",
  grade: "Grade",
  section: "Section",
  phone: "Phone",
};

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function detectColumns(row: unknown[]): ColumnMap {
  const columns: ColumnMap = {
    id: null,
    firstName: null,
    secondName: null,
    lastName: null,
    grade: null,
    section: null,
    phone: null,
    fullName: null,
  };

  row.forEach((value, index) => {
    const header = normalizeHeader(value);
    if (!header) return;

    if (
      columns.fullName === null &&
      FULL_NAME_ALIASES.some((alias) => normalizeHeader(alias) === header)
    ) {
      columns.fullName = index;
    }

    for (const field of FIELDS) {
      if (
        columns[field] === null &&
        (HEADER_ALIASES[field].some((alias) => normalizeHeader(alias) === header) ||
          (field === "phone" &&
            /(phone|phon|mobile|telephone|contact|cell)/.test(header)))
      ) {
        columns[field] = index;
      }
    }
  });

  return columns;
}

function findHeaderRow(rows: unknown[][]): { index: number; columns: ColumnMap } | null {
  const scanLimit = Math.min(rows.length, 20);

  for (let index = 0; index < scanLimit; index += 1) {
    const columns = detectColumns(rows[index]);
    const recognized =
      FIELDS.filter((field) => columns[field] !== null).length +
      (columns.fullName !== null ? 1 : 0);
    const hasNameColumn =
      columns.fullName !== null ||
      columns.firstName !== null ||
      columns.secondName !== null ||
      columns.lastName !== null;

    if (recognized >= 2 && hasNameColumn) {
      return { index, columns };
    }
  }

  return null;
}

function findAmountColumn(header: unknown[]): number | null {
  const index = header.findIndex((value) => {
    const name = normalizeHeader(value);
    return (
      !/(penalt|penalit|fine)/.test(name) &&
      (PAYMENT_HEADERS.some((alias) => normalizeHeader(alias) === name) ||
        /(amount|fee)/.test(name))
    );
  });
  return index < 0 ? null : index;
}

function cellText(value: unknown): string {
  return String(value ?? "").trim();
}

function normalizeGrade(value: string): string {
  const compact = value.toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (
    compact === "NUR" ||
    compact === "NURSERY" ||
    compact === "NURSERY1"
  ) {
    return "Nursery";
  }
  if (compact === "UKG") return "UKG";

  const kgMatch = compact.match(/^KG0?([1-3])$/);
  if (kgMatch) return `KG-${kgMatch[1]}`;

  const gradeMatch = compact.match(/^(?:GRADE|CLASS)?0?([1-9]|1[0-2])(?:ST|ND|RD|TH)?$/);
  if (gradeMatch) return gradeMatch[1];

  return "";
}

function parseGradeAndSection(value: string): {
  grade: string;
  section: string;
} {
  const compact = value.toUpperCase().replace(/[^A-Z0-9]/g, "");
  const ukgMatch = compact.match(/^UKG([A-Z])?$/);
  if (ukgMatch) return { grade: "UKG", section: ukgMatch[1] ?? "" };

  const kgMatch = compact.match(/^KG0?([1-3])([A-Z])?$/);
  if (kgMatch) return { grade: `KG-${kgMatch[1]}`, section: kgMatch[2] ?? "" };

  const nurseryMatch = compact.match(/^(?:NURSERY|NUR)([A-Z])?$/);
  if (nurseryMatch) {
    return { grade: "Nursery", section: nurseryMatch[1] ?? "" };
  }

  if (compact === "NURSERY1") return { grade: "Nursery", section: "" };

  const gradeMatch = compact.match(
    /^(?:GRADE|CLASS)?0?([1-9]|1[0-2])(?:ST|ND|RD|TH)?([A-Z])?$/
  );
  if (gradeMatch) {
    return { grade: gradeMatch[1], section: gradeMatch[2] ?? "" };
  }

  return { grade: normalizeGrade(value), section: "" };
}

function gradeFromSheetName(sheetName: string): string {
  const earlyGradeMatch = sheetName.match(
    /(?:^|[^A-Z0-9])(NURSERY1|NURSERY|NUR|UKG|KG\s*[-_]?\s*[1-3])(?=$|[^A-Z0-9])/i
  );
  if (earlyGradeMatch) return parseGradeAndSection(earlyGradeMatch[1]).grade;

  const match = sheetName.match(
    /(?:^|[^A-Z0-9])(?:GRADE|CLASS|G)\s*[-_ ]?\s*(0?[1-9]|1[0-2])(?:ST|ND|RD|TH)?(?=$|[^A-Z0-9])/i
  );
  return match ? normalizeGrade(match[1]) : "";
}

function sortStudentsByGrade(records: StudentRecord[]): StudentRecord[] {
  return records
    .map((record, index) => ({ record, index }))
    .sort((left, right) => {
      const leftRank = GRADE_OPTIONS.indexOf(left.record.grade);
      const rightRank = GRADE_OPTIONS.indexOf(right.record.grade);
      const normalizedLeftRank =
        leftRank < 0 ? GRADE_OPTIONS.length : leftRank;
      const normalizedRightRank =
        rightRank < 0 ? GRADE_OPTIONS.length : rightRank;
      return normalizedLeftRank - normalizedRightRank || left.index - right.index;
    })
    .map(({ record }) => record);
}

function gradeMatchesFilter(
  grade: string,
  filter: GradeFilter,
  specificGrades: string[]
): boolean {
  if (filter === "all") return true;
  if (filter === "specific") return specificGrades.includes(grade);
  if (filter === "nursery-kg3") return EARLY_GRADE_OPTIONS.includes(grade);

  const numericGrade = Number(grade);
  if (!Number.isInteger(numericGrade)) return false;
  if (filter === "grades-1-4") return numericGrade >= 1 && numericGrade <= 4;
  if (filter === "grades-5-8") return numericGrade >= 5 && numericGrade <= 8;
  return numericGrade >= 9 && numericGrade <= 12;
}

function normalizePhone(value: string): string {
  const digits = value.replace(/\D/g, "");
  if (/^\d{9}$/.test(digits)) return `0${digits}`;
  return value;
}

function createIdGenerator(startValue: string): IdGenerator {
  const trimmedValue = startValue.trim();
  const numericSuffix = trimmedValue.match(/^(.*?)(\d+)$/);
  if (!numericSuffix) {
    return {
      prefix: trimmedValue,
      suffix: "",
      width: 3,
      nextNumber: 1,
    };
  }

  const nextNumber = Number(numericSuffix[2]);
  if (!Number.isSafeInteger(nextNumber)) {
    throw new Error("Enter a starting ID with a valid numeric suffix.");
  }

  return {
    prefix: numericSuffix[1],
    suffix: "",
    width: numericSuffix[2].length,
    nextNumber,
  };
}

function inferIdGenerator(ids: string[]): IdGenerator | null {
  if (ids.length === 0) return null;

  const parsedIds = ids.map((id) => {
    const numbers = [...id.matchAll(/\d+/g)];
    return numbers.map((match) => ({
      value: Number(match[0]),
      width: match[0].length,
      index: match.index ?? 0,
    }));
  });
  const maxNumberCount = Math.max(...parsedIds.map((numbers) => numbers.length));

  if (maxNumberCount === 0) {
    const baseId = ids[0];
    return {
      prefix: `${baseId}/`,
      suffix: "",
      width: 3,
      nextNumber: 1,
    };
  }

  let selectedPosition = 0;
  let selectedUniqueCount = -1;
  let selectedCoverage = -1;

  for (let position = 0; position < maxNumberCount; position += 1) {
    const values = parsedIds.flatMap((numbers) =>
      numbers[position] ? [numbers[position].value] : []
    );
    const uniqueCount = new Set(values).size;
    if (
      uniqueCount > selectedUniqueCount ||
      (uniqueCount === selectedUniqueCount && values.length > selectedCoverage)
    ) {
      selectedPosition = position;
      selectedUniqueCount = uniqueCount;
      selectedCoverage = values.length;
    }
  }

  const representativeIndex = parsedIds.findIndex(
    (numbers) => numbers[selectedPosition] !== undefined
  );
  const representative = ids[representativeIndex];
  const selectedNumber = parsedIds[representativeIndex][selectedPosition];
  const prefix = representative.slice(0, selectedNumber.index);
  const suffix = representative.slice(
    selectedNumber.index + selectedNumber.width
  );
  const matchingNumbers = parsedIds.flatMap((numbers) =>
    numbers[selectedPosition] ? [numbers[selectedPosition]] : []
  );

  return {
    prefix,
    suffix,
    width: Math.max(...matchingNumbers.map((number) => number.width)),
    nextNumber: Math.max(...matchingNumbers.map((number) => number.value)) + 1,
  };
}

function createRecords(
  rows: unknown[][],
  headerIndex: number,
  columns: ColumnMap,
  amountColumn: number | null,
  sourceSheet: string,
  sheetGrade: string,
  idState: IdState,
  idGenerator: IdGenerator | null,
  defaultAmount: string
): StudentRecord[] {
  const sourceRows = rows.slice(headerIndex + 1).filter((row) =>
    [...FIELDS, "fullName" as const].some((field) => {
      const column = columns[field];
      return column !== null && cellText(row[column]) !== "";
    })
  );
  let previousGrade = "";

  return sourceRows.map((row) => {
    const read = (field: Field) => {
      const column = columns[field];
      return column === null ? "" : cellText(row[column]);
    };
    const combinedNameColumn = columns.fullName;
    const nameParts =
      combinedNameColumn === null
        ? []
        : cellText(row[combinedNameColumn]).split(/\s+/).filter(Boolean);
    const id = read("id").replace(/\s+/g, "");
    const firstName = read("firstName") || nameParts[0] || "";
    const secondName = read("secondName") || nameParts[1] || "";
    const lastName =
      read("lastName") || (nameParts.length > 2 ? nameParts.slice(2).join(" ") : "");
    const sourceGrade = read("grade");
    const parsedGrade = parseGradeAndSection(sourceGrade);
    const detectedGrade = parsedGrade.grade;
    const grade = detectedGrade || (!sourceGrade ? sheetGrade || previousGrade : "");
    const sourceSection = read("section") || parsedGrade.section;
    const sourcePhone = read("phone");
    const sourceAmount =
      amountColumn === null ? "" : cellText(row[amountColumn]);
    const issues: string[] = [];

    if (!firstName) issues.push("First name is missing");
    if (!secondName) issues.push("Second name is missing");
    if (!lastName) issues.push("Third name defaulted to AA");
    if (!grade) {
      issues.push(
        sourceGrade
          ? `Grade "${sourceGrade}" is not recognized`
          : "Grade is missing"
      );
    } else if (!detectedGrade && !sheetGrade) {
      issues.push(`Grade carried forward as ${grade}`);
    } else if (!detectedGrade) {
      issues.push(`Grade taken from sheet name: ${grade}`);
    }
    if (!sourceSection) issues.push("Section defaulted to A");
    if (!sourcePhone) issues.push("Phone defaulted to 09");

    const finalLastName = lastName || "AA";
    if (detectedGrade) previousGrade = detectedGrade;

    let generatedId = false;
    let studentId = id;
    if (!studentId) {
      do {
        if (idGenerator) {
          studentId = `${idGenerator.prefix}${String(
            idGenerator.nextNumber
          ).padStart(idGenerator.width, "0")}${idGenerator.suffix}`;
          idGenerator.nextNumber += 1;
        } else {
          studentId = `${idState.fallbackGenerator.prefix}${String(
            idState.fallbackGenerator.nextNumber
          ).padStart(idState.fallbackGenerator.width, "0")}${idState.fallbackGenerator.suffix}`;
          idState.fallbackGenerator.nextNumber += 1;
        }
      } while (idState.usedIds.has(studentId.toLowerCase()));
      idState.usedIds.add(studentId.toLowerCase());
      generatedId = true;
    }

    return {
      id: studentId,
      sourceSheet,
      amount: sourceAmount || defaultAmount,
      firstName,
      secondName,
      lastName: finalLastName,
      grade,
      section: sourceSection || "A",
      phone: sourcePhone ? normalizePhone(sourcePhone) : "09",
      generatedId,
      issues,
    };
  });
}

export default function StudentImportPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [idPrefix, setIdPrefix] = useState("MS/");
  const [defaultAmount, setDefaultAmount] = useState("1700");
  const [students, setStudents] = useState<StudentRecord[]>([]);
  const [textStudents, setTextStudents] = useState<StudentRecord[]>([]);
  const [sources, setSources] = useState<ImportedSheet[]>([]);
  const [showTextConverter, setShowTextConverter] = useState(false);
  const [smsText, setSmsText] = useState("");
  const [gradeFilter, setGradeFilter] = useState<GradeFilter>("all");
  const [specificGrades, setSpecificGrades] = useState<string[]>([]);
  const [error, setError] = useState("");
  const [textError, setTextError] = useState("");
  const [loading, setLoading] = useState(false);

  function getFilteredStudents(records: StudentRecord[]): StudentRecord[] {
    return sortStudentsByGrade(
      records.filter((student) =>
        gradeMatchesFilter(student.grade, gradeFilter, specificGrades)
      )
    );
  }

  async function handleConvert() {
    if (files.length === 0) {
      setError("Choose at least one Excel file first.");
      return;
    }

    if (files.length > 20) {
      setError("Select no more than 20 Excel files at a time.");
      return;
    }

    if (!defaultAmount.trim() || !Number.isFinite(Number(defaultAmount)) || Number(defaultAmount) < 0) {
      setError("Enter a valid default payment amount.");
      return;
    }

    setLoading(true);
    setError("");
    setStudents([]);
    setSources([]);

    try {
      const detectedSheets: DetectedSheet[] = [];
      for (const [fileIndex, file] of files.entries()) {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        if (workbook.SheetNames.length === 0) {
          throw new Error(`${file.name} does not contain any worksheets.`);
        }

        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) continue;

          const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
            header: 1,
            defval: "",
            raw: false,
          });
          const header = findHeaderRow(rows);
          if (!header) continue;
          const amountColumn = findAmountColumn(rows[header.index] ?? []);

          const sourceRows = rows.slice(header.index + 1).filter((row) =>
            [...FIELDS, "fullName" as const].some((field) => {
              const column = header.columns[field];
              return column !== null && cellText(row[column]) !== "";
            })
          );
          if (sourceRows.length === 0) continue;

          detectedSheets.push({
            name: `${file.name} (${fileIndex + 1}) / ${sheetName}`,
            sheetGrade: gradeFromSheetName(sheetName),
            rows,
            header,
            recordCount: sourceRows.length,
            amountColumn,
          });
        }
      }

      if (detectedSheets.length === 0) {
        throw new Error(
          "Could not detect student records in any worksheet. Include headings for at least one student name and another field."
        );
      }

      const idState: IdState = {
        usedIds: new Set<string>(),
        fallbackGenerator: createIdGenerator(idPrefix),
      };
      const sheetIdGenerators = new Map<string, IdGenerator | null>();
      for (const detectedSheet of detectedSheets) {
        const idColumn = detectedSheet.header.columns.id;
        const ids =
          idColumn === null
            ? []
            : detectedSheet.rows
                .slice(detectedSheet.header.index + 1)
                .map((row) => cellText(row[idColumn]).replace(/\s+/g, ""))
                .filter(Boolean);
        for (const id of ids) {
          idState.usedIds.add(id.toLowerCase());
        }
        sheetIdGenerators.set(
          detectedSheet.name,
          inferIdGenerator(ids)
        );
      }

      const importedSheets: ImportedSheet[] = [];
      const records = detectedSheets.flatMap(
        ({ name, sheetGrade, rows, header, amountColumn }) => {
        const sheetRecords = createRecords(
          rows,
          header.index,
          header.columns,
          amountColumn,
          name,
          sheetGrade,
          idState,
          sheetIdGenerators.get(name) ?? null,
          defaultAmount
        );
        importedSheets.push({
          name,
          recordCount: sheetRecords.length,
          columns: header.columns,
          amountColumn,
        });
        return sheetRecords;
        }
      );

      setStudents(sortStudentsByGrade(records));
      setSources(importedSheets);
    } catch (convertError) {
      setError(
        convertError instanceof Error
          ? convertError.message
          : "Could not read this Excel workbook."
      );
    } finally {
      setLoading(false);
    }
  }

  function parseSmsText(): StudentRecord[] {
    const records: StudentRecord[] = [];
    let activeGrade = "";
    let activeSection = "";
    let pendingPhoneIndex: number | null = null;
    const usedIds = new Set<string>();
    const idGenerator = createIdGenerator(idPrefix);
    const lines = smsText.split(/\r?\n/);

    const isGradeLine = (line: string) => {
      const match = line.match(/^G(?:RADE)?\s*(?:[:=]\s*)?(.+)$/i);
      if (!match) return null;

      const grade = parseGradeAndSection(match[1]);
      return grade.grade ? grade : null;
    };

    const isPhoneLine = (line: string) => {
      const digits = line.replace(/\D/g, "");
      return digits.length >= 9 && digits.length <= 13;
    };

    for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
      const rawLine = lines[lineIndex];
      const line = rawLine.trim();
      if (!line) continue;

      const parsedGrade = isGradeLine(line);
      if (parsedGrade) {
        const { grade, section } = parsedGrade;
        activeGrade = grade;
        activeSection = section || "A";
        if (pendingPhoneIndex !== null) {
          records[pendingPhoneIndex].grade = activeGrade;
          records[pendingPhoneIndex].section = activeSection;
          records[pendingPhoneIndex].issues = records[
            pendingPhoneIndex
          ].issues.filter((issue) => issue !== "Grade is missing or not recognized");
        }
        continue;
      }

      if (isPhoneLine(line)) {
        if (pendingPhoneIndex !== null) {
          records[pendingPhoneIndex].phone = normalizePhone(
            line.replace(/\D/g, "")
          );
          pendingPhoneIndex = null;
        }
        continue;
      }

      const numberedName = line.match(/^\s*\d+\s*[.)-]\s*(.+?)\s*$/);
      const name = (numberedName?.[1] ?? line).trim();
      const nameParts = name.split(/\s+/).filter(Boolean);
      if (nameParts.length === 0) continue;

      let nextRecordFollows = false;
      for (let nextIndex = lineIndex + 1; nextIndex < lines.length; nextIndex += 1) {
        const nextLine = lines[nextIndex].trim();
        if (!nextLine) continue;
        if (isGradeLine(nextLine)) continue;
        if (isPhoneLine(nextLine)) nextRecordFollows = true;
        break;
      }

      if (!numberedName && (!activeGrade || !nextRecordFollows)) {
        continue;
      }

      const issues: string[] = [];
      if (!nameParts[1]) issues.push("Second name is missing");
      if (nameParts.length < 3) issues.push("Third name defaulted to AA");
      if (!activeGrade) issues.push("Grade is missing or not recognized");

      let id = "";
      do {
        id = `${idGenerator.prefix}${String(idGenerator.nextNumber).padStart(
          idGenerator.width,
          "0"
        )}${idGenerator.suffix}`;
        idGenerator.nextNumber += 1;
      } while (usedIds.has(id.toLowerCase()));
      usedIds.add(id.toLowerCase());

      records.push({
        id,
        sourceSheet: "SMS text",
        amount: defaultAmount,
        firstName: nameParts[0],
        secondName: nameParts[1] ?? "",
        lastName: nameParts.length > 2 ? nameParts.slice(2).join(" ") : "AA",
        grade: activeGrade,
        section: activeSection || "A",
        phone: "09",
        generatedId: true,
        issues,
      });
      pendingPhoneIndex = records.length - 1;
    }

    for (const record of records) {
      if (record.phone === "09") record.issues.push("Phone defaulted to 09");
    }

    return records;
  }

  function handleConvertSms() {
    setTextError("");
    setTextStudents([]);

    if (!smsText.trim()) {
      setTextError("Enter the student text before converting.");
      return;
    }

    if (!idPrefix.trim()) {
      setTextError("Enter an ID prefix before converting the text.");
      return;
    }

    if (!defaultAmount.trim() || !Number.isFinite(Number(defaultAmount)) || Number(defaultAmount) < 0) {
      setTextError("Enter a valid default payment amount.");
      return;
    }

    let records: StudentRecord[];
    try {
      records = parseSmsText();
    } catch (parseError) {
      setTextError(
        parseError instanceof Error
          ? parseError.message
          : "Could not generate student IDs from this starting ID."
      );
      return;
    }

    if (records.length === 0) {
      setTextError(
        "No student names were detected. Enter each student name on its own line, with an optional grade heading and phone number."
      );
      return;
    }

    setTextStudents(sortStudentsByGrade(records));
  }

  function downloadStudents(records: StudentRecord[], filename: string) {
    if (records.length === 0) {
      throw new Error("Convert student records before downloading.");
    }

    const data = records.map((student) => [
      student.id,
      student.firstName,
      student.secondName,
      student.lastName,
      student.grade,
      student.section,
      student.phone,
    ]);
    const worksheet = XLSX.utils.aoa_to_sheet([[...OUTPUT_HEADERS], ...data]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Students");
    XLSX.writeFile(workbook, filename);
  }

  function downloadPaymentFile(records: StudentRecord[], filename: string) {
    if (records.length === 0) {
      throw new Error("Convert student records before downloading.");
    }

    const data = records.map((student) => {
      const numericAmount = Number(student.amount.replace(/,/g, ""));
      return [
        student.id,
        Number.isFinite(numericAmount) ? numericAmount : student.amount,
      ];
    });
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Student Id", "Amount"],
      ...data,
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Other Payment");
    XLSX.writeFile(workbook, filename);
  }

  function handleDownload() {
    try {
      downloadStudents(getFilteredStudents(students), "converted-students.xlsx");
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the converted workbook."
      );
    }
  }

  function handleDownloadSms() {
    try {
      downloadStudents(getFilteredStudents(textStudents), "sms-students.xlsx");
    } catch (downloadError) {
      setTextError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the SMS student workbook."
      );
    }
  }

  function handleDownloadPayment() {
    try {
      downloadPaymentFile(
        getFilteredStudents(students),
        "other-payment-file.xlsx"
      );
    } catch (downloadError) {
      setError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the other payment file."
      );
    }
  }

  function handleDownloadSmsPayment() {
    try {
      downloadPaymentFile(
        getFilteredStudents(textStudents),
        "other-payment-file.xlsx"
      );
    } catch (downloadError) {
      setTextError(
        downloadError instanceof Error
          ? downloadError.message
          : "Could not download the other payment file."
      );
    }
  }

  const issueCount = students.filter((student) => student.issues.length > 0).length;
  const generatedIdCount = students.filter((student) => student.generatedId).length;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50/70 text-slate-900 md:flex">
      <AppHeader />
      <main className="min-w-0 flex-1">
        <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8 sm:py-10">
          <div className="relative mb-8 overflow-hidden rounded-[2rem] bg-gradient-to-br from-slate-950 via-blue-950 to-indigo-900 px-6 py-8 text-white shadow-2xl shadow-blue-950/15 sm:px-9 sm:py-10">
            <div aria-hidden="true" className="absolute -right-10 -top-28 h-80 w-80 rounded-full border-[40px] border-white/[0.04]" />
            <div aria-hidden="true" className="absolute -bottom-36 right-40 h-64 w-64 rounded-full bg-cyan-400/20 blur-3xl" />
            <div className="relative">
              <p className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.18em] text-blue-100">
                <span className="h-1.5 w-1.5 rounded-full bg-cyan-300" />
                Student records
              </p>
              <h2 className="mt-4 text-3xl font-bold tracking-tight text-white sm:text-4xl">
                Student Import
              </h2>
              <p className="mt-3 max-w-3xl text-sm leading-6 text-blue-100 sm:text-base">
                Convert student spreadsheets into the Uni-Cash template.
              </p>
            </div>
          </div>

          <div className="rounded-3xl border border-slate-200/80 bg-white/95 p-5 shadow-xl shadow-slate-200/50 sm:p-7">
            <div className="mb-6 flex items-start gap-4">
              <div aria-hidden="true" className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-blue-50 text-blue-700 ring-1 ring-blue-100">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" className="h-5 w-5">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M7 3.75h7l4 4v12.5H7a2 2 0 0 1-2-2v-12.5a2 2 0 0 1 2-2Z" />
                  <path strokeLinecap="round" strokeLinejoin="round" d="M14 3.75v4h4M8.5 12h7M8.5 15.5h7" />
                </svg>
              </div>
              <div>
                <h3 className="text-lg font-bold tracking-tight text-slate-900">Upload student workbooks</h3>
                <p className="mt-1 text-sm leading-6 text-slate-500">
                  Select student Excel files to prepare an import-ready workbook.
                </p>
              </div>
            </div>

            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-[minmax(0,1fr)_minmax(14rem,0.6fr)_minmax(10rem,0.45fr)_auto] lg:items-end">
              <label className="text-sm font-medium text-slate-700">
                Excel files (.xlsx or .xls, up to 20)
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 p-3 text-sm shadow-inner shadow-slate-100 file:mr-4 file:cursor-pointer file:rounded-lg file:border-0 file:bg-blue-700 file:px-3 file:py-2 file:font-semibold file:text-white hover:border-blue-300"
                  onChange={(event) => {
                    const selectedFiles = Array.from(event.target.files ?? []);
                    setStudents([]);
                    setSources([]);
                    setTextStudents([]);
                    if (selectedFiles.length > 20) {
                      setFiles([]);
                      setError("Select no more than 20 Excel files at a time.");
                    } else {
                      setFiles(selectedFiles);
                      setError("");
                    }
                  }}
                />
              </label>
              <label className="text-sm font-medium text-slate-700">
                Generated ID prefix or starting ID
                <input
                  type="text"
                  value={idPrefix}
                  onChange={(event) => {
                    setIdPrefix(event.target.value);
                    setStudents([]);
                    setTextStudents([]);
                    setSources([]);
                    setError("");
                    setTextError("");
                  }}
                  placeholder="MS/ or SA/620"
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-blue-400 focus:bg-white focus:outline-none"
                  aria-describedby="id-prefix-help"
                />
                <span id="id-prefix-help" className="mt-1 block text-xs font-normal text-slate-500">
                  Enter a prefix (MS/) to generate MS/001, MS/002, and so on,
                  or a starting ID (SA/620) to generate SA/620, SA/621, and so on.
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
                    setStudents([]);
                    setTextStudents([]);
                    setSources([]);
                    setError("");
                    setTextError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50/80 px-3 py-3 text-sm shadow-inner shadow-slate-100 focus:border-blue-400 focus:bg-white focus:outline-none"
                  aria-describedby="default-amount-help"
                />
                <span id="default-amount-help" className="mt-1 block text-xs font-normal text-slate-500">
                  Used only where a workbook has no amount. Default: 1700.
                </span>
              </label>
              <button
                type="button"
                onClick={handleConvert}
                disabled={
                  files.length === 0 ||
                  loading ||
                  !idPrefix.trim() ||
                  !defaultAmount.trim()
                }
                className="rounded-xl bg-gradient-to-r from-blue-700 to-indigo-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-blue-700/20 transition hover:-translate-y-0.5 hover:shadow-xl hover:shadow-blue-700/25 disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:translate-y-0"
              >
                {loading ? "Converting..." : "Convert and preview"}
              </button>
              <p className="text-sm text-slate-500 sm:col-span-2 lg:col-span-4">
                {files.length === 0
                  ? "No files selected."
                  : `${files.length} file${files.length === 1 ? "" : "s"} selected: ${files.map((selectedFile) => selectedFile.name).join(", ")}`}
              </p>
            </div>

            {error && (
              <p role="alert" className="mt-4 rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700">
                {error}
              </p>
            )}
          </div>

          <div className="mt-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="grid gap-4 sm:grid-cols-[minmax(14rem,0.6fr)_minmax(0,1fr)] sm:items-start">
              <label className="text-sm font-semibold text-slate-800">
                Grade filter for downloads
                <select
                  value={gradeFilter}
                  onChange={(event) => {
                    setGradeFilter(event.target.value as GradeFilter)
                    setError("");
                    setTextError("");
                  }}
                  className="mt-2 block w-full rounded-xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm font-medium focus:border-blue-400 focus:bg-white focus:outline-none"
                >
                  <option value="all">All grades</option>
                  <option value="nursery-kg3">Nursery – KG-3</option>
                  <option value="grades-1-4">Grades 1–4</option>
                  <option value="grades-5-8">Grades 5–8</option>
                  <option value="grades-9-12">Grades 9–12</option>
                  <option value="specific">Choose specific grades</option>
                </select>
              </label>
              {gradeFilter === "specific" && (
                <fieldset>
                  <legend className="text-sm font-semibold text-slate-800">
                    Select one or more grades
                  </legend>
                  <div className="mt-2 flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() => setSpecificGrades(GRADE_OPTIONS)}
                      className="rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
                    >
                      Select all
                    </button>
                    <button
                      type="button"
                      onClick={() => setSpecificGrades([])}
                      className="rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
                    >
                      Clear
                    </button>
                  </div>
                  <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-5">
                    {GRADE_OPTIONS.map((grade) => (
                      <label
                        key={grade}
                        className="flex items-center gap-2 rounded-lg border border-slate-100 px-3 py-2 text-sm text-slate-700"
                      >
                        <input
                          type="checkbox"
                          checked={specificGrades.includes(grade)}
                          onChange={(event) =>
                            setSpecificGrades((current) =>
                              event.target.checked
                                ? [...current, grade]
                                : current.filter((selected) => selected !== grade)
                            )
                          }
                          className="h-4 w-4 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                        />
                        {grade}
                      </label>
                    ))}
                  </div>
                </fieldset>
              )}
            </div>
            <p className="mt-3 text-xs leading-5 text-slate-500">
              Applies to both student and payment downloads, including SMS/text
              records. Previews remain complete. Downloads are ordered Nursery,
              UKG, KG-1–KG-3, then grades 1–12.
            </p>
            {(students.length > 0 || textStudents.length > 0) && (
              <p className="mt-1 text-xs font-medium text-slate-600">
                Matching records: {getFilteredStudents(students).length} from
                workbooks and {getFilteredStudents(textStudents).length} from
                SMS/text.
              </p>
            )}
          </div>

          <div className="mt-7">
            <button
              type="button"
              onClick={() => {
                setShowTextConverter((visible) => !visible);
                setTextError("");
              }}
              aria-expanded={showTextConverter}
              className="inline-flex items-center gap-3 rounded-2xl border border-blue-200/80 bg-white px-5 py-3.5 text-sm font-semibold text-blue-900 shadow-lg shadow-slate-200/50 transition hover:-translate-y-0.5 hover:border-blue-300 hover:shadow-xl"
            >
              <span className="flex h-8 w-8 items-center justify-center rounded-xl bg-blue-700 text-sm text-white">✦</span>
              {showTextConverter
                ? "Close SMS/Text converter"
                : "Convert SMS/Text to Excel"}
            </button>
          </div>

          {showTextConverter && (
            <div className="mt-4 rounded-3xl border border-slate-200/80 bg-white/95 p-5 shadow-xl shadow-slate-200/50 sm:p-7">
              <h3 className="text-lg font-bold tracking-tight text-slate-900">Paste student text</h3>
              <p className="mb-4 mt-1 text-sm text-slate-500">
                Add a grade heading such as G=8 or Grade=6, followed by each
                student&apos;s name and phone number on separate lines. The
                generated IDs use the prefix or starting ID above.
              </p>
              <textarea
                value={smsText}
                onChange={(event) => {
                  setSmsText(event.target.value);
                  setTextStudents([]);
                  setTextError("");
                }}
                rows={10}
                placeholder={`G=8
Yusra Suleyman
0913035983

G=2
monenus belay
0917690223

Grade=6
Selam G/egziabher
0933111586`}
                className="w-full rounded-2xl border border-slate-200 bg-slate-50/80 p-4 font-mono text-sm leading-6 text-slate-800 shadow-inner shadow-slate-100 placeholder:text-slate-400 focus:border-blue-400 focus:bg-white focus:outline-none"
              />
              <div className="mt-4 flex flex-wrap gap-3">
                <button
                  type="button"
                  onClick={handleConvertSms}
                  disabled={!smsText.trim() || !idPrefix.trim()}
                  className="rounded-xl bg-gradient-to-r from-blue-700 to-indigo-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-blue-700/20 transition hover:-translate-y-0.5 hover:shadow-xl disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:translate-y-0"
                >
                  Generate student preview
                </button>
                {textStudents.length > 0 && (
                  <>
                    <button
                      type="button"
                      onClick={handleDownloadSms}
                      disabled={getFilteredStudents(textStudents).length === 0}
                      className="rounded-xl bg-gradient-to-r from-emerald-700 to-teal-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-emerald-700/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      Download Student Excel (
                      {getFilteredStudents(textStudents).length})
                    </button>
                    <button
                      type="button"
                      onClick={handleDownloadSmsPayment}
                      disabled={getFilteredStudents(textStudents).length === 0}
                      className="rounded-xl bg-gradient-to-r from-teal-700 to-cyan-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-teal-700/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      Download Other Payment File (
                      {getFilteredStudents(textStudents).length})
                    </button>
                  </>
                )}
              </div>
              {textError && (
                <p role="alert" className="mt-4 rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700">
                  {textError}
                </p>
              )}
              {textStudents.length > 0 && (
                <div className="mt-6">
                  <p className="mb-3 text-sm font-medium text-slate-700">
                    Preview: {textStudents.length} students ·{" "}
                    {textStudents.filter((student) => student.issues.length > 0).length}{" "}
                    records need attention
                  </p>
                  <div className="max-h-[28rem] overflow-auto rounded-2xl border border-slate-200 shadow-md shadow-slate-200/40">
                    <table className="min-w-full divide-y divide-slate-200 text-left text-sm">
                      <thead className="sticky top-0 bg-slate-100 text-xs uppercase tracking-wide text-slate-600">
                        <tr>
                          {OUTPUT_HEADERS.map((header) => (
                            <th key={header} className="whitespace-nowrap px-4 py-3 font-semibold">
                              {header}
                            </th>
                          ))}
                          <th className="whitespace-nowrap px-4 py-3 font-semibold">Amount</th>
                          <th className="min-w-56 px-4 py-3 font-semibold">Review notes</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {textStudents.map((student, index) => (
                          <tr
                            key={`${student.id}-${index}`}
                            className={student.issues.length ? "bg-amber-50/60" : "bg-white"}
                          >
                            <td className="whitespace-nowrap px-4 py-3">{student.id}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.firstName}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.secondName}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.lastName}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.grade || "—"}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.section}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.phone}</td>
                            <td className="whitespace-nowrap px-4 py-3">{student.amount}</td>
                            <td className="px-4 py-3 text-amber-800">
                              {student.issues.join(" · ") || "—"}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {students.length > 0 && sources.length > 0 && (
            <div className="mt-8 rounded-3xl border border-slate-200/80 bg-white/95 p-5 shadow-xl shadow-slate-200/50 sm:p-7">
              <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
                <div>
                  <h3 className="text-lg font-semibold">Review converted students</h3>
                  <p className="mt-1 text-sm text-slate-500">
                    {students.length} records · {generatedIdCount} generated IDs ·{" "}
                    {issueCount} records need attention
                  </p>
                </div>
                <div className="flex flex-wrap gap-3">
                  <button
                    type="button"
                    onClick={handleDownload}
                    disabled={getFilteredStudents(students).length === 0}
                    className="shrink-0 rounded-xl bg-gradient-to-r from-emerald-700 to-teal-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-emerald-700/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    Download Student Excel ({getFilteredStudents(students).length})
                  </button>
                  <button
                    type="button"
                    onClick={handleDownloadPayment}
                    disabled={getFilteredStudents(students).length === 0}
                    className="shrink-0 rounded-xl bg-gradient-to-r from-teal-700 to-cyan-700 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-teal-700/15 transition hover:-translate-y-0.5 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    Download Other Payment File ({getFilteredStudents(students).length})
                  </button>
                </div>
              </div>

              <div className="mt-5 space-y-3 rounded-2xl border border-slate-100 bg-slate-50/80 p-4 text-sm">
                <p className="font-semibold text-slate-800">
                  Worksheets included ({sources.length})
                </p>
                {sources.map((source) => (
                  <div key={source.name} className="border-t border-slate-200 pt-3">
                    <p className="font-medium text-slate-700">
                      {source.name} · {source.recordCount} records · first matching
                      phone column used
                    </p>
                    <p className="mt-1 text-slate-500">
                      {FIELDS.map((field) => {
                        const column = source.columns[field];
                        const description =
                          column !== null
                            ? `${FIELD_LABELS[field]}: column ${column + 1}`
                            : field === "id"
                              ? "ID: generated"
                              : field === "firstName" ||
                                  field === "secondName" ||
                                  field === "lastName"
                                ? source.columns.fullName !== null
                                  ? `${FIELD_LABELS[field]}: split from column ${source.columns.fullName + 1}`
                                  : `${FIELD_LABELS[field]}: not found`
                                : `${FIELD_LABELS[field]}: defaulted or flagged`;
                        return description;
                      }).join(" · ")}
                    </p>
                  </div>
                ))}
              </div>

              <div className="mt-5 max-h-[36rem] overflow-auto rounded-2xl border border-slate-200 shadow-md shadow-slate-200/40">
                <table className="min-w-full divide-y divide-slate-200 text-left text-sm">
                  <thead className="sticky top-0 bg-slate-100 text-xs uppercase tracking-wide text-slate-600">
                    <tr>
                      {OUTPUT_HEADERS.map((header) => (
                        <th key={header} className="whitespace-nowrap px-4 py-3 font-semibold">
                          {header}
                        </th>
                      ))}
                      <th className="whitespace-nowrap px-4 py-3 font-semibold">Amount</th>
                      <th className="whitespace-nowrap px-4 py-3 font-semibold">
                        Source sheet
                      </th>
                      <th className="min-w-64 px-4 py-3 font-semibold">Review notes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {students.map((student, index) => (
                      <tr
                        key={`${student.id}-${index}`}
                        className={student.issues.length ? "bg-amber-50/60" : "bg-white"}
                      >
                        <td className="whitespace-nowrap px-4 py-3">{student.id}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.firstName}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.secondName}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.lastName}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.grade || "—"}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.section}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.phone}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.amount}</td>
                        <td className="whitespace-nowrap px-4 py-3">{student.sourceSheet}</td>
                        <td className="px-4 py-3 text-amber-800">
                          {[
                            ...(student.generatedId ? ["ID generated"] : []),
                            ...student.issues,
                          ].join(" · ") || "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="mt-3 text-xs text-slate-500">
                The downloaded workbook contains only the required template
                columns: ID, FERST NAME, SECOND NAME, LAST NAME, GRADE, SECTION,
                and PHONE.
              </p>
            </div>
          )}
        </section>
      </main>
    </div>
  );
}
