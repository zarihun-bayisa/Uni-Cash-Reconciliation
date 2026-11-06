"use client";

import {
  createContext,
  useContext,
  useState,
  type Dispatch,
  type ReactNode,
  type SetStateAction,
} from "react";

export interface PaymentDifference {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidBillId: string;
  paidAmount: number | null;
  status: string;
  occurrences: number;
  matchNote: string;
}

export interface MatchedPayment {
  customerId: string;
  centralAmount: number;
  centralMerchantReference: string;
  paidBillId: string;
  paidAmount: number;
  status: string;
}

export interface ReconciliationResult {
  success: boolean;
  totalPayments: number;
  duplicateCount: number;
  duplicateRows: Record<string, unknown>[];
  matchedCount: number;
  matchedPayments: MatchedPayment[];
  differenceCount: number;
  differences: PaymentDifference[];
}

interface ReconciliationContextValue {
  result: ReconciliationResult | null;
  setResult: Dispatch<SetStateAction<ReconciliationResult | null>>;
  schoolResult: ReconciliationResult | null;
  setSchoolResult: Dispatch<SetStateAction<ReconciliationResult | null>>;
}

const ReconciliationContext =
  createContext<ReconciliationContextValue | null>(null);

export function ReconciliationProvider({
  children,
}: {
  children: ReactNode;
}) {
  const [result, setResult] = useState<ReconciliationResult | null>(null);
  const [schoolResult, setSchoolResult] =
    useState<ReconciliationResult | null>(null);

  return (
    <ReconciliationContext.Provider
      value={{ result, setResult, schoolResult, setSchoolResult }}
    >
      {children}
    </ReconciliationContext.Provider>
  );
}

export function useReconciliation() {
  const context = useContext(ReconciliationContext);

  if (!context) {
    throw new Error(
      "useReconciliation must be used within ReconciliationProvider."
    );
  }

  return context;
}
