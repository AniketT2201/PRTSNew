import { useEffect, useState } from "react";

export const sanitize = (obj: any) =>
  Object.fromEntries(
    Object.entries(obj || {}).filter(
      ([_, v]) => v !== null && v !== undefined
    )
  );

import { parse, format, isValid } from "date-fns";

export const formatDate = (dateString: string) => {
    if (!dateString) return "";

    const currentYear = new Date().getFullYear();

    const formats = [
        "yyyy-MM-dd'T'HH:mm:ssX", // ISO
        "dd-MMM-yy",
        "M/d/yy",
        "M/d/yyyy",
        "d-MMM" // ⚠️ no year
    ];

    for (const f of formats) {
        let parsed = parse(dateString, f, new Date());

        // 🔥 Fix for "6-Oct" (no year)
        if (f === "d-MMM" && isValid(parsed)) {
            parsed.setFullYear(currentYear);
        }

        if (isValid(parsed)) {
            return format(parsed, "dd/MM/yyyy");
        }
    }

    return "-"; // fallback
};

// only for inputs without year, use baseDate to infer year
export const newFormatDate = (dateString: string, baseDate?: string) => {
    if (!dateString) return "";

    const baseYear = baseDate
        ? new Date(baseDate).getFullYear()
        : new Date().getFullYear();

    const formats = [
        "yyyy-MM-dd'T'HH:mm:ssX",
        "dd-MMM-yy",
        "M/d/yy",
        "d-MMM"
    ];

    for (const f of formats) {
        let parsed = parse(dateString, f, new Date());

        if (f === "d-MMM" && isValid(parsed)) {
            parsed.setFullYear(baseYear); // 🔥 correct logic
        }

        if (isValid(parsed)) {
            return format(parsed, "dd-MM-yyyy");
        }
    }

    return "-";
};

export const parseAmount = (value: any) => {
  if (value == null) return 0;

  return Number(String(value).replace(/,/g, ''));
};

export const formatAmount = (value: any): string => {
  if (value === null || value === undefined || value === "") {
    return "-";
  }

  const num = Number(value);

  if (isNaN(num)) {
    return "-";
  }

  return num.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
};

export const ConvertDatetoInputValue = (dateValue: any) => {
  if (!dateValue) return '';

  const date = new Date(dateValue);
  return date.toISOString().split('T')[0]; // YYYY-MM-DD
};

export const dateInputToISO = (dateValue: any) => {
  if (!dateValue) return null;

  const [y, m, d] = dateValue.split('-');
  return new Date(y, m - 1, d).toISOString();
};


export const useDebounce = <T>(value: T, delay = 500): T => {
  const [debouncedValue, setDebouncedValue] = useState<T>(value);

  useEffect(() => {
    const handler = setTimeout(() => {
      setDebouncedValue(value);
    }, delay);

    return () => {
      clearTimeout(handler);
    };
  }, [value, delay]);

  return debouncedValue;
}

export const getMonthFilters = () => {
    const currentDate = new Date();

    const getFormatted = (offset: number) => {
        const d = new Date(currentDate);
        d.setMonth(currentDate.getMonth() - offset);

        const month = d.toLocaleString('default', { month: 'short' });
        const year = d.getFullYear();

        return `${month}-${year}`;
    };

    return {
        Month1: getFormatted(0), // current month
        Month2: getFormatted(1), // previous month
        Month3: getFormatted(2),  // 2 months ago
        Month4: getFormatted(3)
    };
}

export const getFileType = (name: string) => {
  const ext = name?.split(".").pop()?.toLowerCase();

  if (["jpg", "jpeg", "png", "gif", "bmp", "webp"].includes(ext)) return "image";
  if (["pdf"].includes(ext)) return "pdf";
  if (["xls", "xlsx", "csv"].includes(ext)) return "excel";
  if (["doc", "docx"].includes(ext)) return "word";
  return "other";
};

export const formatFileSize = (bytes: number) => {
  if (!bytes) return "—";
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return (bytes / Math.pow(1024, i)).toFixed(1) + " " + sizes[i];
};
