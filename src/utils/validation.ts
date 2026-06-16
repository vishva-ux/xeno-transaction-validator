export interface ValidationRule {
  columnName: string;
  type: 'required' | 'email' | 'phone' | 'date' | 'numeric';
  required?: boolean;
  min?: number;
  max?: number;
  dateFormat?: string;
  phoneCountryColumn?: string; // Optional: column pointing to country name or code
  phoneDefaultCountry?: string; // Default country code
}

export interface CountryPhoneRule {
  code: string; // e.g. "IN" or "SG"
  name: string; // e.g. "India" or "Singapore"
  length: number; // expected digits length
  regex?: string; // optional custom regex
}

export interface ValidationConfig {
  phoneRules: CountryPhoneRule[];
  defaultPhoneCountry: string;
  dateFormat: string;
  requiredColumns: string[];
  numericColumns: string[];
}

export interface RowValidationError {
  column: string;
  message: string;
  severity: 'error' | 'warning';
}

export interface ProcessedRow {
  index: number;
  originalData: Record<string, any>;
  cleanedData: Record<string, any>;
  isValid: boolean;
  errors: RowValidationError[];
}

export interface ValidationSummary {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  errorRate: number;
  errorCountsByType: {
    missingValue: number;
    emailFormat: number;
    phoneFormat: number;
    dateFormat: number;
    numericFormat: number;
  };
  cityDistribution: Record<string, number>;
  countryDistribution: Record<string, number>;
}

// Default configuration rules
export const DEFAULT_COUNTRY_PHONE_RULES: CountryPhoneRule[] = [
  { code: 'IN', name: 'India', length: 10 },
  { code: 'SG', name: 'Singapore', length: 8 },
  { code: 'US', name: 'United States', length: 10 },
  { code: 'UK', name: 'United Kingdom', length: 10 },
  { code: 'AE', name: 'United Arab Emirates', length: 9 },
  { code: 'MY', name: 'Malaysia', length: 9 },
];

export const AVAILABLE_DATE_FORMATS = [
  { value: 'YYYY-MM-DD', label: 'YYYY-MM-DD (e.g. 2025-04-16)' },
  { value: 'DD-MM-YYYY', label: 'DD-MM-YYYY (e.g. 16-04-2025)' },
  { value: 'MM/DD/YYYY', label: 'MM/DD/YYYY (e.g. 04/16/2025)' },
  { value: 'DD/MM/YYYY', label: 'DD/MM/YYYY (e.g. 16/04/2025)' },
  { value: 'YYYY/MM/DD', label: 'YYYY/MM/DD (e.g. 2025/04/16)' },
];

/**
 * Validates an email address format.
 */
export function isValidEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates date based on format.
 */
export function isValidDate(dateStr: string, format: string): boolean {
  if (!dateStr) return false;
  
  // Clean up any extra whitespace
  dateStr = dateStr.trim();
  
  let regex: RegExp;
  let yearIdx: number, monthIdx: number, dayIdx: number;

  switch (format) {
    case 'YYYY-MM-DD':
      regex = /^(\d{4})-(\d{2})-(\d{2})$/;
      yearIdx = 1; monthIdx = 2; dayIdx = 3;
      break;
    case 'DD-MM-YYYY':
      regex = /^(\d{2})-(\d{2})-(\d{4})$/;
      yearIdx = 3; monthIdx = 2; dayIdx = 1;
      break;
    case 'MM/DD/YYYY':
      regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      yearIdx = 3; monthIdx = 1; dayIdx = 2;
      break;
    case 'DD/MM/YYYY':
      regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      yearIdx = 3; monthIdx = 2; dayIdx = 1;
      break;
    case 'YYYY/MM/DD':
      regex = /^(\d{4})\/(\d{2})\/(\d{2})$/;
      yearIdx = 1; monthIdx = 2; dayIdx = 3;
      break;
    default:
      // Fallback: check if standard Date constructor works
      const testDate = new Date(dateStr);
      return !isNaN(testDate.getTime());
  }

  const match = dateStr.match(regex);
  if (!match) return false;

  const y = parseInt(match[yearIdx], 10);
  const m = parseInt(match[monthIdx], 10) - 1; // JS Date months are 0-indexed
  const d = parseInt(match[dayIdx], 10);

  const parsedDate = new Date(y, m, d);
  
  // Verify date details to catch Feb 30th etc.
  return (
    parsedDate.getFullYear() === y &&
    parsedDate.getMonth() === m &&
    parsedDate.getDate() === d
  );
}

/**
 * Validates a phone number based on country rules.
 */
export function validatePhoneNumber(
  phoneVal: any,
  countryVal: string | undefined,
  rules: CountryPhoneRule[],
  defaultCountryCode: string
): { isValid: boolean; errorMsg?: string } {
  if (phoneVal === undefined || phoneVal === null || phoneVal === '') {
    return { isValid: false, errorMsg: 'Phone number is missing.' };
  }

  // Convert phone to clean numeric string (remove +,-, spaces, brackets)
  const phoneStr = String(phoneVal).replace(/[^0-9]/g, '');
  
  if (!phoneStr) {
    return { isValid: false, errorMsg: 'Phone number must contain numbers.' };
  }

  // Try to find the rule based on country name or code
  let cleanCountry = (countryVal || '').trim().toUpperCase();
  let rule = rules.find(r => 
    r.code.toUpperCase() === cleanCountry || 
    r.name.toUpperCase() === cleanCountry
  );

  // If no country rule matched, fallback to defaultCountryCode
  if (!rule) {
    rule = rules.find(r => r.code === defaultCountryCode) || rules[0];
  }

  // Validate digit length
  if (phoneStr.length !== rule.length) {
    return { 
      isValid: false, 
      errorMsg: `Phone must be exactly ${rule.length} digits for ${rule.name} (Found ${phoneStr.length} digits).` 
    };
  }

  return { isValid: true };
}

/**
 * Primary processing function.
 */
export function validateDataset(
  data: Record<string, any>[],
  config: ValidationConfig,
  columnMapping: {
    email?: string;
    phone?: string;
    country?: string;
    date?: string;
    city?: string;
  }
): { rows: ProcessedRow[]; summary: ValidationSummary } {
  const processedRows: ProcessedRow[] = [];
  
  const summary: ValidationSummary = {
    totalRows: data.length,
    validRows: 0,
    invalidRows: 0,
    errorRate: 0,
    errorCountsByType: {
      missingValue: 0,
      emailFormat: 0,
      phoneFormat: 0,
      dateFormat: 0,
      numericFormat: 0,
    },
    cityDistribution: {},
    countryDistribution: {},
  };

  data.forEach((row, idx) => {
    const errors: RowValidationError[] = [];
    const cleanedData: Record<string, any> = { ...row };
    
    // 1. Missing Value / Required Field Checks
    config.requiredColumns.forEach(col => {
      const val = row[col];
      if (val === undefined || val === null || String(val).trim() === '') {
        errors.push({
          column: col,
          message: `Field '${col}' is required and cannot be empty.`,
          severity: 'error'
        });
        summary.errorCountsByType.missingValue++;
      }
    });

    // 2. Email Validation
    const emailCol = columnMapping.email || 'email';
    if (row[emailCol] !== undefined && row[emailCol] !== null && String(row[emailCol]).trim() !== '') {
      const emailStr = String(row[emailCol]).trim();
      if (!isValidEmail(emailStr)) {
        errors.push({
          column: emailCol,
          message: `Invalid email address format.`,
          severity: 'error'
        });
        summary.errorCountsByType.emailFormat++;
      }
    }

    // 3. Phone Number Validation
    const phoneCol = columnMapping.phone || 'phone_number';
    const countryCol = columnMapping.country || 'country';
    const hasPhoneVal = row[phoneCol] !== undefined && row[phoneCol] !== null && String(row[phoneCol]).trim() !== '';
    
    if (hasPhoneVal) {
      const countryVal = row[countryCol];
      const phoneValidation = validatePhoneNumber(
        row[phoneCol],
        countryVal,
        config.phoneRules,
        config.defaultPhoneCountry
      );
      
      if (!phoneValidation.isValid) {
        errors.push({
          column: phoneCol,
          message: phoneValidation.errorMsg || 'Invalid phone format.',
          severity: 'error'
        });
        summary.errorCountsByType.phoneFormat++;
      }
    }

    // 4. Date Validation
    const dateCol = columnMapping.date || 'signup_date';
    const hasDateVal = row[dateCol] !== undefined && row[dateCol] !== null && String(row[dateCol]).trim() !== '';
    
    if (hasDateVal) {
      const rawDate = String(row[dateCol]).trim();
      // If datetime format has time segment but we only parse date (e.g. YYYY-MM-DD 00:00:00)
      const datePart = rawDate.split(' ')[0]; // Split space to drop time
      
      if (!isValidDate(datePart, config.dateFormat)) {
        errors.push({
          column: dateCol,
          message: `Date '${rawDate}' does not match format ${config.dateFormat} or is invalid.`,
          severity: 'error'
        });
        summary.errorCountsByType.dateFormat++;
      }
    }

    // 5. Numeric Column Checks (Prices, Amounts, Quantities must be positive numbers)
    config.numericColumns.forEach(col => {
      const val = row[col];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        const num = Number(val);
        if (isNaN(num)) {
          errors.push({
            column: col,
            message: `Field '${col}' must be a valid number.`,
            severity: 'error'
          });
          summary.errorCountsByType.numericFormat++;
        } else if (num < 0) {
          errors.push({
            column: col,
            message: `Field '${col}' cannot be negative (Found: ${num}).`,
            severity: 'error'
          });
          summary.errorCountsByType.numericFormat++;
        }
      }
    });

    const isValid = errors.length === 0;
    if (isValid) {
      summary.validRows++;
    } else {
      summary.invalidRows++;
    }

    // Capture distribution statistics for dashboard
    const cityCol = columnMapping.city || 'city';
    if (row[cityCol]) {
      const city = String(row[cityCol]).trim();
      summary.cityDistribution[city] = (summary.cityDistribution[city] || 0) + 1;
    }
    
    if (row[countryCol]) {
      const country = String(row[countryCol]).trim();
      summary.countryDistribution[country] = (summary.countryDistribution[country] || 0) + 1;
    }

    processedRows.push({
      index: idx + 1,
      originalData: row,
      cleanedData,
      isValid,
      errors
    });
  });

  summary.errorRate = summary.totalRows > 0 ? (summary.invalidRows / summary.totalRows) * 100 : 0;
  
  return {
    rows: processedRows,
    summary
  };
}

/**
 * Splits rows into smaller arrays of specified chunk size.
 */
export function chunkData<T>(data: T[], chunkSize: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < data.length; i += chunkSize) {
    chunks.push(data.slice(i, i + chunkSize));
  }
  return chunks;
}
