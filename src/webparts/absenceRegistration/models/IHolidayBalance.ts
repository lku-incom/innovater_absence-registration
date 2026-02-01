/**
 * Data models for Holiday Balance tracking
 * Based on Danish Holiday Law (Ferieloven) - "Samtidighedsferie" from September 2020
 *
 * Holiday year runs from September 1st to August 31st
 * Monthly accrual rate: 2.08 days (25 days / 12 months)
 */

export interface IHolidayBalance {
  Id?: number;
  DataverseId?: string; // GUID from Dataverse (cr_holidaybalanceid)
  Name?: string; // Auto-generated: "EmployeeName - HolidayYear"
  EmployeeEmail: string;
  EmployeeName: string;
  HolidayYear: string; // e.g., "2024-2025" (Sept 2024 - Aug 2025)

  // Statutory vacation (Ferie) - 25 days/year under Danish law
  AccruedDays: number; // Total accrued this holiday year (2.08 per month)
  UsedDays: number; // Approved absences of type "Ferie"
  PendingDays: number; // Pending approval absences
  AvailableDays: number; // Calculated: accrued + carriedOver + transferredIn - used - pending
  CarriedOverDays: number; // From previous year (legacy field, use TransferredInDays for new logic)

  // Transfer tracking (Ferieloven §19 - max 5 days beyond 20 can be transferred)
  TransferredInDays: number; // Days transferred IN from previous holiday year (requires agreement)
  TransferredOutDays: number; // Days agreed to transfer OUT to next holiday year
  HasTransferAgreement: boolean; // Whether written agreement exists for transfer
  TransferAgreementDate?: Date; // Date agreement was made (must be by Dec 31)

  // Feriefridage (contract-based extra days, typically 5/year - NOT covered by Ferieloven)
  FeriefridageAccrued: number;
  FeriefridageUsed: number;
  FeriefridageAvailable: number;
  FeriefridageTransferredIn: number; // Company policy dependent
  FeriefridageTransferredOut: number; // Company policy dependent

  // Tracking
  LastAccrualDate?: Date;
  EmploymentStartDate?: Date;
  IsActive: boolean; // True for current holiday year

  // Timestamps
  Created?: Date;
  Modified?: Date;
}

export interface IAccrualHistory {
  Id?: number;
  DataverseId?: string; // GUID from Dataverse (cr_accrualhistoryid)
  Name?: string; // Auto-generated: "EmployeeName - Month Year"
  EmployeeEmail: string;
  EmployeeName: string;
  HolidayYear: string;

  // Accrual details
  AccrualDate: Date;
  AccrualMonth: number; // 1-12
  AccrualYear: number;
  DaysAccrued: number; // Typically 2.08
  FeriefridageAccrued?: number;
  BalanceAfterAccrual?: number;
  AccrualType: AccrualType;
  Notes?: string;

  // Timestamps
  Created?: Date;
}

export type AccrualType =
  | 'Månedlig optjening' // Monthly Accrual (100000000)
  | 'Årsstart overførsel' // Year Start Carryover (100000001)
  | 'Manuel justering' // Manual Adjustment (100000002)
  | 'Startsaldo'; // Initial Balance (100000003)

/**
 * Calculated holiday balance - computed from accrual history and absence registrations
 * This is not stored in Dataverse, but calculated on-the-fly
 */
export interface ICalculatedHolidayBalance {
  EmployeeEmail: string;
  EmployeeName: string;
  HolidayYear: string;

  // Feriedage (statutory vacation)
  TotalAccruedDays: number; // Sum of all accrual history DaysAccrued
  UsedDays: number; // Sum of approved Ferie absences
  PendingDays: number; // Sum of pending Ferie absences
  AvailableDays: number; // TotalAccruedDays - UsedDays - PendingDays

  // Feriefridage (contract-based extra days)
  TotalAccruedFeriefridage: number; // Sum of all accrual history FeriefridageAccrued
  UsedFeriefridage: number; // Sum of approved Feriefridage absences
  PendingFeriefridage: number; // Sum of pending Feriefridage absences
  AvailableFeriefridage: number; // TotalAccruedFeriefridage - UsedFeriefridage - PendingFeriefridage

  // Source data counts (for debugging/display)
  AccrualHistoryCount: number;
  AbsenceRegistrationCount: number;
}

/**
 * Constants for Danish Holiday Law calculations
 *
 * Key rules from Ferieloven (Danish Holiday Act):
 * - Holiday year (ferieår): September 1 - August 31
 * - Holiday-taking period (ferieafholdelsesperiode): September 1 - December 31 (16 months)
 * - First 4 weeks (20 days) MUST be taken - cannot be transferred
 * - 5th week (5 days) can be transferred with written agreement by December 31
 * - No legal limit on accumulation over years (with continuous agreement)
 * - Feriefridage are NOT covered by Ferieloven - rules depend on contract/overenskomst
 */
export const DANISH_HOLIDAY_CONSTANTS = {
  // Monthly accrual rate: 25 days / 12 months = 2.08333...
  MONTHLY_ACCRUAL_RATE: 2.08,

  // Total statutory vacation days per year
  ANNUAL_VACATION_DAYS: 25,

  // Mandatory vacation days that MUST be taken (first 4 weeks)
  MANDATORY_VACATION_DAYS: 20,

  // Maximum days that can be transferred per year (5th week only)
  MAX_TRANSFER_DAYS_PER_YEAR: 5,

  // Holiday year starts on September 1st
  HOLIDAY_YEAR_START_MONTH: 9, // September (1-indexed)

  // Holiday year ends on August 31st
  HOLIDAY_YEAR_END_MONTH: 8, // August (1-indexed)

  // Holiday-taking period ends December 31 (of the following year)
  HOLIDAY_TAKING_PERIOD_END_MONTH: 12, // December (1-indexed)
  HOLIDAY_TAKING_PERIOD_END_DAY: 31,

  // Transfer agreement deadline (December 31)
  TRANSFER_DEADLINE_MONTH: 12,
  TRANSFER_DEADLINE_DAY: 31,

  // Typical feriefridage per year (contract-based, not statutory)
  TYPICAL_FERIEFRIDAGE_PER_YEAR: 5,

  // Monthly feriefridage accrual (if accrued monthly)
  MONTHLY_FERIEFRIDAGE_RATE: 0.42, // 5 / 12 = 0.4166...
} as const;

/**
 * Dataverse field mapping for Holiday Balance
 */
export const HOLIDAY_BALANCE_FIELD_MAP = {
  id: 'cr_holidaybalanceid',
  name: 'cr_name',
  employeeEmail: 'cr_employeeemail',
  employeeName: 'cr_employeename',
  holidayYear: 'cr_holidayyear',
  accruedDays: 'cr_accrueddays',
  usedDays: 'cr_useddays',
  pendingDays: 'cr_pendingdays',
  availableDays: 'cr_availabledays',
  carriedOverDays: 'cr_carriedoverdays',
  // Transfer tracking fields
  transferredInDays: 'cr_transferredindays',
  transferredOutDays: 'cr_transferredoutdays',
  hasTransferAgreement: 'cr_hastransferagreement',
  transferAgreementDate: 'cr_transferagreementdate',
  // Feriefridage fields
  feriefridageAccrued: 'cr_feriefridageaccrued',
  feriefridageUsed: 'cr_feriefridageused',
  feriefridageAvailable: 'cr_feriefridageavailable',
  feriefridageTransferredIn: 'cr_feriefridagetransferredin',
  feriefridageTransferredOut: 'cr_feriefridagetransferredout',
  // Tracking fields
  lastAccrualDate: 'cr_lastaccrualdate',
  employmentStartDate: 'cr_employmentstartdate',
  isActive: 'cr_isactive',
} as const;

/**
 * Dataverse field mapping for Accrual History
 */
export const ACCRUAL_HISTORY_FIELD_MAP = {
  id: 'cr_accrualhistoryid',
  name: 'cr_name',
  employeeEmail: 'cr_employeeemail',
  employeeName: 'cr_employeename',
  holidayYear: 'cr_holidayyear',
  accrualDate: 'cr_accrualdate',
  accrualMonth: 'cr_accrualmonth',
  accrualYear: 'cr_accrualyear',
  daysAccrued: 'cr_daysaccrued',
  feriefridageAccrued: 'cr_feriefridageaccrued',
  balanceAfterAccrual: 'cr_balanceafteraccrual',
  accrualType: 'cr_accrualtype',
  notes: 'cr_notes',
} as const;

/**
 * Accrual type option set values in Dataverse
 * These must match the option set values defined in Dataverse
 */
export const ACCRUAL_TYPE_VALUES: Record<AccrualType, number> = {
  'Månedlig optjening': 100000000, // Monthly Accrual
  'Årsstart overførsel': 100000001, // Year Start Carryover (transfer IN)
  'Manuel justering': 100000002, // Manual Adjustment
  'Startsaldo': 100000003, // Initial Balance
};

/**
 * Helper function to get current holiday year string
 * Holiday year runs Sept 1 - Aug 31
 * @param date Optional date to calculate for (defaults to now)
 * @returns Holiday year string, e.g., "2024-2025"
 */
export function getHolidayYear(date?: Date): string {
  const d = date || new Date();
  const month = d.getMonth() + 1; // 1-indexed
  const year = d.getFullYear();

  // If September or later, holiday year is currentYear-nextYear
  // If before September, holiday year is previousYear-currentYear
  if (month >= DANISH_HOLIDAY_CONSTANTS.HOLIDAY_YEAR_START_MONTH) {
    return `${year}-${year + 1}`;
  } else {
    return `${year - 1}-${year}`;
  }
}

/**
 * Helper function to get the start date of a holiday year
 * @param holidayYear Holiday year string, e.g., "2024-2025"
 * @returns Start date (September 1st of the first year)
 */
export function getHolidayYearStartDate(holidayYear: string): Date {
  const startYear = parseInt(holidayYear.split('-')[0], 10);
  return new Date(startYear, 8, 1); // September 1st (month is 0-indexed)
}

/**
 * Helper function to get the end date of a holiday year
 * @param holidayYear Holiday year string, e.g., "2024-2025"
 * @returns End date (August 31st of the second year)
 */
export function getHolidayYearEndDate(holidayYear: string): Date {
  const endYear = parseInt(holidayYear.split('-')[1], 10);
  return new Date(endYear, 7, 31); // August 31st (month is 0-indexed)
}

/**
 * Calculate available days from a holiday balance record
 * Includes transferred-in days from previous year
 */
export function calculateAvailableDays(balance: IHolidayBalance): number {
  const transferredIn = balance.TransferredInDays || 0;
  const carriedOver = balance.CarriedOverDays || 0; // Legacy support
  return balance.AccruedDays + carriedOver + transferredIn - balance.UsedDays - balance.PendingDays;
}

/**
 * Calculate available feriefridage from a holiday balance record
 */
export function calculateAvailableFeriefridage(balance: IHolidayBalance): number {
  const transferredIn = balance.FeriefridageTransferredIn || 0;
  return balance.FeriefridageAccrued + transferredIn - balance.FeriefridageUsed;
}

/**
 * Get the end date of the holiday-taking period (ferieafholdelsesperiode)
 * This is December 31st of the year following the holiday year start
 * Example: For holiday year 2024-2025, period ends December 31, 2025
 * @param holidayYear Holiday year string, e.g., "2024-2025"
 * @returns End date of holiday-taking period (December 31st)
 */
export function getHolidayTakingPeriodEndDate(holidayYear: string): Date {
  const endYear = parseInt(holidayYear.split('-')[1], 10);
  return new Date(
    endYear,
    DANISH_HOLIDAY_CONSTANTS.HOLIDAY_TAKING_PERIOD_END_MONTH - 1,
    DANISH_HOLIDAY_CONSTANTS.HOLIDAY_TAKING_PERIOD_END_DAY
  );
}

/**
 * Get the transfer agreement deadline date for a holiday year
 * This is December 31st at the end of the holiday-taking period
 * @param holidayYear Holiday year string, e.g., "2024-2025"
 * @returns Transfer deadline date (December 31st)
 */
export function getTransferDeadline(holidayYear: string): Date {
  return getHolidayTakingPeriodEndDate(holidayYear);
}

/**
 * Check if we're still within the transfer agreement deadline
 * @param holidayYear Holiday year string
 * @param currentDate Optional current date (defaults to now)
 * @returns True if transfer agreement can still be made
 */
export function isWithinTransferDeadline(holidayYear: string, currentDate?: Date): boolean {
  const deadline = getTransferDeadline(holidayYear);
  const now = currentDate || new Date();
  return now <= deadline;
}

/**
 * Calculate how many days can be transferred from this holiday year
 * Only the 5th week (days beyond 20) can be transferred
 *
 * @param balance The holiday balance record
 * @returns Maximum days that can be transferred (0-5)
 */
export function calculateMaxTransferableDays(balance: IHolidayBalance): number {
  const { MANDATORY_VACATION_DAYS, MAX_TRANSFER_DAYS_PER_YEAR } = DANISH_HOLIDAY_CONSTANTS;

  // Total available = accrued + transferred in from previous years
  const totalDays = balance.AccruedDays + (balance.TransferredInDays || 0) + (balance.CarriedOverDays || 0);

  // Days remaining after usage
  const remainingDays = totalDays - balance.UsedDays - balance.PendingDays;

  // Only days beyond the mandatory 20 can be transferred
  const daysAboveMandatory = Math.max(0, remainingDays - MANDATORY_VACATION_DAYS);

  // But maximum is 5 days per year (new transfers)
  // Note: Previously transferred days that haven't been used can be transferred again
  return Math.min(daysAboveMandatory, MAX_TRANSFER_DAYS_PER_YEAR);
}

/**
 * Check if an employee has met the mandatory 4-week (20 days) requirement
 * First 4 weeks MUST be taken - they cannot be transferred
 *
 * @param balance The holiday balance record
 * @returns True if mandatory vacation requirement is met
 */
export function hasMandatoryVacationBeenTaken(balance: IHolidayBalance): boolean {
  const { MANDATORY_VACATION_DAYS } = DANISH_HOLIDAY_CONSTANTS;
  return balance.UsedDays >= MANDATORY_VACATION_DAYS;
}

/**
 * Calculate days that will be forfeited if not taken by deadline
 * First 4 weeks (20 days) are lost if not taken
 * 5th week can be paid out or transferred with agreement
 *
 * @param balance The holiday balance record
 * @returns Number of days at risk of forfeiture
 */
export function calculateDaysAtRiskOfForfeiture(balance: IHolidayBalance): number {
  const { MANDATORY_VACATION_DAYS } = DANISH_HOLIDAY_CONSTANTS;

  const totalAvailable = calculateAvailableDays(balance);

  // If using less than 20 days, the unused portion of mandatory days will be lost
  const mandatoryDaysRemaining = Math.max(0, MANDATORY_VACATION_DAYS - balance.UsedDays);

  // The 5th week can be transferred (with agreement) or paid out, so it's not "forfeited"
  return Math.min(mandatoryDaysRemaining, totalAvailable);
}

/**
 * Validate a transfer request
 * @param balance Current balance
 * @param daysToTransfer Number of days requested to transfer
 * @param transferDate Date of transfer agreement
 * @returns Validation result with error message if invalid
 */
export function validateTransferRequest(
  balance: IHolidayBalance,
  daysToTransfer: number,
  transferDate: Date
): { isValid: boolean; errorMessage?: string } {
  const { MAX_TRANSFER_DAYS_PER_YEAR, MANDATORY_VACATION_DAYS } = DANISH_HOLIDAY_CONSTANTS;

  // Check deadline
  if (!isWithinTransferDeadline(balance.HolidayYear, transferDate)) {
    return {
      isValid: false,
      errorMessage: `Transfer agreement must be made by December 31. Deadline has passed for ${balance.HolidayYear}.`,
    };
  }

  // Check max transfer per year
  if (daysToTransfer > MAX_TRANSFER_DAYS_PER_YEAR) {
    return {
      isValid: false,
      errorMessage: `Maximum ${MAX_TRANSFER_DAYS_PER_YEAR} days can be transferred per year.`,
    };
  }

  // Check if there are enough transferable days
  const maxTransferable = calculateMaxTransferableDays(balance);
  if (daysToTransfer > maxTransferable) {
    return {
      isValid: false,
      errorMessage: `Only ${maxTransferable} days are available for transfer. The first 20 days must be taken.`,
    };
  }

  // Check mandatory vacation has been taken
  const usedPlusTransfer = balance.UsedDays + balance.PendingDays;
  const remainingAfterTransfer = balance.AccruedDays + (balance.TransferredInDays || 0) - usedPlusTransfer - daysToTransfer;
  if (remainingAfterTransfer > 0 && balance.UsedDays < MANDATORY_VACATION_DAYS) {
    // Warning: Employee hasn't taken mandatory 20 days yet
    // This is a warning, not necessarily an error
  }

  return { isValid: true };
}

/**
 * Process year-end transfer for a holiday balance
 * Creates the transfer-out from current year and transfer-in for next year
 *
 * @param currentBalance Current year's balance
 * @param daysToTransfer Number of days to transfer
 * @param agreementDate Date of transfer agreement
 * @returns Updated balances for current and next year
 */
export interface IYearEndTransferResult {
  updatedCurrentBalance: Partial<IHolidayBalance>;
  newYearTransferIn: {
    holidayYear: string;
    transferredInDays: number;
    transferSource: string; // e.g., "2024-2025"
  };
}

export function prepareYearEndTransfer(
  currentBalance: IHolidayBalance,
  daysToTransfer: number,
  agreementDate: Date
): IYearEndTransferResult {
  // Parse current holiday year to get next year
  const [startYear, endYear] = currentBalance.HolidayYear.split('-').map(Number);
  const nextHolidayYear = `${endYear}-${endYear + 1}`;

  return {
    updatedCurrentBalance: {
      TransferredOutDays: daysToTransfer,
      HasTransferAgreement: true,
      TransferAgreementDate: agreementDate,
    },
    newYearTransferIn: {
      holidayYear: nextHolidayYear,
      transferredInDays: daysToTransfer,
      transferSource: currentBalance.HolidayYear,
    },
  };
}
