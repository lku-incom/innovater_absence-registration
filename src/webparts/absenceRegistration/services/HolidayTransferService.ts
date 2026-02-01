/**
 * Service for handling Danish Holiday Transfer operations
 *
 * Based on Ferieloven (Danish Holiday Act) §19:
 * - Only the 5th week (5 days beyond mandatory 20) can be transferred per year
 * - Written agreement required by December 31
 * - First 4 weeks (20 days) MUST be taken or are forfeited
 * - Feriefridage rules depend on company policy (not covered by Ferieloven)
 */

import {
  IHolidayBalance,
  IAccrualHistory,
  DANISH_HOLIDAY_CONSTANTS,
  calculateMaxTransferableDays,
  validateTransferRequest,
  prepareYearEndTransfer,
  getTransferDeadline,
  isWithinTransferDeadline,
  hasMandatoryVacationBeenTaken,
  calculateDaysAtRiskOfForfeiture,
  getHolidayYear,
  IYearEndTransferResult,
  ACCRUAL_TYPE_VALUES,
} from '../models/IHolidayBalance';

export interface ITransferRequest {
  employeeEmail: string;
  employeeName: string;
  currentHolidayYear: string;
  feriedageToTransfer: number;
  feriefridageToTransfer?: number; // Optional - depends on company policy
  agreementDate: Date;
  notes?: string;
}

export interface ITransferValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  maxFeriedageTransferable: number;
  deadlineDate: Date;
  isDeadlinePassed: boolean;
  mandatoryDaysTaken: number;
  daysAtRisk: number;
}

export interface IYearEndProcessingResult {
  success: boolean;
  error?: string;
  updatedCurrentBalance?: IHolidayBalance;
  newYearBalance?: Partial<IHolidayBalance>;
  transferAccrualRecord?: Partial<IAccrualHistory>;
}

export class HolidayTransferService {
  /**
   * Validate a transfer request before processing
   * @param balance Current holiday balance
   * @param request Transfer request details
   * @returns Validation result with errors and warnings
   */
  public static validateTransfer(
    balance: IHolidayBalance,
    request: ITransferRequest
  ): ITransferValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];
    const { MAX_TRANSFER_DAYS_PER_YEAR, MANDATORY_VACATION_DAYS } = DANISH_HOLIDAY_CONSTANTS;

    const deadline = getTransferDeadline(balance.HolidayYear);
    const isDeadlinePassed = !isWithinTransferDeadline(balance.HolidayYear, request.agreementDate);
    const maxTransferable = calculateMaxTransferableDays(balance);
    const mandatoryTaken = hasMandatoryVacationBeenTaken(balance);
    const daysAtRisk = calculateDaysAtRiskOfForfeiture(balance);

    // Check deadline
    if (isDeadlinePassed) {
      errors.push(
        `Transfer deadline has passed. Agreement must be made by December 31, ${deadline.getFullYear()}.`
      );
    }

    // Check feriedage transfer amount
    if (request.feriedageToTransfer > MAX_TRANSFER_DAYS_PER_YEAR) {
      errors.push(
        `Cannot transfer more than ${MAX_TRANSFER_DAYS_PER_YEAR} feriedage per year (requested: ${request.feriedageToTransfer}).`
      );
    }

    if (request.feriedageToTransfer > maxTransferable) {
      errors.push(
        `Only ${maxTransferable} days are available for transfer. First ${MANDATORY_VACATION_DAYS} days must be taken.`
      );
    }

    if (request.feriedageToTransfer < 0) {
      errors.push('Transfer amount cannot be negative.');
    }

    // Warnings for mandatory vacation
    if (!mandatoryTaken) {
      warnings.push(
        `Employee has not taken mandatory ${MANDATORY_VACATION_DAYS} vacation days yet. ` +
        `${daysAtRisk} days may be forfeited if not taken by deadline.`
      );
    }

    // Warning for feriefridage transfer
    if (request.feriefridageToTransfer && request.feriefridageToTransfer > 0) {
      warnings.push(
        'Feriefridage transfer is not covered by Ferieloven. ' +
        'Please verify company policy allows this transfer.'
      );
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      maxFeriedageTransferable: maxTransferable,
      deadlineDate: deadline,
      isDeadlinePassed,
      mandatoryDaysTaken: balance.UsedDays,
      daysAtRisk,
    };
  }

  /**
   * Process a year-end transfer
   * Creates the necessary updates for current and next year balances
   *
   * @param currentBalance Current year's balance
   * @param request Transfer request
   * @returns Processing result with updated balances
   */
  public static processTransfer(
    currentBalance: IHolidayBalance,
    request: ITransferRequest
  ): IYearEndProcessingResult {
    // Validate first
    const validation = this.validateTransfer(currentBalance, request);
    if (!validation.isValid) {
      return {
        success: false,
        error: validation.errors.join(' '),
      };
    }

    // Prepare transfer
    const transferResult = prepareYearEndTransfer(
      currentBalance,
      request.feriedageToTransfer,
      request.agreementDate
    );

    // Parse holiday years
    const [, endYear] = currentBalance.HolidayYear.split('-').map(Number);
    const nextHolidayYear = `${endYear}-${endYear + 1}`;

    // Create updated current balance
    const updatedCurrentBalance: IHolidayBalance = {
      ...currentBalance,
      TransferredOutDays: request.feriedageToTransfer,
      HasTransferAgreement: true,
      TransferAgreementDate: request.agreementDate,
      FeriefridageTransferredOut: request.feriefridageToTransfer || 0,
    };

    // Create new year balance template (needs to be merged with existing if it exists)
    const newYearBalance: Partial<IHolidayBalance> = {
      EmployeeEmail: currentBalance.EmployeeEmail,
      EmployeeName: currentBalance.EmployeeName,
      HolidayYear: nextHolidayYear,
      TransferredInDays: request.feriedageToTransfer,
      FeriefridageTransferredIn: request.feriefridageToTransfer || 0,
      // Other fields would be initialized or merged as needed
    };

    // Create accrual history record for the transfer
    const transferAccrualRecord: Partial<IAccrualHistory> = {
      EmployeeEmail: currentBalance.EmployeeEmail,
      EmployeeName: currentBalance.EmployeeName,
      HolidayYear: nextHolidayYear,
      AccrualDate: request.agreementDate,
      AccrualMonth: request.agreementDate.getMonth() + 1,
      AccrualYear: request.agreementDate.getFullYear(),
      DaysAccrued: request.feriedageToTransfer,
      FeriefridageAccrued: request.feriefridageToTransfer || 0,
      AccrualType: 'Årsstart overførsel',
      Notes: request.notes || `Transferred from ${currentBalance.HolidayYear}`,
    };

    return {
      success: true,
      updatedCurrentBalance,
      newYearBalance,
      transferAccrualRecord,
    };
  }

  /**
   * Calculate the year-end summary for an employee
   * Shows what will happen at the end of the holiday-taking period
   *
   * @param balance Current balance
   * @returns Summary of year-end situation
   */
  public static calculateYearEndSummary(balance: IHolidayBalance): {
    totalAvailable: number;
    mandatoryDaysRemaining: number;
    daysToBeForfeited: number;
    daysTransferableWithAgreement: number;
    daysToBePaidOut: number;
    feriefridageRemaining: number;
    recommendations: string[];
  } {
    const { MANDATORY_VACATION_DAYS, MAX_TRANSFER_DAYS_PER_YEAR } = DANISH_HOLIDAY_CONSTANTS;

    const totalAvailable =
      balance.AccruedDays +
      (balance.TransferredInDays || 0) +
      (balance.CarriedOverDays || 0) -
      balance.UsedDays -
      balance.PendingDays;

    const mandatoryDaysRemaining = Math.max(0, MANDATORY_VACATION_DAYS - balance.UsedDays);
    const daysAboveMandatory = Math.max(0, totalAvailable - MANDATORY_VACATION_DAYS);

    // Days that will be forfeited (mandatory days not taken)
    const daysToBeForfeited = Math.min(mandatoryDaysRemaining, totalAvailable);

    // Days that can be transferred (5th week, max 5 per year)
    const daysTransferableWithAgreement = Math.min(
      daysAboveMandatory,
      MAX_TRANSFER_DAYS_PER_YEAR
    );

    // Days that will be paid out (anything above 4 weeks that's not transferred)
    const daysToBePaidOut = balance.HasTransferAgreement
      ? Math.max(0, daysAboveMandatory - (balance.TransferredOutDays || 0))
      : daysAboveMandatory;

    // Feriefridage remaining
    const feriefridageRemaining =
      balance.FeriefridageAccrued +
      (balance.FeriefridageTransferredIn || 0) -
      balance.FeriefridageUsed -
      (balance.FeriefridageTransferredOut || 0);

    // Generate recommendations
    const recommendations: string[] = [];

    if (mandatoryDaysRemaining > 0) {
      recommendations.push(
        `You need to take ${mandatoryDaysRemaining} more mandatory vacation days to avoid forfeiture.`
      );
    }

    if (daysTransferableWithAgreement > 0 && !balance.HasTransferAgreement) {
      recommendations.push(
        `You can transfer up to ${daysTransferableWithAgreement} days to next year with a written agreement by December 31.`
      );
    }

    if (daysToBePaidOut > 0 && !balance.HasTransferAgreement) {
      recommendations.push(
        `${daysToBePaidOut} days (5th week) will be paid out if no transfer agreement is made.`
      );
    }

    if (feriefridageRemaining > 0) {
      recommendations.push(
        `${feriefridageRemaining} feriefridage remaining. Check company policy for transfer/payout rules.`
      );
    }

    return {
      totalAvailable,
      mandatoryDaysRemaining,
      daysToBeForfeited,
      daysTransferableWithAgreement,
      daysToBePaidOut,
      feriefridageRemaining,
      recommendations,
    };
  }

  /**
   * Check if we're approaching the transfer deadline
   * Useful for sending reminders
   *
   * @param holidayYear Holiday year to check
   * @param daysBeforeDeadline Days before deadline to trigger warning
   * @returns Whether we're in the warning period
   */
  public static isApproachingTransferDeadline(
    holidayYear: string,
    daysBeforeDeadline: number = 30
  ): boolean {
    const deadline = getTransferDeadline(holidayYear);
    const warningDate = new Date(deadline.getTime());
    warningDate.setDate(warningDate.getDate() - daysBeforeDeadline);

    const now = new Date();
    return now >= warningDate && now <= deadline;
  }

  /**
   * Get the previous holiday year string
   */
  public static getPreviousHolidayYear(holidayYear: string): string {
    const [startYear] = holidayYear.split('-').map(Number);
    return `${startYear - 1}-${startYear}`;
  }

  /**
   * Get the next holiday year string
   */
  public static getNextHolidayYear(holidayYear: string): string {
    const [, endYear] = holidayYear.split('-').map(Number);
    return `${endYear}-${endYear + 1}`;
  }
}

export default HolidayTransferService;
