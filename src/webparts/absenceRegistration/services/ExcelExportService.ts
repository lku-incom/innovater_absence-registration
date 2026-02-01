/**
 * Excel Export Service
 * Generates CSV files that can be opened in Excel
 */

import { IAbsenceRegistration } from '../models/IAbsenceRegistration';
import { IAccrualHistory, ICalculatedHolidayBalance } from '../models/IHolidayBalance';

export class ExcelExportService {
  /**
   * Export user's holiday data to Excel (CSV format)
   */
  public static exportUserHolidayData(
    employeeName: string,
    balance: ICalculatedHolidayBalance | undefined,
    registrations: IAbsenceRegistration[],
    accrualHistory: IAccrualHistory[]
  ): void {
    // Generate CSV content with all data in one file
    const csv = this.generateCombinedCsv(employeeName, balance, registrations, accrualHistory);
    const filename = `Feriedata_${employeeName.replace(/\s+/g, '_')}_${this.formatDateForFilename(new Date())}.csv`;

    // Add BOM for Excel to recognize UTF-8
    const bom = '\uFEFF';
    this.downloadFile(bom + csv, filename, 'text/csv;charset=utf-8');
  }

  /**
   * Generate combined CSV with all sections
   */
  private static generateCombinedCsv(
    employeeName: string,
    balance: ICalculatedHolidayBalance | undefined,
    registrations: IAbsenceRegistration[],
    accrualHistory: IAccrualHistory[]
  ): string {
    const lines: string[] = [];
    const sep = ';'; // Use semicolon for better Excel compatibility in Danish locale

    // Section 1: Summary
    lines.push('FERIESALDO');
    lines.push(`Medarbejder${sep}${this.escapeCsv(employeeName)}`);
    lines.push(`Eksporteret${sep}${this.formatDateTime(new Date())}`);

    if (balance) {
      lines.push(`Ferieår${sep}${balance.HolidayYear}`);
      lines.push('');
      lines.push('FERIEDAGE');
      lines.push(`Beskrivelse${sep}Dage`);
      lines.push(`Optjent${sep}${this.formatNumber(balance.TotalAccruedDays)}`);
      lines.push(`Brugt (godkendt)${sep}${this.formatNumber(balance.UsedDays)}`);
      lines.push(`Afventer godkendelse${sep}${this.formatNumber(balance.PendingDays)}`);
      lines.push(`Til rådighed${sep}${this.formatNumber(balance.AvailableDays)}`);
      lines.push('');
      lines.push('FERIEFRIDAGE');
      lines.push(`Beskrivelse${sep}Dage`);
      lines.push(`Optjent${sep}${this.formatNumber(balance.TotalAccruedFeriefridage)}`);
      lines.push(`Brugt (godkendt)${sep}${this.formatNumber(balance.UsedFeriefridage)}`);
      lines.push(`Afventer godkendelse${sep}${this.formatNumber(balance.PendingFeriefridage)}`);
      lines.push(`Til rådighed${sep}${this.formatNumber(balance.AvailableFeriefridage)}`);
    } else {
      lines.push('Ingen feriesaldo fundet');
    }

    // Section 2: Absence Registrations
    lines.push('');
    lines.push('');
    lines.push('FRAVÆRSREGISTRERINGER');
    lines.push([
      'Fraværstype',
      'Startdato',
      'Slutdato',
      'Antal dage',
      'Status',
      'Godkender',
      'Noter',
      'Oprettet'
    ].join(sep));

    if (registrations.length > 0) {
      for (const reg of registrations) {
        lines.push([
          this.escapeCsv(reg.AbsenceType || ''),
          reg.StartDate ? this.formatDate(reg.StartDate) : '',
          reg.EndDate ? this.formatDate(reg.EndDate) : '',
          this.formatNumber(reg.NumberOfDays),
          this.escapeCsv(reg.Status || ''),
          this.escapeCsv(reg.ApproverName || ''),
          this.escapeCsv(reg.Notes || ''),
          reg.Created ? this.formatDateTime(reg.Created) : ''
        ].join(sep));
      }
    } else {
      lines.push('Ingen fraværsregistreringer fundet');
    }

    // Section 3: Accrual History
    lines.push('');
    lines.push('');
    lines.push('OPTJENINGSHISTORIK');
    lines.push([
      'Dato',
      'Type',
      'Ferieår',
      'Feriedage',
      'Feriefridage',
      'Noter'
    ].join(sep));

    if (accrualHistory.length > 0) {
      for (const accrual of accrualHistory) {
        lines.push([
          accrual.AccrualDate ? this.formatDate(accrual.AccrualDate) : '',
          this.escapeCsv(accrual.AccrualType || ''),
          this.escapeCsv(accrual.HolidayYear || ''),
          this.formatNumber(accrual.DaysAccrued),
          this.formatNumber(accrual.FeriefridageAccrued || 0),
          this.escapeCsv(accrual.Notes || '')
        ].join(sep));
      }
    } else {
      lines.push('Ingen optjeningshistorik fundet');
    }

    return lines.join('\r\n');
  }

  /**
   * Download file to user's computer
   */
  private static downloadFile(content: string, filename: string, mimeType: string): void {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
  }

  /**
   * Escape CSV value (handle quotes and special characters)
   */
  private static escapeCsv(value: string): string {
    if (!value) return '';
    // If value contains semicolon, quote, or newline, wrap in quotes and escape quotes
    if (value.indexOf(';') >= 0 || value.indexOf('"') >= 0 || value.indexOf('\n') >= 0 || value.indexOf('\r') >= 0) {
      return '"' + value.replace(/"/g, '""') + '"';
    }
    return value;
  }

  /**
   * Format number for CSV (use comma as decimal separator for Danish locale)
   */
  private static formatNumber(num: number): string {
    if (num === undefined || num === null) return '0';
    // Use comma as decimal separator for Danish Excel
    return num.toFixed(1).replace('.', ',');
  }

  /**
   * Format date for display (dd-mm-yyyy)
   */
  private static formatDate(date: Date | string): string {
    const d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) return '';

    const day = this.padZero(d.getDate(), 2);
    const month = this.padZero(d.getMonth() + 1, 2);
    const year = d.getFullYear();
    return day + '-' + month + '-' + year;
  }

  /**
   * Format date and time for display
   */
  private static formatDateTime(date: Date | string): string {
    const d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) return '';

    const day = this.padZero(d.getDate(), 2);
    const month = this.padZero(d.getMonth() + 1, 2);
    const year = d.getFullYear();
    const hours = this.padZero(d.getHours(), 2);
    const minutes = this.padZero(d.getMinutes(), 2);
    return day + '-' + month + '-' + year + ' ' + hours + ':' + minutes;
  }

  /**
   * Format date for filename (yyyy-mm-dd)
   */
  private static formatDateForFilename(date: Date): string {
    const year = date.getFullYear();
    const month = this.padZero(date.getMonth() + 1, 2);
    const day = this.padZero(date.getDate(), 2);
    return year + '-' + month + '-' + day;
  }

  /**
   * Pad string with leading zeros (ES5 compatible)
   */
  private static padZero(num: number, length: number): string {
    let result = String(num);
    while (result.length < length) {
      result = '0' + result;
    }
    return result;
  }
}

export default ExcelExportService;
