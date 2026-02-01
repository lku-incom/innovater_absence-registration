/**
 * Service for calculating Danish public holidays and working days
 *
 * Danish public holidays include fixed dates and moveable feasts based on Easter.
 * Note: Store Bededag was abolished as a public holiday starting 2024.
 */

export interface IDanishHoliday {
  date: Date;
  name: string;
}

export class DanishHolidayService {
  /**
   * Calculate Easter Sunday for a given year using the Anonymous Gregorian algorithm
   */
  private static calculateEaster(year: number): Date {
    const a = year % 19;
    const b = Math.floor(year / 100);
    const c = year % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31);
    const day = ((h + l - 7 * m + 114) % 31) + 1;

    return new Date(year, month - 1, day);
  }

  /**
   * Add days to a date
   */
  private static addDays(date: Date, days: number): Date {
    const result = new Date(date.getTime());
    result.setDate(result.getDate() + days);
    return result;
  }

  /**
   * Get all Danish public holidays for a given year
   */
  public static getHolidaysForYear(year: number): IDanishHoliday[] {
    const holidays: IDanishHoliday[] = [];
    const easter = this.calculateEaster(year);

    // Fixed holidays
    holidays.push({ date: new Date(year, 0, 1), name: 'Nytårsdag' }); // Jan 1
    holidays.push({ date: new Date(year, 11, 24), name: 'Juleaftensdag' }); // Dec 24
    holidays.push({ date: new Date(year, 11, 25), name: '1. Juledag' }); // Dec 25
    holidays.push({ date: new Date(year, 11, 26), name: '2. Juledag' }); // Dec 26

    // Easter-based moveable feasts
    holidays.push({
      date: this.addDays(easter, -3),
      name: 'Skærtorsdag', // Maundy Thursday
    });
    holidays.push({
      date: this.addDays(easter, -2),
      name: 'Langfredag', // Good Friday
    });
    holidays.push({
      date: easter,
      name: 'Påskedag', // Easter Sunday
    });
    holidays.push({
      date: this.addDays(easter, 1),
      name: '2. Påskedag', // Easter Monday
    });

    // Store Bededag (Great Prayer Day) - ABOLISHED from 2024
    // Only include for years before 2024
    if (year < 2024) {
      holidays.push({
        date: this.addDays(easter, 26),
        name: 'Store Bededag',
      });
    }

    holidays.push({
      date: this.addDays(easter, 39),
      name: 'Kristi Himmelfartsdag', // Ascension Day
    });
    holidays.push({
      date: this.addDays(easter, 49),
      name: 'Pinsedag', // Whit Sunday
    });
    holidays.push({
      date: this.addDays(easter, 50),
      name: '2. Pinsedag', // Whit Monday
    });

    return holidays;
  }

  /**
   * Check if a date is a Danish public holiday
   */
  public static isHoliday(date: Date): boolean {
    const year = date.getFullYear();
    const holidays = this.getHolidaysForYear(year);

    return holidays.some(
      (holiday) =>
        holiday.date.getFullYear() === date.getFullYear() &&
        holiday.date.getMonth() === date.getMonth() &&
        holiday.date.getDate() === date.getDate()
    );
  }

  /**
   * Check if a date is a weekend (Saturday or Sunday)
   */
  public static isWeekend(date: Date): boolean {
    const day = date.getDay();
    return day === 0 || day === 6; // Sunday = 0, Saturday = 6
  }

  /**
   * Check if a date is a working day (not weekend and not holiday)
   */
  public static isWorkingDay(date: Date): boolean {
    return !this.isWeekend(date) && !this.isHoliday(date);
  }

  /**
   * Calculate the number of working days between two dates (inclusive)
   * Excludes weekends and Danish public holidays
   */
  public static calculateWorkingDays(startDate: Date, endDate: Date): number {
    if (startDate > endDate) {
      return 0;
    }

    let workingDays = 0;

    // Create dates using only year, month, day to avoid timezone issues
    // This ensures we're working with the local date the user selected
    const start = new Date(
      startDate.getFullYear(),
      startDate.getMonth(),
      startDate.getDate(),
      12, 0, 0, 0  // Use noon to avoid any DST edge cases
    );

    const end = new Date(
      endDate.getFullYear(),
      endDate.getMonth(),
      endDate.getDate(),
      12, 0, 0, 0
    );

    const currentDate = new Date(start.getTime());

    while (currentDate <= end) {
      if (this.isWorkingDay(currentDate)) {
        workingDays++;
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    return workingDays;
  }

  /**
   * Get holidays between two dates
   */
  public static getHolidaysBetweenDates(
    startDate: Date,
    endDate: Date
  ): IDanishHoliday[] {
    const holidays: IDanishHoliday[] = [];
    const startYear = startDate.getFullYear();
    const endYear = endDate.getFullYear();

    for (let year = startYear; year <= endYear; year++) {
      const yearHolidays = this.getHolidaysForYear(year);
      yearHolidays.forEach((holiday) => {
        if (holiday.date >= startDate && holiday.date <= endDate) {
          holidays.push(holiday);
        }
      });
    }

    return holidays;
  }
}

export default DanishHolidayService;
