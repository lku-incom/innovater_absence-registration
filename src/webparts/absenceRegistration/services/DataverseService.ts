/**
 * Service for Dataverse (Power Apps tables) operations
 * Handles CRUD operations for absence registrations using Dataverse Web API
 */

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IAbsenceRegistration,
  AbsenceType,
  RegistrationStatus,
} from '../models/IAbsenceRegistration';
import {
  IAccrualHistory,
  ICalculatedHolidayBalance,
  AccrualType,
  getHolidayYear,
  getHolidayYearStartDate,
  getHolidayYearEndDate,
  ACCRUAL_TYPE_VALUES,
} from '../models/IHolidayBalance';

// Dataverse table name (logical name) - uses publisher prefix cr153_
// Note: Dataverse API uses plural form (entity set name)
const TABLE_NAME = 'cr153_absenceregistrations';
const ACCRUAL_HISTORY_TABLE = 'cr_accrualhistories';

// Dataverse environment URL - Innovater
const DEFAULT_DATAVERSE_URL = 'https://orgab6f6874.crm4.dynamics.com';

interface IDataverseEntity {
  cr153_absenceregistrationid?: string;
  cr153_name?: string; // Primary name column (Display name: Titel)
  cr153_employeeemail: string;
  cr153_employeename: string;
  cr153_approveremail: string;
  cr153_approvername: string;
  cr153_startdate: string;
  cr153_enddate: string;
  cr153_numberofdays: number;
  cr153_absencetype: number; // Option set value
  cr153_notes?: string;
  cr153_status: number; // Option set value
  cr153_approvaldate?: string;
  cr153_approvercomments?: string;
  createdon?: string;
  modifiedon?: string;
}

interface IAccrualHistoryEntity {
  cr_accrualhistoryid?: string;
  cr_name?: string;
  cr_employeeemail: string;
  cr_employeename: string;
  cr_holidayyear: string;
  cr_accrualdate: string;
  cr_accrualmonth: number;
  cr_accrualyear: number;
  cr_daysaccrued: number;
  cr_feriefridageaccrued?: number;
  cr_balanceafteraccrual?: number;
  cr_accrualtype: number; // Option set value
  cr_notes?: string;
  createdon?: string;
}

// Mapping for AbsenceType to Dataverse option set values
const AbsenceTypeOptionSet: Record<AbsenceType, number> = {
  Ferie: 100000000,
  Sygdom: 100000001,
  Barselsorlov: 100000002,
  Feriefridage: 100000003,
  'Flex/afspadsering': 100000004,
  'Andet fravær': 100000005,
};

const AbsenceTypeFromOptionSet: Record<number, AbsenceType> = {
  100000000: 'Ferie',
  100000001: 'Sygdom',
  100000002: 'Barselsorlov',
  100000003: 'Feriefridage',
  100000004: 'Flex/afspadsering',
  100000005: 'Andet fravær',
};

// Mapping for Status to Dataverse option set values
const StatusOptionSet: Record<RegistrationStatus, number> = {
  Kladde: 100000000,
  'Afventer godkendelse': 100000001,
  Godkendt: 100000002,
  Afvist: 100000003,
};

const StatusFromOptionSet: Record<number, RegistrationStatus> = {
  100000000: 'Kladde',
  100000001: 'Afventer godkendelse',
  100000002: 'Godkendt',
  100000003: 'Afvist',
};

export class DataverseService {
  private static _instance: DataverseService;
  private _context: WebPartContext;
  private _aadClient: AadHttpClient | undefined;
  private _dataverseUrl: string;

  private constructor() {
    this._dataverseUrl = DEFAULT_DATAVERSE_URL;
  }

  /**
   * Initialize the Dataverse service with SPFx context
   */
  public static getInstance(context?: WebPartContext): DataverseService {
    if (!DataverseService._instance) {
      DataverseService._instance = new DataverseService();
    }

    if (context) {
      DataverseService._instance._context = context;
    }

    return DataverseService._instance;
  }

  /**
   * Set the Dataverse environment URL
   */
  public setDataverseUrl(url: string): void {
    this._dataverseUrl = url;
  }

  /**
   * Get the AAD HTTP Client for Dataverse
   */
  private async getClient(): Promise<AadHttpClient> {
    if (!this._aadClient) {
      this._aadClient = await this._context.aadHttpClientFactory.getClient(
        this._dataverseUrl
      );
    }
    return this._aadClient;
  }

  /**
   * Make a GET request to Dataverse
   */
  private async get<T>(endpoint: string): Promise<T> {
    const client = await this.getClient();
    const response: HttpClientResponse = await client.get(
      `${this._dataverseUrl}/api/data/v9.2/${endpoint}`,
      AadHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
      }
    );

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Dataverse GET fejl: ${error}`);
    }

    return response.json();
  }

  /**
   * Make a POST request to Dataverse
   */
  private async post<T>(endpoint: string, data: object): Promise<T> {
    const client = await this.getClient();
    const response: HttpClientResponse = await client.post(
      `${this._dataverseUrl}/api/data/v9.2/${endpoint}`,
      AadHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        body: JSON.stringify(data),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Dataverse POST fejl: ${error}`);
    }

    // For create operations, the ID is in the OData-EntityId header
    const entityId = response.headers.get('OData-EntityId');
    if (entityId) {
      const match = entityId.match(/\(([^)]+)\)/);
      if (match) {
        return { id: match[1] } as unknown as T;
      }
    }

    return response.json();
  }

  /**
   * Make a PATCH request to Dataverse (update)
   */
  private async patch(endpoint: string, data: object): Promise<void> {
    const client = await this.getClient();
    const response: HttpClientResponse = await client.fetch(
      `${this._dataverseUrl}/api/data/v9.2/${endpoint}`,
      AadHttpClient.configurations.v1,
      {
        method: 'PATCH',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        body: JSON.stringify(data),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Dataverse PATCH fejl: ${error}`);
    }
  }

  /**
   * Make a DELETE request to Dataverse
   */
  private async delete(endpoint: string): Promise<void> {
    const client = await this.getClient();
    const response: HttpClientResponse = await client.fetch(
      `${this._dataverseUrl}/api/data/v9.2/${endpoint}`,
      AadHttpClient.configurations.v1,
      {
        method: 'DELETE',
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
      }
    );

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Dataverse DELETE fejl: ${error}`);
    }
  }

  /**
   * Create a new absence registration
   */
  public async createRegistration(
    registration: Omit<IAbsenceRegistration, 'Id' | 'Created' | 'Modified'>
  ): Promise<IAbsenceRegistration> {
    const entity: Partial<IDataverseEntity> = {
      cr153_name: `${registration.EmployeeName} - ${registration.AbsenceType}`,
      cr153_employeeemail: registration.EmployeeEmail,
      cr153_employeename: registration.EmployeeName,
      cr153_approveremail: registration.ApproverEmail,
      cr153_approvername: registration.ApproverName,
      cr153_startdate: registration.StartDate.toISOString(),
      cr153_enddate: registration.EndDate.toISOString(),
      cr153_numberofdays: registration.NumberOfDays,
      cr153_absencetype: AbsenceTypeOptionSet[registration.AbsenceType],
      cr153_notes: registration.Notes || '',
      cr153_status: StatusOptionSet[registration.Status],
    };

    const result = await this.post<{ id: string }>(TABLE_NAME, entity);

    return {
      ...registration,
      Id: 0, // Dataverse uses GUIDs, but we'll use a numeric ID for compatibility
      Title: `${registration.EmployeeName} - ${registration.AbsenceType}`,
      Created: new Date(),
      Modified: new Date(),
    };
  }

  /**
   * Update an existing absence registration
   */
  public async updateRegistration(
    id: string,
    updates: Partial<IAbsenceRegistration>
  ): Promise<void> {
    const entity: Partial<IDataverseEntity> = {};

    if (updates.StartDate) {
      entity.cr153_startdate = updates.StartDate.toISOString();
    }
    if (updates.EndDate) {
      entity.cr153_enddate = updates.EndDate.toISOString();
    }
    if (updates.NumberOfDays !== undefined) {
      entity.cr153_numberofdays = updates.NumberOfDays;
    }
    if (updates.AbsenceType) {
      entity.cr153_absencetype = AbsenceTypeOptionSet[updates.AbsenceType];
    }
    if (updates.Notes !== undefined) {
      entity.cr153_notes = updates.Notes;
    }
    if (updates.Status) {
      entity.cr153_status = StatusOptionSet[updates.Status];
    }
    if (updates.ApprovalDate) {
      entity.cr153_approvaldate = updates.ApprovalDate.toISOString();
    }
    if (updates.ApproverComments !== undefined) {
      entity.cr153_approvercomments = updates.ApproverComments;
    }

    await this.patch(`${TABLE_NAME}(${id})`, entity);
  }

  /**
   * Delete an absence registration
   */
  public async deleteRegistration(id: string): Promise<void> {
    await this.delete(`${TABLE_NAME}(${id})`);
  }

  /**
   * Get all registrations for the current user
   */
  public async getMyRegistrations(
    userEmail: string
  ): Promise<IAbsenceRegistration[]> {
    // Direct email comparison (Dataverse doesn't support tolower)
    const filter = `cr153_employeeemail eq '${userEmail}'`;
    const orderby = 'cr153_startdate desc';
    const select =
      'cr153_absenceregistrationid,cr153_name,cr153_employeeemail,cr153_employeename,cr153_approveremail,cr153_approvername,cr153_startdate,cr153_enddate,cr153_numberofdays,cr153_absencetype,cr153_notes,cr153_status,cr153_approvaldate,cr153_approvercomments,createdon,modifiedon';

    const result = await this.get<{ value: IDataverseEntity[] }>(
      `${TABLE_NAME}?$filter=${encodeURIComponent(filter)}&$orderby=${orderby}&$select=${select}`
    );

    return result.value.map((entity) => this.mapEntityToRegistration(entity));
  }

  /**
   * Get registrations by status
   */
  public async getRegistrationsByStatus(
    userEmail: string,
    status: RegistrationStatus
  ): Promise<IAbsenceRegistration[]> {
    const statusValue = StatusOptionSet[status];
    // Direct email comparison (Dataverse doesn't support tolower)
    const filter = `cr153_employeeemail eq '${userEmail}' and cr153_status eq ${statusValue}`;
    const orderby = 'cr153_startdate desc';
    const select =
      'cr153_absenceregistrationid,cr153_name,cr153_employeeemail,cr153_employeename,cr153_approveremail,cr153_approvername,cr153_startdate,cr153_enddate,cr153_numberofdays,cr153_absencetype,cr153_notes,cr153_status,cr153_approvaldate,cr153_approvercomments,createdon,modifiedon';

    const result = await this.get<{ value: IDataverseEntity[] }>(
      `${TABLE_NAME}?$filter=${encodeURIComponent(filter)}&$orderby=${orderby}&$select=${select}`
    );

    return result.value.map((entity) => this.mapEntityToRegistration(entity));
  }

  /**
   * Get a single registration by ID
   */
  public async getRegistrationById(
    id: string
  ): Promise<IAbsenceRegistration | undefined> {
    try {
      const select =
        'cr153_absenceregistrationid,cr153_name,cr153_employeeemail,cr153_employeename,cr153_approveremail,cr153_approvername,cr153_startdate,cr153_enddate,cr153_numberofdays,cr153_absencetype,cr153_notes,cr153_status,cr153_approvaldate,cr153_approvercomments,createdon,modifiedon';

      const entity = await this.get<IDataverseEntity>(
        `${TABLE_NAME}(${id})?$select=${select}`
      );

      return this.mapEntityToRegistration(entity);
    } catch {
      return undefined;
    }
  }

  /**
   * Submit registration for approval
   */
  public async submitForApproval(id: string): Promise<void> {
    await this.updateRegistration(id, {
      Status: 'Afventer godkendelse',
    });
  }

  /**
   * Get registrations pending approval for the current approver
   */
  public async getPendingApprovals(
    approverEmail: string
  ): Promise<IAbsenceRegistration[]> {
    const statusValue = StatusOptionSet['Afventer godkendelse'];
    // Filter by approver email and pending status
    const filter = `cr153_approveremail eq '${approverEmail}' and cr153_status eq ${statusValue}`;
    const orderby = 'cr153_startdate asc';
    const select =
      'cr153_absenceregistrationid,cr153_name,cr153_employeeemail,cr153_employeename,cr153_approveremail,cr153_approvername,cr153_startdate,cr153_enddate,cr153_numberofdays,cr153_absencetype,cr153_notes,cr153_status,cr153_approvaldate,cr153_approvercomments,createdon,modifiedon';

    const result = await this.get<{ value: IDataverseEntity[] }>(
      `${TABLE_NAME}?$filter=${encodeURIComponent(filter)}&$orderby=${orderby}&$select=${select}`
    );

    return result.value.map((entity) => this.mapEntityToRegistration(entity));
  }

  /**
   * Approve a registration
   */
  public async approveRegistration(
    id: string,
    comments?: string
  ): Promise<void> {
    await this.updateRegistration(id, {
      Status: 'Godkendt',
      ApprovalDate: new Date(),
      ApproverComments: comments || '',
    });
  }

  /**
   * Reject a registration
   */
  public async rejectRegistration(
    id: string,
    comments: string
  ): Promise<void> {
    await this.updateRegistration(id, {
      Status: 'Afvist',
      ApprovalDate: new Date(),
      ApproverComments: comments,
    });
  }

  /**
   * Get all registrations (admin function)
   */
  public async getAllRegistrations(): Promise<IAbsenceRegistration[]> {
    const orderby = 'createdon desc';
    const select =
      'cr153_absenceregistrationid,cr153_name,cr153_employeeemail,cr153_employeename,cr153_approveremail,cr153_approvername,cr153_startdate,cr153_enddate,cr153_numberofdays,cr153_absencetype,cr153_notes,cr153_status,cr153_approvaldate,cr153_approvercomments,createdon,modifiedon';

    const result = await this.get<{ value: IDataverseEntity[] }>(
      `${TABLE_NAME}?$orderby=${orderby}&$select=${select}`
    );

    return result.value.map((entity) => this.mapEntityToRegistration(entity));
  }

  /**
   * Delete all registrations (admin function)
   * Warning: This is a destructive operation
   */
  public async deleteAllRegistrations(): Promise<number> {
    const registrations = await this.getAllRegistrations();
    let deletedCount = 0;

    for (const registration of registrations) {
      if (registration.DataverseId) {
        await this.deleteRegistration(registration.DataverseId);
        deletedCount++;
      }
    }

    return deletedCount;
  }

  /**
   * Get count of all accrual history records
   */
  public async getAccrualHistoryCount(): Promise<number> {
    try {
      // Dataverse doesn't support $top=0, so just fetch IDs and count
      const result = await this.get<{ value: { cr_accrualhistoryid: string }[] }>(
        `${ACCRUAL_HISTORY_TABLE}?$select=cr_accrualhistoryid`
      );
      return result.value.length;
    } catch {
      return 0;
    }
  }

  /**
   * Get accrual history records for a user for a specific holiday year
   * @param userEmail Employee email
   * @param holidayYear Holiday year (e.g., "2024-2025")
   * @returns Array of accrual history records
   */
  public async getAccrualHistoryForUser(
    userEmail: string,
    holidayYear: string
  ): Promise<IAccrualHistory[]> {
    try {
      const filter = `cr_employeeemail eq '${userEmail}' and cr_holidayyear eq '${holidayYear}'`;
      const select =
        'cr_accrualhistoryid,cr_name,cr_employeeemail,cr_employeename,cr_holidayyear,' +
        'cr_accrualdate,cr_accrualmonth,cr_accrualyear,cr_daysaccrued,' +
        'cr_feriefridageaccrued,cr_balanceafteraccrual,cr_accrualtype,cr_notes,createdon';
      const orderby = 'cr_accrualdate asc';

      const result = await this.get<{ value: IAccrualHistoryEntity[] }>(
        `${ACCRUAL_HISTORY_TABLE}?$filter=${encodeURIComponent(filter)}&$select=${select}&$orderby=${orderby}`
      );

      return result.value.map((entity) => this.mapEntityToAccrualHistory(entity));
    } catch {
      return [];
    }
  }

  /**
   * Calculate holiday balance for a user by aggregating accrual history and absence registrations
   * This is the client-side calculation approach - no stored balance needed
   *
   * @param userEmail Employee email
   * @param holidayYear Holiday year (defaults to current)
   * @returns Calculated holiday balance
   */
  public async calculateHolidayBalanceForUser(
    userEmail: string,
    holidayYear?: string
  ): Promise<ICalculatedHolidayBalance> {
    const year = holidayYear || getHolidayYear();
    const yearStartDate = getHolidayYearStartDate(year);
    const yearEndDate = getHolidayYearEndDate(year);

    // Fetch accrual history for this user/year
    const accrualHistory = await this.getAccrualHistoryForUser(userEmail, year);

    // Fetch all absence registrations for this user
    // Filter by date range (within holiday year) and type (Ferie or Feriefridage)
    const allRegistrations = await this.getMyRegistrations(userEmail);

    // Filter registrations to those within the holiday year and of relevant types
    const relevantRegistrations = allRegistrations.filter((reg) => {
      // Check if absence falls within holiday year (based on start date)
      // reg.StartDate is already a Date object from mapEntityToRegistration
      const absenceStart = reg.StartDate;
      const isWithinYear = absenceStart >= yearStartDate && absenceStart <= yearEndDate;

      // Only count Ferie and Feriefridage types
      const isRelevantType = reg.AbsenceType === 'Ferie' || reg.AbsenceType === 'Feriefridage';

      return isWithinYear && isRelevantType;
    });

    // Calculate accrued totals from accrual history
    let totalAccruedDays = 0;
    let totalAccruedFeriefridage = 0;

    for (const accrual of accrualHistory) {
      totalAccruedDays += accrual.DaysAccrued || 0;
      totalAccruedFeriefridage += accrual.FeriefridageAccrued || 0;
    }

    // Calculate used/pending from absence registrations
    let usedDays = 0;
    let pendingDays = 0;
    let usedFeriefridage = 0;
    let pendingFeriefridage = 0;

    for (const reg of relevantRegistrations) {
      if (reg.AbsenceType === 'Ferie') {
        if (reg.Status === 'Godkendt') {
          usedDays += reg.NumberOfDays;
        } else if (reg.Status === 'Afventer godkendelse') {
          pendingDays += reg.NumberOfDays;
        }
      } else if (reg.AbsenceType === 'Feriefridage') {
        if (reg.Status === 'Godkendt') {
          usedFeriefridage += reg.NumberOfDays;
        } else if (reg.Status === 'Afventer godkendelse') {
          pendingFeriefridage += reg.NumberOfDays;
        }
      }
    }

    // Get employee name from accrual history or registrations
    const employeeName =
      accrualHistory[0]?.EmployeeName ||
      relevantRegistrations[0]?.EmployeeName ||
      userEmail;

    return {
      EmployeeEmail: userEmail,
      EmployeeName: employeeName,
      HolidayYear: year,

      // Feriedage
      TotalAccruedDays: totalAccruedDays,
      UsedDays: usedDays,
      PendingDays: pendingDays,
      AvailableDays: totalAccruedDays - usedDays - pendingDays,

      // Feriefridage
      TotalAccruedFeriefridage: totalAccruedFeriefridage,
      UsedFeriefridage: usedFeriefridage,
      PendingFeriefridage: pendingFeriefridage,
      AvailableFeriefridage: totalAccruedFeriefridage - usedFeriefridage - pendingFeriefridage,

      // Source data counts
      AccrualHistoryCount: accrualHistory.length,
      AbsenceRegistrationCount: relevantRegistrations.length,
    };
  }

  /**
   * Map Dataverse entity to IAccrualHistory
   */
  private mapEntityToAccrualHistory(entity: IAccrualHistoryEntity): IAccrualHistory {
    // Reverse lookup for accrual type
    const accrualTypeFromValue: Record<number, AccrualType> = {
      100000000: 'Månedlig optjening',
      100000001: 'Årsstart overførsel',
      100000002: 'Manuel justering',
      100000003: 'Startsaldo',
    };

    return {
      DataverseId: entity.cr_accrualhistoryid,
      Name: entity.cr_name,
      EmployeeEmail: entity.cr_employeeemail,
      EmployeeName: entity.cr_employeename,
      HolidayYear: entity.cr_holidayyear,
      AccrualDate: new Date(entity.cr_accrualdate),
      AccrualMonth: entity.cr_accrualmonth,
      AccrualYear: entity.cr_accrualyear,
      DaysAccrued: entity.cr_daysaccrued || 0,
      FeriefridageAccrued: entity.cr_feriefridageaccrued,
      BalanceAfterAccrual: entity.cr_balanceafteraccrual,
      AccrualType: accrualTypeFromValue[entity.cr_accrualtype] || 'Manuel justering',
      Notes: entity.cr_notes,
      Created: entity.createdon ? new Date(entity.createdon) : undefined,
    };
  }

  /**
   * Delete all accrual history records (admin function)
   * Warning: This is a destructive operation
   */
  public async deleteAllAccrualHistory(): Promise<number> {
    const result = await this.get<{ value: { cr_accrualhistoryid: string }[] }>(
      `${ACCRUAL_HISTORY_TABLE}?$select=cr_accrualhistoryid`
    );

    let deletedCount = 0;
    for (const record of result.value) {
      if (record.cr_accrualhistoryid) {
        await this.delete(`${ACCRUAL_HISTORY_TABLE}(${record.cr_accrualhistoryid})`);
        deletedCount++;
      }
    }

    return deletedCount;
  }

  /**
   * Create a new accrual history record (admin function)
   * Used for manually adding accrual entries (initial balance, carryover, adjustments)
   */
  public async createAccrualHistory(
    accrual: Omit<IAccrualHistory, 'Id' | 'DataverseId' | 'Name' | 'Created'>
  ): Promise<{ id: string }> {
    // Generate name based on accrual type
    const monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'Maj', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dec'];
    const name = `${accrual.EmployeeName} - ${accrual.AccrualType} - ${monthNames[accrual.AccrualMonth]} ${accrual.AccrualYear}`;

    const entity: Partial<IAccrualHistoryEntity> = {
      cr_name: name,
      cr_employeeemail: accrual.EmployeeEmail,
      cr_employeename: accrual.EmployeeName,
      cr_holidayyear: accrual.HolidayYear,
      cr_accrualdate: accrual.AccrualDate.toISOString(),
      cr_accrualmonth: accrual.AccrualMonth,
      cr_accrualyear: accrual.AccrualYear,
      cr_daysaccrued: accrual.DaysAccrued,
      cr_accrualtype: ACCRUAL_TYPE_VALUES[accrual.AccrualType],
    };

    if (accrual.FeriefridageAccrued !== undefined) {
      entity.cr_feriefridageaccrued = accrual.FeriefridageAccrued;
    }
    if (accrual.BalanceAfterAccrual !== undefined) {
      entity.cr_balanceafteraccrual = accrual.BalanceAfterAccrual;
    }
    if (accrual.Notes) {
      entity.cr_notes = accrual.Notes;
    }

    return this.post<{ id: string }>(ACCRUAL_HISTORY_TABLE, entity);
  }

  /**
   * Get all employee balances for admin reporting
   * Fetches unique employees from accrual history and calculates balance for each
   * @param holidayYear Holiday year (e.g., "2024-2025")
   * @returns Array of calculated holiday balances for all employees
   */
  public async getAllEmployeeBalances(
    holidayYear: string
  ): Promise<ICalculatedHolidayBalance[]> {
    try {
      // Get all accrual history records for the holiday year to find unique employees
      const filter = `cr_holidayyear eq '${holidayYear}'`;
      const select = 'cr_employeeemail,cr_employeename';

      const result = await this.get<{ value: { cr_employeeemail: string; cr_employeename: string }[] }>(
        `${ACCRUAL_HISTORY_TABLE}?$filter=${encodeURIComponent(filter)}&$select=${select}`
      );

      // Get unique employees by email (case-insensitive)
      const employeeEmails: string[] = [];
      const seenEmails = new Set<string>();
      for (const record of result.value) {
        const emailKey = record.cr_employeeemail.toLowerCase();
        if (!seenEmails.has(emailKey)) {
          seenEmails.add(emailKey);
          employeeEmails.push(emailKey);
        }
      }

      // Calculate balance for each unique employee
      const balances: ICalculatedHolidayBalance[] = [];
      for (const email of employeeEmails) {
        const balance = await this.calculateHolidayBalanceForUser(email, holidayYear);
        balances.push(balance);
      }

      return balances;
    } catch {
      return [];
    }
  }

  /**
   * Map Dataverse entity to IAbsenceRegistration
   */
  private mapEntityToRegistration(
    entity: IDataverseEntity
  ): IAbsenceRegistration {
    return {
      Id: 0, // Using string ID internally
      Title: entity.cr153_name || '',
      EmployeeId: 0,
      EmployeeEmail: entity.cr153_employeeemail,
      EmployeeName: entity.cr153_employeename,
      Department: '', // Not used
      ApproverId: 0,
      ApproverName: entity.cr153_approvername,
      ApproverEmail: entity.cr153_approveremail,
      StartDate: new Date(entity.cr153_startdate),
      EndDate: new Date(entity.cr153_enddate),
      NumberOfDays: entity.cr153_numberofdays,
      AbsenceType: AbsenceTypeFromOptionSet[entity.cr153_absencetype] || 'Ferie',
      Notes: entity.cr153_notes,
      Status: StatusFromOptionSet[entity.cr153_status] || 'Kladde',
      ApprovalDate: entity.cr153_approvaldate
        ? new Date(entity.cr153_approvaldate)
        : undefined,
      ApproverComments: entity.cr153_approvercomments,
      Created: entity.createdon ? new Date(entity.createdon) : undefined,
      Modified: entity.modifiedon ? new Date(entity.modifiedon) : undefined,
      // Store the Dataverse GUID for update/delete operations
      DataverseId: entity.cr153_absenceregistrationid,
    };
  }
}

export default DataverseService;
