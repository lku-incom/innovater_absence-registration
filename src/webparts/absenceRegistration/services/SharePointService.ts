/**
 * Service for SharePoint list operations
 * Handles CRUD operations for absence registrations
 */

import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IAbsenceRegistration,
  AbsenceType,
  RegistrationStatus,
} from '../models/IAbsenceRegistration';

const LIST_NAME = 'AbsenceRegistrations';

interface ISharePointListItem {
  Id: number;
  Title: string;
  EmployeeId: number;
  EmployeeEmail: string;
  Department: string;
  ApproverId: number;
  StartDate: string;
  EndDate: string;
  NumberOfDays: number;
  AbsenceType: string;
  Notes: string;
  Status: string;
  ApprovalDate: string;
  ApproverComments: string;
  Created: string;
  Modified: string;
  Employee?: { Title: string; EMail: string };
  Approver?: { Title: string; EMail: string };
}

export class SharePointService {
  private static _instance: SharePointService;
  private _sp: ReturnType<typeof spfi>;
  private _context: WebPartContext;

  private constructor() {
    // Private constructor for singleton
  }

  /**
   * Initialize the SharePoint service with SPFx context
   */
  public static getInstance(context?: WebPartContext): SharePointService {
    if (!SharePointService._instance) {
      SharePointService._instance = new SharePointService();
    }

    if (context) {
      SharePointService._instance._context = context;
      SharePointService._instance._sp = spfi().using(SPFx(context));
    }

    return SharePointService._instance;
  }

  /**
   * Get user ID by email for Person fields
   */
  public async getUserIdByEmail(email: string): Promise<number> {
    try {
      const user = await this._sp.web.ensureUser(email);
      return user.data.Id;
    } catch (error) {
      console.error('Error getting user ID:', error);
      throw new Error(`Kunne ikke finde bruger: ${email}`);
    }
  }

  /**
   * Get current user's SharePoint ID
   */
  public async getCurrentUserId(): Promise<number> {
    try {
      const user = await this._sp.web.currentUser();
      return user.Id;
    } catch (error) {
      console.error('Error getting current user ID:', error);
      throw new Error('Kunne ikke hente nuværende bruger');
    }
  }

  /**
   * Create a new absence registration
   */
  public async createRegistration(
    registration: Omit<IAbsenceRegistration, 'Id' | 'Created' | 'Modified'>
  ): Promise<IAbsenceRegistration> {
    try {
      const item = await this._sp.web.lists.getByTitle(LIST_NAME).items.add({
        Title: `${registration.EmployeeName} - ${registration.AbsenceType}`,
        EmployeeId: registration.EmployeeId,
        EmployeeEmail: registration.EmployeeEmail,
        Department: registration.Department,
        ApproverId: registration.ApproverId,
        StartDate: registration.StartDate.toISOString(),
        EndDate: registration.EndDate.toISOString(),
        NumberOfDays: registration.NumberOfDays,
        AbsenceType: registration.AbsenceType,
        Notes: registration.Notes || '',
        Status: registration.Status,
      });

      return {
        ...registration,
        Id: item.data.Id,
        Created: new Date(item.data.Created),
        Modified: new Date(item.data.Modified),
      };
    } catch (error) {
      console.error('Error creating registration:', error);
      throw new Error('Kunne ikke oprette fraværsregistrering');
    }
  }

  /**
   * Update an existing absence registration
   */
  public async updateRegistration(
    id: number,
    updates: Partial<IAbsenceRegistration>
  ): Promise<void> {
    try {
      const updateData: Record<string, unknown> = {};

      if (updates.StartDate) {
        updateData.StartDate = updates.StartDate.toISOString();
      }
      if (updates.EndDate) {
        updateData.EndDate = updates.EndDate.toISOString();
      }
      if (updates.NumberOfDays !== undefined) {
        updateData.NumberOfDays = updates.NumberOfDays;
      }
      if (updates.AbsenceType) {
        updateData.AbsenceType = updates.AbsenceType;
      }
      if (updates.Notes !== undefined) {
        updateData.Notes = updates.Notes;
      }
      if (updates.Status) {
        updateData.Status = updates.Status;
      }
      if (updates.ApprovalDate) {
        updateData.ApprovalDate = updates.ApprovalDate.toISOString();
      }
      if (updates.ApproverComments !== undefined) {
        updateData.ApproverComments = updates.ApproverComments;
      }

      await this._sp.web.lists
        .getByTitle(LIST_NAME)
        .items.getById(id)
        .update(updateData);
    } catch (error) {
      console.error('Error updating registration:', error);
      throw new Error('Kunne ikke opdatere fraværsregistrering');
    }
  }

  /**
   * Delete an absence registration
   */
  public async deleteRegistration(id: number): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle(LIST_NAME).items.getById(id).delete();
    } catch (error) {
      console.error('Error deleting registration:', error);
      throw new Error('Kunne ikke slette fraværsregistrering');
    }
  }

  /**
   * Get all registrations for the current user
   */
  public async getMyRegistrations(): Promise<IAbsenceRegistration[]> {
    try {
      const userId = await this.getCurrentUserId();

      const items: ISharePointListItem[] = await this._sp.web.lists
        .getByTitle(LIST_NAME)
        .items.filter(`EmployeeId eq ${userId}`)
        .select(
          'Id',
          'Title',
          'EmployeeId',
          'EmployeeEmail',
          'Department',
          'ApproverId',
          'StartDate',
          'EndDate',
          'NumberOfDays',
          'AbsenceType',
          'Notes',
          'Status',
          'ApprovalDate',
          'ApproverComments',
          'Created',
          'Modified',
          'Employee/Title',
          'Employee/EMail',
          'Approver/Title',
          'Approver/EMail'
        )
        .expand('Employee', 'Approver')
        .orderBy('StartDate', false)();

      return items.map((item) => this.mapItemToRegistration(item));
    } catch (error) {
      console.error('Error fetching registrations:', error);
      throw new Error('Kunne ikke hente fraværsregistreringer');
    }
  }

  /**
   * Get registrations by status
   */
  public async getRegistrationsByStatus(
    status: RegistrationStatus
  ): Promise<IAbsenceRegistration[]> {
    try {
      const userId = await this.getCurrentUserId();

      const items: ISharePointListItem[] = await this._sp.web.lists
        .getByTitle(LIST_NAME)
        .items.filter(`EmployeeId eq ${userId} and Status eq '${status}'`)
        .select(
          'Id',
          'Title',
          'EmployeeId',
          'EmployeeEmail',
          'Department',
          'ApproverId',
          'StartDate',
          'EndDate',
          'NumberOfDays',
          'AbsenceType',
          'Notes',
          'Status',
          'ApprovalDate',
          'ApproverComments',
          'Created',
          'Modified',
          'Employee/Title',
          'Employee/EMail',
          'Approver/Title',
          'Approver/EMail'
        )
        .expand('Employee', 'Approver')
        .orderBy('StartDate', false)();

      return items.map((item) => this.mapItemToRegistration(item));
    } catch (error) {
      console.error('Error fetching registrations by status:', error);
      throw new Error('Kunne ikke hente fraværsregistreringer');
    }
  }

  /**
   * Get a single registration by ID
   */
  public async getRegistrationById(
    id: number
  ): Promise<IAbsenceRegistration | undefined> {
    try {
      const item: ISharePointListItem = await this._sp.web.lists
        .getByTitle(LIST_NAME)
        .items.getById(id)
        .select(
          'Id',
          'Title',
          'EmployeeId',
          'EmployeeEmail',
          'Department',
          'ApproverId',
          'StartDate',
          'EndDate',
          'NumberOfDays',
          'AbsenceType',
          'Notes',
          'Status',
          'ApprovalDate',
          'ApproverComments',
          'Created',
          'Modified',
          'Employee/Title',
          'Employee/EMail',
          'Approver/Title',
          'Approver/EMail'
        )
        .expand('Employee', 'Approver')();

      return this.mapItemToRegistration(item);
    } catch (error) {
      console.error('Error fetching registration by ID:', error);
      return undefined;
    }
  }

  /**
   * Submit registration for approval (change status to Pending)
   */
  public async submitForApproval(id: number): Promise<void> {
    await this.updateRegistration(id, {
      Status: 'Afventer godkendelse',
    });
  }

  /**
   * Map SharePoint list item to IAbsenceRegistration
   */
  private mapItemToRegistration(
    item: ISharePointListItem
  ): IAbsenceRegistration {
    return {
      Id: item.Id,
      Title: item.Title,
      EmployeeId: item.EmployeeId,
      EmployeeEmail: item.EmployeeEmail,
      EmployeeName: item.Employee?.Title || '',
      Department: item.Department,
      ApproverId: item.ApproverId,
      ApproverName: item.Approver?.Title || '',
      ApproverEmail: item.Approver?.EMail || '',
      StartDate: new Date(item.StartDate),
      EndDate: new Date(item.EndDate),
      NumberOfDays: item.NumberOfDays,
      AbsenceType: item.AbsenceType as AbsenceType,
      Notes: item.Notes,
      Status: item.Status as RegistrationStatus,
      ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
      ApproverComments: item.ApproverComments,
      Created: new Date(item.Created),
      Modified: new Date(item.Modified),
    };
  }

  /**
   * Ensure the AbsenceRegistrations list exists
   * This can be called during web part initialization
   */
  public async ensureListExists(): Promise<boolean> {
    try {
      await this._sp.web.lists.getByTitle(LIST_NAME)();
      return true;
    } catch {
      console.warn(
        `List "${LIST_NAME}" does not exist. Please create it manually or use a provisioning script.`
      );
      return false;
    }
  }
}

export default SharePointService;
