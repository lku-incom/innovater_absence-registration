/**
 * Interface definitions for component props
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IAbsenceRegistration,
  IUserInfo,
  AbsenceType,
} from '../models/IAbsenceRegistration';
import { IAccrualHistory } from '../models/IHolidayBalance';

export interface IAbsenceRegistrationProps {
  context: WebPartContext;
  title: string;
  dataverseUrl?: string;
}

export interface IAbsenceFormProps {
  context: WebPartContext;
  currentUser: IUserInfo | undefined;
  onSave: (registration: IAbsenceRegistration) => Promise<void>;
  onSubmit: (registration: IAbsenceRegistration) => Promise<void>;
  onCancel: () => void;
  editingRegistration?: IAbsenceRegistration;
  isLoading: boolean;
  readOnly?: boolean;
}

export interface IMyRegistrationsProps {
  context: WebPartContext;
  registrations: IAbsenceRegistration[];
  onEdit: (registration: IAbsenceRegistration) => void;
  onView: (registration: IAbsenceRegistration) => void;
  onDelete: (id: string) => Promise<void>;
  onSubmitForApproval: (id: string) => Promise<void>;
  isLoading: boolean;
  onRefresh: () => void;
  currentUserEmail?: string;
}

export interface IPendingApprovalsProps {
  context: WebPartContext;
  pendingApprovals: IAbsenceRegistration[];
  onView: (registration: IAbsenceRegistration) => void;
  isLoading: boolean;
  onRefresh: () => void;
}

export interface IAbsenceFormState {
  startDate: Date | undefined;
  endDate: Date | undefined;
  absenceType: AbsenceType | undefined;
  notes: string;
  numberOfDays: number;
  isSaving: boolean;
  errorMessage: string;
}

export interface IAbsenceRegistrationState {
  activeTab: 'new' | 'list' | 'approvals' | 'admin';
  currentUser: IUserInfo | undefined;
  registrations: IAbsenceRegistration[];
  pendingApprovals: IAbsenceRegistration[];
  editingRegistration: IAbsenceRegistration | undefined;
  isLoading: boolean;
  errorMessage: string;
}

export interface IAdminPanelProps {
  context: WebPartContext;
  allRegistrations: IAbsenceRegistration[];
  onDeleteAll: () => Promise<void>;
  onDeleteSingle: (id: string) => Promise<void>;
  onDeleteAllAccrualHistory: () => Promise<number>;
  onCreateAccrualHistory: (accrual: Omit<IAccrualHistory, 'Id' | 'DataverseId' | 'Name' | 'Created'>) => Promise<void>;
  accrualHistoryCount: number;
  isLoading: boolean;
  onRefresh: () => void;
  // Impersonation props
  impersonatedUser?: IUserInfo;
  onImpersonate: (user: IUserInfo | undefined) => void;
}
