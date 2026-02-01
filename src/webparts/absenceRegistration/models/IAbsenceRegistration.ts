/**
 * Data models for the Absence Registration solution
 */

export interface IAbsenceRegistration {
  Id?: number;
  DataverseId?: string; // GUID from Dataverse
  Title?: string;
  EmployeeId: number;
  EmployeeEmail: string;
  EmployeeName: string;
  Department: string;
  ApproverId: number;
  ApproverName: string;
  ApproverEmail: string;
  StartDate: Date;
  EndDate: Date;
  NumberOfDays: number;
  AbsenceType: AbsenceType;
  Notes?: string;
  Status: RegistrationStatus;
  ApprovalDate?: Date;
  ApproverComments?: string;
  Created?: Date;
  Modified?: Date;
}

export type AbsenceType =
  | 'Ferie'
  | 'Sygdom'
  | 'Barselsorlov'
  | 'Feriefridage'
  | 'Flex/afspadsering'
  | 'Andet fravær';

export type RegistrationStatus =
  | 'Kladde'
  | 'Afventer godkendelse'
  | 'Godkendt'
  | 'Afvist';

export interface IAbsenceTypeOption {
  key: AbsenceType;
  text: string;
}

export const AbsenceTypeOptions: IAbsenceTypeOption[] = [
  { key: 'Ferie', text: 'Ferie' },
  { key: 'Sygdom', text: 'Sygdom' },
  { key: 'Barselsorlov', text: 'Barselsorlov' },
  { key: 'Feriefridage', text: 'Feriefridage' },
  { key: 'Flex/afspadsering', text: 'Flex/afspadsering' },
  { key: 'Andet fravær', text: 'Andet fravær' },
];

export interface IUserInfo {
  id: number;
  email: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  manager?: IManagerInfo;
}

export interface IManagerInfo {
  id: string;
  email: string;
  displayName: string;
}
