import * as React from 'react';
import { useState, useCallback, useRef, useEffect } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { DefaultButton, PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Modal } from '@fluentui/react/lib/Modal';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Separator } from '@fluentui/react/lib/Separator';
import { TextField } from '@fluentui/react/lib/TextField';
import { SpinButton } from '@fluentui/react/lib/SpinButton';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { List } from '@fluentui/react/lib/List';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import styles from './AbsenceRegistration.module.scss';
import { IAdminPanelProps } from './IAbsenceRegistrationProps';
import { IAbsenceRegistration, RegistrationStatus, IUserInfo } from '../models/IAbsenceRegistration';
import { AccrualType, getHolidayYear, IAccrualHistory } from '../models/IHolidayBalance';
import { GraphService, IGroupMember } from '../services/GraphService';
import AdminReport from './AdminReport';

// Azure AD security group ID for holiday-eligible employees
const HOLIDAY_ELIGIBLE_GROUP_ID = '775b14c1-65ed-472f-915d-6e88c446318a';

// Number of group members to show initially
const INITIAL_MEMBERS_TO_SHOW = 5;

// Accrual type options for the dropdown (must match Dataverse option set values 100000000-100000003)
const accrualTypeOptions: IDropdownOption[] = [
  { key: 'Startsaldo', text: 'Startsaldo (Initial Balance)' },
  { key: 'Årsstart overførsel', text: 'Årsstart overførsel (Year Start Carryover)' },
  { key: 'Månedlig optjening', text: 'Månedlig optjening (Monthly Accrual)' },
  { key: 'Manuel justering', text: 'Manuel justering (Manual Adjustment)' },
];

// Generate holiday year options (current year and previous 2 years)
const generateHolidayYearOptions = (): IDropdownOption[] => {
  const currentYear = getHolidayYear();
  const [startYear] = currentYear.split('-').map(Number);
  const options: IDropdownOption[] = [];

  for (let i = -1; i <= 1; i++) {
    const year = startYear + i;
    options.push({ key: `${year}-${year + 1}`, text: `${year}-${year + 1}` });
  }

  return options;
};

interface IAccrualHistoryFormState {
  employeeEmail: string;
  employeeName: string;
  holidayYear: string;
  accrualType: AccrualType;
  accrualDate: Date;
  daysAccrued: number;
  feriefridageAccrued: number;
  notes: string;
}

const initialFormState: IAccrualHistoryFormState = {
  employeeEmail: '',
  employeeName: '',
  holidayYear: getHolidayYear(),
  accrualType: 'Startsaldo',
  accrualDate: new Date(),
  daysAccrued: 0,
  feriefridageAccrued: 0,
  notes: '',
};

const AdminPanel: React.FC<IAdminPanelProps> = (props) => {
  const {
    context,
    allRegistrations,
    onDeleteAll,
    onDeleteSingle,
    onDeleteAllAccrualHistory,
    onCreateAccrualHistory,
    accrualHistoryCount,
    isLoading,
    onRefresh,
    impersonatedUser,
    onImpersonate,
  } = props;

  const [isDeleteAllDialogOpen, setIsDeleteAllDialogOpen] = useState<boolean>(false);
  const [isDeleteSingleDialogOpen, setIsDeleteSingleDialogOpen] = useState<boolean>(false);
  const [isDeleteHistoryDialogOpen, setIsDeleteHistoryDialogOpen] = useState<boolean>(false);
  const [showReport, setShowReport] = useState<boolean>(false);
  const [deletingId, setDeletingId] = useState<string | undefined>(undefined);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);
  const [deleteResult, setDeleteResult] = useState<{ message: string; type: MessageBarType } | undefined>(undefined);

  // Accrual History Form State
  const [formState, setFormState] = useState<IAccrualHistoryFormState>(initialFormState);
  const [isAddingAccrual, setIsAddingAccrual] = useState<boolean>(false);
  const [showAddAccrualForm, setShowAddAccrualForm] = useState<boolean>(false);
  const [formError, setFormError] = useState<string>('');

  // User Search State
  const [userSearchText, setUserSearchText] = useState<string>('');
  const [userSearchResults, setUserSearchResults] = useState<IUserInfo[]>([]);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const [showUserDropdown, setShowUserDropdown] = useState<boolean>(false);
  const [selectedUser, setSelectedUser] = useState<IUserInfo | undefined>(undefined);
  const searchInputRef = useRef<HTMLDivElement>(null);
  const searchTimeoutRef = useRef<number | undefined>(undefined);

  // Group Management State
  const [groupMembers, setGroupMembers] = useState<IGroupMember[]>([]);
  const [isLoadingGroupMembers, setIsLoadingGroupMembers] = useState<boolean>(false);
  const [groupUserSearchText, setGroupUserSearchText] = useState<string>('');
  const [groupUserSearchResults, setGroupUserSearchResults] = useState<IGroupMember[]>([]);
  const [isSearchingGroupUser, setIsSearchingGroupUser] = useState<boolean>(false);
  const [showGroupUserDropdown, setShowGroupUserDropdown] = useState<boolean>(false);
  const [isAddingToGroup, setIsAddingToGroup] = useState<boolean>(false);
  const [isRemovingFromGroup, setIsRemovingFromGroup] = useState<string | undefined>(undefined);
  const [groupError, setGroupError] = useState<string>('');
  const [groupSuccess, setGroupSuccess] = useState<string>('');
  const [showAllMembers, setShowAllMembers] = useState<boolean>(false);
  const groupSearchInputRef = useRef<HTMLDivElement>(null);
  const groupSearchTimeoutRef = useRef<number | undefined>(undefined);

  // Impersonation Search State
  const [impersonationSearchText, setImpersonationSearchText] = useState<string>('');
  const [impersonationSearchResults, setImpersonationSearchResults] = useState<IGroupMember[]>([]);
  const [isSearchingImpersonation, setIsSearchingImpersonation] = useState<boolean>(false);
  const [showImpersonationDropdown, setShowImpersonationDropdown] = useState<boolean>(false);
  const impersonationSearchInputRef = useRef<HTMLDivElement>(null);
  const impersonationSearchTimeoutRef = useRef<number | undefined>(undefined);

  const getStatusBadgeClass = (status: RegistrationStatus): string => {
    switch (status) {
      case 'Kladde':
        return styles.draft;
      case 'Afventer godkendelse':
        return styles.pending;
      case 'Godkendt':
        return styles.approved;
      case 'Afvist':
        return styles.rejected;
      default:
        return '';
    }
  };

  const formatDate = (date: Date | undefined): string => {
    if (!date) return '-';
    return date.toLocaleDateString('da-DK', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  };

  const handleDeleteAllClick = useCallback((): void => {
    setIsDeleteAllDialogOpen(true);
  }, []);

  const handleDeleteAllConfirm = useCallback(async (): Promise<void> => {
    setIsDeleting(true);
    setDeleteResult(undefined);
    try {
      await onDeleteAll();
      setIsDeleteAllDialogOpen(false);
      setDeleteResult({ message: 'Alle registreringer er slettet.', type: MessageBarType.success });
    } catch (error) {
      setDeleteResult({ message: 'Fejl ved sletning af registreringer.', type: MessageBarType.error });
    } finally {
      setIsDeleting(false);
    }
  }, [onDeleteAll]);

  const handleDeleteSingleClick = useCallback((id: string): void => {
    setDeletingId(id);
    setIsDeleteSingleDialogOpen(true);
  }, []);

  const handleDeleteSingleConfirm = useCallback(async (): Promise<void> => {
    if (!deletingId) return;
    setIsDeleting(true);
    setDeleteResult(undefined);
    try {
      await onDeleteSingle(deletingId);
      setIsDeleteSingleDialogOpen(false);
      setDeletingId(undefined);
    } finally {
      setIsDeleting(false);
    }
  }, [deletingId, onDeleteSingle]);

  const handleDeleteHistoryClick = useCallback((): void => {
    setIsDeleteHistoryDialogOpen(true);
  }, []);

  const handleDeleteHistoryConfirm = useCallback(async (): Promise<void> => {
    setIsDeleting(true);
    setDeleteResult(undefined);
    try {
      const count = await onDeleteAllAccrualHistory();
      setIsDeleteHistoryDialogOpen(false);
      setDeleteResult({ message: `${count} optjeningshistorik poster er slettet.`, type: MessageBarType.success });
      onRefresh();
    } catch (error) {
      setDeleteResult({ message: 'Fejl ved sletning af optjeningshistorik.', type: MessageBarType.error });
    } finally {
      setIsDeleting(false);
    }
  }, [onDeleteAllAccrualHistory, onRefresh]);

  const handleDialogDismiss = useCallback((): void => {
    setIsDeleteAllDialogOpen(false);
    setIsDeleteSingleDialogOpen(false);
    setIsDeleteHistoryDialogOpen(false);
    setDeletingId(undefined);
  }, []);

  // Accrual History Form Handlers
  const handleFormChange = useCallback((field: keyof IAccrualHistoryFormState, value: string | number | Date): void => {
    setFormState((prev) => ({ ...prev, [field]: value }));
    setFormError('');
  }, []);

  const handleSpinButtonChange = useCallback((field: keyof IAccrualHistoryFormState) => {
    return (_event: React.SyntheticEvent<HTMLElement>, newValue?: string): void => {
      const numValue = parseFloat(newValue || '0') || 0;
      handleFormChange(field, numValue);
    };
  }, [handleFormChange]);

  const handleAddAccrualSubmit = useCallback(async (): Promise<void> => {
    // Validate form
    if (!selectedUser) {
      setFormError('Vælg venligst en medarbejder');
      return;
    }
    if (!formState.holidayYear) {
      setFormError('Ferieår er påkrævet');
      return;
    }
    if (!formState.accrualType) {
      setFormError('Optjeningstype er påkrævet');
      return;
    }
    if (formState.daysAccrued === 0 && formState.feriefridageAccrued === 0) {
      setFormError('Angiv mindst ét antal dage (feriedage eller feriefridage)');
      return;
    }

    setIsAddingAccrual(true);
    setFormError('');
    setDeleteResult(undefined);

    try {
      const accrualDate = formState.accrualDate;
      const accrual: Omit<IAccrualHistory, 'Id' | 'DataverseId' | 'Name' | 'Created'> = {
        EmployeeEmail: selectedUser.email,
        EmployeeName: selectedUser.displayName,
        HolidayYear: formState.holidayYear,
        AccrualDate: accrualDate,
        AccrualMonth: accrualDate.getMonth() + 1, // 1-indexed
        AccrualYear: accrualDate.getFullYear(),
        DaysAccrued: formState.daysAccrued,
        FeriefridageAccrued: formState.feriefridageAccrued,
        AccrualType: formState.accrualType,
        Notes: formState.notes || undefined,
      };

      await onCreateAccrualHistory(accrual);
      setDeleteResult({ message: `Optjening oprettet for ${selectedUser.displayName}`, type: MessageBarType.success });
      setFormState(initialFormState);
      setShowAddAccrualForm(false);
      setUserSearchText('');
      setUserSearchResults([]);
      setSelectedUser(undefined);
      onRefresh();
    } catch {
      setFormError('Kunne ikke oprette optjening. Prøv venligst igen.');
    } finally {
      setIsAddingAccrual(false);
    }
  }, [formState, selectedUser, onCreateAccrualHistory, onRefresh]);

  const handleCancelAddAccrual = useCallback((): void => {
    setFormState(initialFormState);
    setShowAddAccrualForm(false);
    setFormError('');
    setUserSearchText('');
    setUserSearchResults([]);
    setSelectedUser(undefined);
    setShowUserDropdown(false);
  }, []);

  // User Search Handlers
  const handleUserSearch = useCallback(async (searchText: string): Promise<void> => {
    if (searchText.length < 2) {
      setUserSearchResults([]);
      setShowUserDropdown(false);
      return;
    }

    setIsSearching(true);
    try {
      const graphService = GraphService.getInstance();
      const results = await graphService.searchUsers(searchText);
      setUserSearchResults(results);
      setShowUserDropdown(results.length > 0);
    } catch {
      setUserSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  }, []);

  const handleUserSearchChange = useCallback((
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const value = newValue || '';
    setUserSearchText(value);
    setSelectedUser(undefined);
    handleFormChange('employeeEmail', '');
    handleFormChange('employeeName', '');

    // Debounce search
    if (searchTimeoutRef.current) {
      window.clearTimeout(searchTimeoutRef.current);
    }
    searchTimeoutRef.current = window.setTimeout(() => {
      handleUserSearch(value);
    }, 300);
  }, [handleUserSearch, handleFormChange]);

  const handleUserSelect = useCallback((user: IUserInfo): void => {
    setSelectedUser(user);
    setUserSearchText(user.displayName);
    handleFormChange('employeeEmail', user.email);
    handleFormChange('employeeName', user.displayName);
    setShowUserDropdown(false);
    setUserSearchResults([]);
  }, [handleFormChange]);

  const handleSearchBlur = useCallback((): void => {
    // Delay hiding to allow click on dropdown item
    setTimeout(() => {
      setShowUserDropdown(false);
    }, 200);
  }, []);

  // Cleanup timeout on unmount
  useEffect(() => {
    return () => {
      if (searchTimeoutRef.current) {
        window.clearTimeout(searchTimeoutRef.current);
      }
      if (groupSearchTimeoutRef.current) {
        window.clearTimeout(groupSearchTimeoutRef.current);
      }
      if (impersonationSearchTimeoutRef.current) {
        window.clearTimeout(impersonationSearchTimeoutRef.current);
      }
    };
  }, []);

  // Group Management Handlers
  const loadGroupMembers = useCallback(async (): Promise<void> => {
    setIsLoadingGroupMembers(true);
    setGroupError('');
    try {
      const graphService = GraphService.getInstance();
      const members = await graphService.getGroupMembers(HOLIDAY_ELIGIBLE_GROUP_ID);
      setGroupMembers(members);
    } catch {
      setGroupError('Kunne ikke hente gruppemedlemmer. Kontroller tilladelser.');
      setGroupMembers([]);
    } finally {
      setIsLoadingGroupMembers(false);
    }
  }, []);

  const handleGroupUserSearch = useCallback(async (searchText: string): Promise<void> => {
    if (searchText.length < 2) {
      setGroupUserSearchResults([]);
      setShowGroupUserDropdown(false);
      return;
    }

    setIsSearchingGroupUser(true);
    try {
      const graphService = GraphService.getInstance();
      const results = await graphService.searchUsersWithId(searchText);
      // Filter out users already in group
      const filteredResults = results.filter(
        (user) => !groupMembers.some((member) => member.id === user.id)
      );
      setGroupUserSearchResults(filteredResults);
      setShowGroupUserDropdown(filteredResults.length > 0);
    } catch {
      setGroupUserSearchResults([]);
    } finally {
      setIsSearchingGroupUser(false);
    }
  }, [groupMembers]);

  const handleGroupUserSearchChange = useCallback((
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const value = newValue || '';
    setGroupUserSearchText(value);

    // Debounce search
    if (groupSearchTimeoutRef.current) {
      window.clearTimeout(groupSearchTimeoutRef.current);
    }
    groupSearchTimeoutRef.current = window.setTimeout(() => {
      handleGroupUserSearch(value);
    }, 300);
  }, [handleGroupUserSearch]);

  const handleAddUserToGroup = useCallback(async (user: IGroupMember): Promise<void> => {
    setIsAddingToGroup(true);
    setGroupError('');
    setGroupSuccess('');
    setShowGroupUserDropdown(false);
    setGroupUserSearchText('');
    setGroupUserSearchResults([]);

    try {
      const graphService = GraphService.getInstance();
      await graphService.addUserToGroup(HOLIDAY_ELIGIBLE_GROUP_ID, user.id);
      setGroupMembers((prev) => [...prev, user]);
      setGroupSuccess(`${user.displayName} er tilføjet til gruppen`);
    } catch {
      setGroupError(`Kunne ikke tilføje ${user.displayName} til gruppen`);
    } finally {
      setIsAddingToGroup(false);
    }
  }, []);

  const handleRemoveUserFromGroup = useCallback(async (user: IGroupMember): Promise<void> => {
    setIsRemovingFromGroup(user.id);
    setGroupError('');
    setGroupSuccess('');

    try {
      const graphService = GraphService.getInstance();
      await graphService.removeUserFromGroup(HOLIDAY_ELIGIBLE_GROUP_ID, user.id);
      setGroupMembers((prev) => prev.filter((m) => m.id !== user.id));
      setGroupSuccess(`${user.displayName} er fjernet fra gruppen`);
    } catch {
      setGroupError(`Kunne ikke fjerne ${user.displayName} fra gruppen`);
    } finally {
      setIsRemovingFromGroup(undefined);
    }
  }, []);

  const handleGroupSearchBlur = useCallback((): void => {
    // Delay hiding to allow click on dropdown item
    setTimeout(() => {
      setShowGroupUserDropdown(false);
    }, 200);
  }, []);

  // Impersonation Search Handlers
  const handleImpersonationSearch = useCallback(async (searchText: string): Promise<void> => {
    if (searchText.length < 2) {
      setImpersonationSearchResults([]);
      setShowImpersonationDropdown(false);
      return;
    }

    setIsSearchingImpersonation(true);
    try {
      const graphService = GraphService.getInstance();
      const results = await graphService.searchUsersWithId(searchText);
      setImpersonationSearchResults(results);
      setShowImpersonationDropdown(results.length > 0);
    } catch {
      setImpersonationSearchResults([]);
    } finally {
      setIsSearchingImpersonation(false);
    }
  }, []);

  const handleImpersonationSearchChange = useCallback((
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const value = newValue || '';
    setImpersonationSearchText(value);

    // Debounce search
    if (impersonationSearchTimeoutRef.current) {
      window.clearTimeout(impersonationSearchTimeoutRef.current);
    }
    impersonationSearchTimeoutRef.current = window.setTimeout(() => {
      handleImpersonationSearch(value);
    }, 300);
  }, [handleImpersonationSearch]);

  const handleSelectUserToImpersonate = useCallback((user: IGroupMember): void => {
    setShowImpersonationDropdown(false);
    setImpersonationSearchText('');
    setImpersonationSearchResults([]);
    onImpersonate({
      id: 0,
      email: user.email,
      displayName: user.displayName,
      jobTitle: user.jobTitle,
    });
  }, [onImpersonate]);

  const handleImpersonationSearchBlur = useCallback((): void => {
    // Delay hiding to allow click on dropdown item
    setTimeout(() => {
      setShowImpersonationDropdown(false);
    }, 200);
  }, []);

  // Load group members on mount
  useEffect(() => {
    loadGroupMembers();
  }, [loadGroupMembers]);

  const columns: IColumn[] = [
    {
      key: 'employee',
      name: 'Medarbejder',
      fieldName: 'EmployeeName',
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: 'type',
      name: 'Type',
      fieldName: 'AbsenceType',
      minWidth: 100,
      maxWidth: 140,
      isResizable: true,
    },
    {
      key: 'period',
      name: 'Periode',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => (
        <span>
          {formatDate(item.StartDate)} - {formatDate(item.EndDate)}
        </span>
      ),
    },
    {
      key: 'days',
      name: 'Dage',
      fieldName: 'NumberOfDays',
      minWidth: 50,
      maxWidth: 60,
      isResizable: true,
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => (
        <span className={`${styles.statusBadge} ${getStatusBadgeClass(item.Status)}`}>
          {item.Status}
        </span>
      ),
    },
    {
      key: 'created',
      name: 'Oprettet',
      minWidth: 90,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => (
        <span>{formatDate(item.Created)}</span>
      ),
    },
    {
      key: 'actions',
      name: '',
      minWidth: 50,
      maxWidth: 50,
      onRender: (item: IAbsenceRegistration) => (
        <div className={styles.actionButtons}>
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Slet"
            ariaLabel="Slet registrering"
            onClick={() => item.DataverseId && handleDeleteSingleClick(item.DataverseId)}
            disabled={isLoading || isDeleting}
          />
        </div>
      ),
    },
  ];

  if (isLoading && allRegistrations.length === 0) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Indlæser..." />
      </div>
    );
  }

  return (
    <div className={styles.formSection}>
      <h2 className={styles.sectionTitle}>Administration</h2>

      <MessageBar messageBarType={MessageBarType.warning} style={{ marginBottom: 16 }}>
        Administratorfunktioner kan slette data permanent. Handlingerne kan ikke fortrydes.
      </MessageBar>

      {/* User Impersonation Section */}
      <div style={{
        backgroundColor: impersonatedUser ? '#fff3cd' : '#f0f9ff',
        borderRadius: 8,
        padding: 20,
        border: impersonatedUser ? '1px solid #ffc107' : '1px solid #b3d9ff',
        marginBottom: 24,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <div>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="ContactInfo" />
              Se appen som en anden bruger
            </h3>
            <p style={{ fontSize: 12, color: '#666', margin: '4px 0 0 0' }}>
              Vælg en medarbejder for at se deres registreringer og feriesaldo (skrivebeskyttet)
            </p>
          </div>
        </div>

        {impersonatedUser ? (
          <div style={{
            backgroundColor: '#fff',
            borderRadius: 6,
            padding: 16,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <Icon iconName="Contact" style={{ fontSize: 24, color: '#856404' }} />
              <div>
                <div style={{ fontWeight: 600 }}>{impersonatedUser.displayName}</div>
                <div style={{ fontSize: 12, color: '#666' }}>{impersonatedUser.email}</div>
              </div>
            </div>
            <DefaultButton
              text="Afslut visning"
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => onImpersonate(undefined)}
            />
          </div>
        ) : (
          <div ref={impersonationSearchInputRef} style={{ position: 'relative', maxWidth: 400 }}>
            <TextField
              placeholder="Søg på navn eller e-mail..."
              value={impersonationSearchText}
              onChange={handleImpersonationSearchChange}
              onBlur={handleImpersonationSearchBlur}
              onFocus={() => impersonationSearchResults.length > 0 && setShowImpersonationDropdown(true)}
              iconProps={isSearchingImpersonation ? undefined : { iconName: 'Search' }}
              styles={{ root: { marginBottom: 0 } }}
            />
            {isSearchingImpersonation && (
              <div style={{ position: 'absolute', right: 8, top: 8 }}>
                <Spinner size={SpinnerSize.small} />
              </div>
            )}
            {showImpersonationDropdown && impersonationSearchResults.length > 0 && impersonationSearchInputRef.current && (
              <Callout
                target={impersonationSearchInputRef.current}
                isBeakVisible={false}
                directionalHint={DirectionalHint.bottomLeftEdge}
                onDismiss={() => setShowImpersonationDropdown(false)}
                setInitialFocus={false}
                styles={{
                  root: { width: impersonationSearchInputRef.current.offsetWidth },
                  calloutMain: { maxHeight: 300, overflowY: 'auto' },
                }}
              >
                <List
                  items={impersonationSearchResults}
                  onRenderCell={(user) => user && (
                    <div
                      style={{
                        padding: '10px 14px',
                        cursor: 'pointer',
                        borderBottom: '1px solid #edebe9',
                        transition: 'background-color 0.15s',
                      }}
                      onClick={() => handleSelectUserToImpersonate(user)}
                      onMouseDown={(e) => e.preventDefault()}
                      onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#f3f2f1')}
                      onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                    >
                      <Persona
                        text={user.displayName}
                        secondaryText={user.email}
                        tertiaryText={user.jobTitle}
                        size={PersonaSize.size32}
                        showSecondaryText
                      />
                    </div>
                  )}
                />
              </Callout>
            )}
          </div>
        )}
      </div>

      {deleteResult && (
        <MessageBar
          messageBarType={deleteResult.type}
          onDismiss={() => setDeleteResult(undefined)}
          style={{ marginBottom: 16 }}
        >
          {deleteResult.message}
        </MessageBar>
      )}

      {/* Group Management Section - At the top */}
      <div style={{
        backgroundColor: '#fafafa',
        borderRadius: 8,
        padding: 20,
        border: '1px solid #e0e0e0',
        marginBottom: 24,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <div>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="Group" />
              Ferieberettigede medarbejdere
            </h3>
            <p style={{ fontSize: 12, color: '#666', margin: '4px 0 0 0' }}>
              Medarbejdere der modtager månedlig ferieoptjening via Power Automate
            </p>
          </div>
          <DefaultButton
            iconProps={{ iconName: 'Refresh' }}
            text="Opdater"
            onClick={loadGroupMembers}
            disabled={isLoadingGroupMembers}
          />
        </div>

        {groupError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setGroupError('')}
            style={{ marginBottom: 12 }}
          >
            {groupError}
          </MessageBar>
        )}

        {groupSuccess && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => setGroupSuccess('')}
            style={{ marginBottom: 12 }}
          >
            {groupSuccess}
          </MessageBar>
        )}

        {/* Add User to Group */}
        <div style={{ marginBottom: 16 }}>
          <div ref={groupSearchInputRef} style={{ position: 'relative', maxWidth: 350 }}>
            <TextField
              placeholder="Tilføj medarbejder (søg på navn eller e-mail)..."
              value={groupUserSearchText}
              onChange={handleGroupUserSearchChange}
              onBlur={handleGroupSearchBlur}
              onFocus={() => groupUserSearchResults.length > 0 && setShowGroupUserDropdown(true)}
              disabled={isAddingToGroup || isLoadingGroupMembers}
              iconProps={isSearchingGroupUser ? undefined : { iconName: 'AddFriend' }}
              styles={{ root: { marginBottom: 0 } }}
            />
            {isSearchingGroupUser && (
              <div style={{ position: 'absolute', right: 8, top: 8 }}>
                <Spinner size={SpinnerSize.small} />
              </div>
            )}
            {showGroupUserDropdown && groupUserSearchResults.length > 0 && groupSearchInputRef.current && (
              <Callout
                target={groupSearchInputRef.current}
                isBeakVisible={false}
                directionalHint={DirectionalHint.bottomLeftEdge}
                onDismiss={() => setShowGroupUserDropdown(false)}
                setInitialFocus={false}
                styles={{
                  root: { width: groupSearchInputRef.current.offsetWidth },
                  calloutMain: { maxHeight: 300, overflowY: 'auto' },
                }}
              >
                <List
                  items={groupUserSearchResults}
                  onRenderCell={(user) => user && (
                    <div
                      style={{
                        padding: '10px 14px',
                        cursor: 'pointer',
                        borderBottom: '1px solid #edebe9',
                        transition: 'background-color 0.15s',
                      }}
                      onClick={() => handleAddUserToGroup(user)}
                      onMouseDown={(e) => e.preventDefault()}
                      onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#f3f2f1')}
                      onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                    >
                      <Persona
                        text={user.displayName}
                        secondaryText={user.email}
                        tertiaryText={user.jobTitle}
                        size={PersonaSize.size32}
                        showSecondaryText
                      />
                    </div>
                  )}
                />
              </Callout>
            )}
          </div>
        </div>

        {/* Group Members List */}
        {isLoadingGroupMembers ? (
          <div style={{ padding: 16, textAlign: 'center' }}>
            <Spinner size={SpinnerSize.small} label="Henter..." />
          </div>
        ) : groupMembers.length === 0 ? (
          <div style={{
            padding: 16,
            textAlign: 'center',
            color: '#666',
            backgroundColor: '#fff',
            borderRadius: 6,
            border: '1px dashed #c8c8c8',
          }}>
            <Icon iconName="Group" style={{ fontSize: 20, marginBottom: 4 }} />
            <div style={{ fontSize: 13 }}>Ingen medlemmer i gruppen</div>
          </div>
        ) : (
          <>
            <div style={{ fontSize: 12, color: '#666', marginBottom: 8 }}>
              <strong>{groupMembers.length}</strong> medlemmer
            </div>
            <div style={{
              backgroundColor: '#fff',
              borderRadius: 6,
              border: '1px solid #e0e0e0',
            }}>
              {(showAllMembers ? groupMembers : groupMembers.slice(0, INITIAL_MEMBERS_TO_SHOW)).map((member) => (
                <div
                  key={member.id}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '8px 12px',
                    borderBottom: '1px solid #edebe9',
                  }}
                >
                  <Persona
                    text={member.displayName}
                    secondaryText={member.email}
                    size={PersonaSize.size24}
                    showSecondaryText
                  />
                  <IconButton
                    iconProps={{ iconName: 'Cancel' }}
                    title="Fjern fra gruppe"
                    ariaLabel={`Fjern ${member.displayName} fra gruppe`}
                    onClick={() => handleRemoveUserFromGroup(member)}
                    disabled={isRemovingFromGroup === member.id}
                    styles={{
                      root: { color: '#c0392b', width: 28, height: 28 },
                      rootHovered: { color: '#a93226', backgroundColor: '#fce4e4' },
                      icon: { fontSize: 12 },
                    }}
                  />
                </div>
              ))}
            </div>
            {groupMembers.length > INITIAL_MEMBERS_TO_SHOW && (
              <DefaultButton
                text={showAllMembers ? `Vis færre` : `Vis alle ${groupMembers.length} medlemmer`}
                iconProps={{ iconName: showAllMembers ? 'ChevronUp' : 'ChevronDown' }}
                onClick={() => setShowAllMembers(!showAllMembers)}
                styles={{
                  root: { marginTop: 8, width: '100%' },
                }}
              />
            )}
          </>
        )}
      </div>

      {/* Report Section */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          backgroundColor: '#f3f2f1',
          borderRadius: 8,
          padding: '16px 20px',
          marginBottom: 24,
          border: '1px solid #e0e0e0',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <Icon iconName="ReportDocument" style={{ fontSize: 24, color: '#004e6b' }} />
          <div>
            <div style={{ fontSize: 14, fontWeight: 600, color: '#333' }}>Ferieoversigt</div>
            <div style={{ fontSize: 12, color: '#666' }}>
              Aggregeret rapport over alle medarbejderes feriesaldo
            </div>
          </div>
        </div>
        <PrimaryButton
          text="Åbn rapport"
          iconProps={{ iconName: 'BarChart4' }}
          onClick={() => setShowReport(true)}
          disabled={isLoading}
        />
      </div>

      {/* Absence Registrations Section */}
      <h3 style={{ marginBottom: 8 }}>Fraværsregistreringer (Absence Registrations)</h3>
      <div className={styles.registrationsHeader}>
        <span>
          <strong>{allRegistrations.length}</strong> registreringer
        </span>
        <div className={styles.buttonRow} style={{ marginTop: 0 }}>
          <DefaultButton
            iconProps={{ iconName: 'Refresh' }}
            text="Opdater"
            onClick={onRefresh}
            disabled={isLoading}
          />
          <PrimaryButton
            iconProps={{ iconName: 'Delete' }}
            text="Slet alle registreringer"
            onClick={handleDeleteAllClick}
            disabled={isLoading || isDeleting || allRegistrations.length === 0}
            style={{ backgroundColor: '#c0392b', borderColor: '#c0392b' }}
          />
        </div>
      </div>

      {allRegistrations.length === 0 ? (
        <div className={styles.emptyState}>
          <Icon iconName="DocumentSet" className={styles.emptyIcon} />
          <p className={styles.emptyText}>Ingen registreringer i databasen</p>
        </div>
      ) : (
        <DetailsList
          items={allRegistrations}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
        />
      )}

      <Separator style={{ marginTop: 32, marginBottom: 16 }} />

      {/* Accrual History Section */}
      <h3 style={{ marginBottom: 8 }}>Optjeningshistorik (Accrual History)</h3>
      <div className={styles.registrationsHeader}>
        <span>
          <strong>{accrualHistoryCount}</strong> historik poster
        </span>
        <div className={styles.buttonRow} style={{ marginTop: 0 }}>
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            text="Tilføj optjening"
            onClick={() => setShowAddAccrualForm(true)}
            disabled={isLoading || isDeleting || isAddingAccrual}
          />
          <PrimaryButton
            iconProps={{ iconName: 'Delete' }}
            text="Slet alle historik"
            onClick={handleDeleteHistoryClick}
            disabled={isLoading || isDeleting || accrualHistoryCount === 0}
            style={{ backgroundColor: '#c0392b', borderColor: '#c0392b' }}
          />
        </div>
      </div>

      {/* Add Accrual History Modal */}
      <Modal
        isOpen={showAddAccrualForm}
        onDismiss={handleCancelAddAccrual}
        isBlocking={isAddingAccrual}
        containerClassName={styles.modalContainer}
        layerProps={{
          eventBubblingEnabled: true,
          hostId: undefined,
        }}
        styles={{
          root: { zIndex: 1000000 },
          main: { maxWidth: 700, minWidth: 500, maxHeight: '90vh' },
          scrollableContent: { overflowY: 'auto' },
        }}
      >
        <div style={{ padding: 24 }}>
          <h2 style={{ margin: '0 0 20px 0', fontSize: 20, fontWeight: 600 }}>Tilføj ny optjening</h2>

          {formError && (
            <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 20 }}>
              {formError}
            </MessageBar>
          )}

          {/* Section 1: Employee Selection */}
          <div style={{
            backgroundColor: '#fafafa',
            borderRadius: 6,
            padding: 16,
            marginBottom: 20,
          }}>
            <div style={{
              fontSize: 13,
              fontWeight: 600,
              color: '#00565a',
              textTransform: 'uppercase',
              letterSpacing: 0.5,
              marginBottom: 12,
            }}>
              <Icon iconName="Contact" style={{ marginRight: 8, fontSize: 14 }} />
              Medarbejder
            </div>
            <Stack horizontal tokens={{ childrenGap: 24 }}>
              <Stack.Item grow={1}>
                <div ref={searchInputRef} style={{ position: 'relative' }}>
                  <TextField
                    label="Søg medarbejder"
                    placeholder="Skriv navn eller e-mail..."
                    value={userSearchText}
                    onChange={handleUserSearchChange}
                    onBlur={handleSearchBlur}
                    onFocus={() => userSearchResults.length > 0 && setShowUserDropdown(true)}
                    required
                    disabled={isAddingAccrual}
                    iconProps={isSearching ? undefined : { iconName: 'Search' }}
                  />
                  {isSearching && (
                    <div style={{ position: 'absolute', right: 8, top: 32 }}>
                      <Spinner size={SpinnerSize.small} />
                    </div>
                  )}
                  {showUserDropdown && userSearchResults.length > 0 && searchInputRef.current && (
                    <Callout
                      target={searchInputRef.current}
                      isBeakVisible={false}
                      directionalHint={DirectionalHint.bottomLeftEdge}
                      onDismiss={() => setShowUserDropdown(false)}
                      setInitialFocus={false}
                      styles={{
                        root: { width: searchInputRef.current.offsetWidth },
                        calloutMain: { maxHeight: 300, overflowY: 'auto' },
                      }}
                    >
                      <List
                        items={userSearchResults}
                        onRenderCell={(user) => user && (
                          <div
                            style={{
                              padding: '10px 14px',
                              cursor: 'pointer',
                              borderBottom: '1px solid #edebe9',
                              transition: 'background-color 0.15s',
                            }}
                            onClick={() => handleUserSelect(user)}
                            onMouseDown={(e) => e.preventDefault()}
                            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#f3f2f1')}
                            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                          >
                            <Persona
                              text={user.displayName}
                              secondaryText={user.email}
                              tertiaryText={user.jobTitle}
                              size={PersonaSize.size32}
                              showSecondaryText
                            />
                          </div>
                        )}
                      />
                    </Callout>
                  )}
                </div>
              </Stack.Item>
              <Stack.Item grow={1}>
                <div style={{
                  backgroundColor: selectedUser ? '#e6f2f2' : '#f5f5f5',
                  borderRadius: 6,
                  padding: 12,
                  minHeight: 60,
                  display: 'flex',
                  alignItems: 'center',
                  border: selectedUser ? '1px solid #00565a' : '1px dashed #c8c8c8',
                }}>
                  {selectedUser ? (
                    <Persona
                      text={selectedUser.displayName}
                      secondaryText={selectedUser.email}
                      size={PersonaSize.size40}
                      showSecondaryText
                    />
                  ) : (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, color: '#a19f9d' }}>
                      <Icon iconName="Contact" style={{ fontSize: 20 }} />
                      <span style={{ fontSize: 14 }}>Ingen medarbejder valgt</span>
                    </div>
                  )}
                </div>
              </Stack.Item>
            </Stack>
          </div>

          {/* Section 2: Period & Type */}
          <div style={{
            backgroundColor: '#fafafa',
            borderRadius: 6,
            padding: 16,
            marginBottom: 20,
          }}>
            <div style={{
              fontSize: 13,
              fontWeight: 600,
              color: '#00565a',
              textTransform: 'uppercase',
              letterSpacing: 0.5,
              marginBottom: 12,
            }}>
              <Icon iconName="Calendar" style={{ marginRight: 8, fontSize: 14 }} />
              Periode & Type
            </div>
            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              <Stack.Item styles={{ root: { minWidth: 200, flex: 1 } }}>
                <Dropdown
                  label="Optjeningstype"
                  selectedKey={formState.accrualType}
                  options={accrualTypeOptions}
                  onChange={(_, option) => option && handleFormChange('accrualType', option.key as string)}
                  required
                  disabled={isAddingAccrual}
                />
              </Stack.Item>
              <Stack.Item styles={{ root: { minWidth: 150, flex: 1 } }}>
                <Dropdown
                  label="Ferieår"
                  selectedKey={formState.holidayYear}
                  options={generateHolidayYearOptions()}
                  onChange={(_, option) => option && handleFormChange('holidayYear', option.key as string)}
                  required
                  disabled={isAddingAccrual}
                />
              </Stack.Item>
              <Stack.Item styles={{ root: { minWidth: 180, flex: 1 } }}>
                <DatePicker
                  label="Optjeningsdato"
                  value={formState.accrualDate}
                  onSelectDate={(date) => date && handleFormChange('accrualDate', date)}
                  firstDayOfWeek={DayOfWeek.Monday}
                  placeholder="Vælg dato..."
                  ariaLabel="Vælg optjeningsdato"
                  disabled={isAddingAccrual}
                  isRequired
                />
              </Stack.Item>
            </Stack>
          </div>

          {/* Section 3: Days to Add */}
          <div style={{
            backgroundColor: '#fafafa',
            borderRadius: 6,
            padding: 16,
            marginBottom: 20,
          }}>
            <div style={{
              fontSize: 13,
              fontWeight: 600,
              color: '#00565a',
              textTransform: 'uppercase',
              letterSpacing: 0.5,
              marginBottom: 12,
            }}>
              <Icon iconName="NumberField" style={{ marginRight: 8, fontSize: 14 }} />
              Antal dage
            </div>
            <Stack horizontal tokens={{ childrenGap: 24 }}>
              <Stack.Item grow={1}>
                <div style={{
                  backgroundColor: '#ffffff',
                  borderRadius: 6,
                  padding: 16,
                  border: '1px solid #e0e0e0',
                  textAlign: 'center',
                }}>
                  <div style={{ fontSize: 12, color: '#605e5c', marginBottom: 8, fontWeight: 600 }}>
                    Feriedage (25 dage/år)
                  </div>
                  <SpinButton
                    value={formState.daysAccrued.toString()}
                    min={-50}
                    max={50}
                    step={0.5}
                    onChange={handleSpinButtonChange('daysAccrued')}
                    disabled={isAddingAccrual}
                    incrementButtonAriaLabel="Forøg"
                    decrementButtonAriaLabel="Formindsk"
                    styles={{
                      spinButtonWrapper: { width: '100%' },
                      input: { textAlign: 'center', fontSize: 18, fontWeight: 600 },
                    }}
                  />
                </div>
              </Stack.Item>
              <Stack.Item grow={1}>
                <div style={{
                  backgroundColor: '#ffffff',
                  borderRadius: 6,
                  padding: 16,
                  border: '1px solid #e0e0e0',
                  textAlign: 'center',
                }}>
                  <div style={{ fontSize: 12, color: '#605e5c', marginBottom: 8, fontWeight: 600 }}>
                    Feriefridage (5 dage/år)
                  </div>
                  <SpinButton
                    value={formState.feriefridageAccrued.toString()}
                    min={-20}
                    max={20}
                    step={0.5}
                    onChange={handleSpinButtonChange('feriefridageAccrued')}
                    disabled={isAddingAccrual}
                    incrementButtonAriaLabel="Forøg"
                    decrementButtonAriaLabel="Formindsk"
                    styles={{
                      spinButtonWrapper: { width: '100%' },
                      input: { textAlign: 'center', fontSize: 18, fontWeight: 600 },
                    }}
                  />
                </div>
              </Stack.Item>
            </Stack>
          </div>

          {/* Section 4: Notes */}
          <div style={{ marginBottom: 24 }}>
            <TextField
              label="Noter (valgfrit)"
              placeholder="Fx 'Startsaldo ved ansættelse', 'Overførsel fra forrige år'..."
              value={formState.notes}
              onChange={(_, val) => handleFormChange('notes', val || '')}
              disabled={isAddingAccrual}
              multiline
              rows={2}
              styles={{
                fieldGroup: { borderRadius: 6 },
              }}
            />
          </div>

          {/* Action Buttons */}
          <div style={{
            display: 'flex',
            justifyContent: 'flex-end',
            gap: 12,
            paddingTop: 16,
            borderTop: '1px solid #e0e0e0',
          }}>
            <DefaultButton
              text="Annuller"
              onClick={handleCancelAddAccrual}
              disabled={isAddingAccrual}
              styles={{ root: { minWidth: 100 } }}
            />
            <PrimaryButton
              text={isAddingAccrual ? 'Opretter...' : 'Opret optjening'}
              onClick={handleAddAccrualSubmit}
              disabled={isAddingAccrual}
              iconProps={{ iconName: 'Add' }}
              styles={{ root: { minWidth: 140 } }}
            />
          </div>
        </div>
      </Modal>

      {/* Delete All Registrations Dialog */}
      <Dialog
        hidden={!isDeleteAllDialogOpen}
        onDismiss={handleDialogDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Slet alle registreringer',
          subText: `Er du sikker på, at du vil slette alle ${allRegistrations.length} registreringer? Denne handling kan ikke fortrydes.`,
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDeleteAllConfirm}
            text={isDeleting ? 'Sletter...' : 'Ja, slet alle'}
            disabled={isDeleting}
            style={{ backgroundColor: '#c0392b', borderColor: '#c0392b' }}
          />
          <DefaultButton
            onClick={handleDialogDismiss}
            text="Annuller"
            disabled={isDeleting}
          />
        </DialogFooter>
      </Dialog>

      {/* Delete Single Dialog */}
      <Dialog
        hidden={!isDeleteSingleDialogOpen}
        onDismiss={handleDialogDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Slet registrering',
          subText: 'Er du sikker på, at du vil slette denne registrering? Denne handling kan ikke fortrydes.',
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDeleteSingleConfirm}
            text={isDeleting ? 'Sletter...' : 'Ja, slet'}
            disabled={isDeleting}
            style={{ backgroundColor: '#c0392b', borderColor: '#c0392b' }}
          />
          <DefaultButton
            onClick={handleDialogDismiss}
            text="Annuller"
            disabled={isDeleting}
          />
        </DialogFooter>
      </Dialog>

      {/* Delete Accrual History Dialog */}
      <Dialog
        hidden={!isDeleteHistoryDialogOpen}
        onDismiss={handleDialogDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Slet alle optjeningshistorik',
          subText: `Er du sikker på, at du vil slette alle ${accrualHistoryCount} historik poster? Denne handling kan ikke fortrydes.`,
        }}
        modalProps={{
          isBlocking: true,
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDeleteHistoryConfirm}
            text={isDeleting ? 'Sletter...' : 'Ja, slet alle'}
            disabled={isDeleting}
            style={{ backgroundColor: '#c0392b', borderColor: '#c0392b' }}
          />
          <DefaultButton
            onClick={handleDialogDismiss}
            text="Annuller"
            disabled={isDeleting}
          />
        </DialogFooter>
      </Dialog>

      {/* Admin Report Modal */}
      <Modal
        isOpen={showReport}
        onDismiss={() => setShowReport(false)}
        isBlocking={false}
        containerClassName={styles.modalContainer}
        styles={{
          root: { zIndex: 1000000 },
          main: { maxWidth: 1100, minWidth: 900, maxHeight: '90vh' },
          scrollableContent: { overflowY: 'auto', padding: 24 },
        }}
      >
        <AdminReport context={context} onClose={() => setShowReport(false)} />
      </Modal>
    </div>
  );
};

export default AdminPanel;
