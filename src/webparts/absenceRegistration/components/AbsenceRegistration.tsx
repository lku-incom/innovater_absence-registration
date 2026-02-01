import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import styles from './AbsenceRegistration.module.scss';
import { IAbsenceRegistrationProps } from './IAbsenceRegistrationProps';
import AbsenceForm from './AbsenceForm';
import MyRegistrations from './MyRegistrations';
import PendingApprovals from './PendingApprovals';
import AdminPanel from './AdminPanel';
import { IAbsenceRegistration, IUserInfo, RegistrationStatus } from '../models/IAbsenceRegistration';
import { GraphService } from '../services/GraphService';
import { DataverseService } from '../services/DataverseService';

// Admin users who can see the admin tab
const ADMIN_EMAILS = [
  'lkt@innovater.dk',
];

export interface IAbsenceRegistrationComponentProps extends IAbsenceRegistrationProps {
  dataverseUrl?: string;
}

const AbsenceRegistration: React.FC<IAbsenceRegistrationComponentProps> = (props) => {
  const { context, dataverseUrl } = props;

  const [activeTab, setActiveTab] = useState<string>('list');
  const [currentUser, setCurrentUser] = useState<IUserInfo | undefined>(undefined);
  const [registrations, setRegistrations] = useState<IAbsenceRegistration[]>([]);
  const [pendingApprovals, setPendingApprovals] = useState<IAbsenceRegistration[]>([]);
  const [allRegistrations, setAllRegistrations] = useState<IAbsenceRegistration[]>([]);
  const [editingRegistration, setEditingRegistration] = useState<IAbsenceRegistration | undefined>(undefined);
  const [isViewOnly, setIsViewOnly] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isLoadingApprovals, setIsLoadingApprovals] = useState<boolean>(false);
  const [isLoadingAdmin, setIsLoadingAdmin] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [accrualHistoryCount, setAccrualHistoryCount] = useState<number>(0);

  // Check if current user is an admin
  const isAdmin = ADMIN_EMAILS.some(
    (email) => email.toLowerCase() === context.pageContext.user.email.toLowerCase()
  );

  // Get status class for the status badge
  const getStatusClass = (status: RegistrationStatus | string): string => {
    switch (status) {
      case 'Afventer godkendelse':
        return styles.pending;
      case 'Godkendt':
        return styles.approved;
      case 'Afvist':
        return styles.rejected;
      case 'Kladde':
      default:
        return styles.draft;
    }
  };

  // Initialize services
  useEffect(() => {
    const graphService = GraphService.getInstance(context);
    const dataverseService = DataverseService.getInstance(context);

    if (dataverseUrl) {
      dataverseService.setDataverseUrl(dataverseUrl);
    }

    const initializeData = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setErrorMessage('');

        // Get current user info including manager
        let userInfo;
        try {
          userInfo = await graphService.getFullUserInfo();
          setCurrentUser(userInfo);
        } catch (graphError) {
          console.error('Graph API error:', graphError);
          // Create minimal user info from context if Graph fails
          userInfo = {
            id: 0,
            email: context.pageContext.user.email,
            displayName: context.pageContext.user.displayName,
            manager: undefined,
          };
          setCurrentUser(userInfo);
          console.warn('Using fallback user info from context');
        }

        // Load user's registrations and pending approvals
        if (userInfo.email) {
          try {
            const [userRegistrations, approvals] = await Promise.all([
              dataverseService.getMyRegistrations(userInfo.email),
              dataverseService.getPendingApprovals(userInfo.email),
            ]);
            setRegistrations(userRegistrations);
            setPendingApprovals(approvals);
          } catch (dataverseError) {
            console.error('Dataverse error:', dataverseError);
            setRegistrations([]);
            setPendingApprovals([]);
          }
        }
      } catch (error) {
        console.error('Error initializing:', error);
        setErrorMessage('Kunne ikke indlæse data. Prøv venligst igen.');
      } finally {
        setIsLoading(false);
      }
    };

    initializeData().catch(console.error);
  }, [context, dataverseUrl]);

  const loadRegistrations = useCallback(async (): Promise<void> => {
    if (!currentUser?.email) return;

    try {
      setIsLoading(true);
      const dataverseService = DataverseService.getInstance();
      const userRegistrations = await dataverseService.getMyRegistrations(currentUser.email);
      setRegistrations(userRegistrations);
    } catch (error) {
      console.error('Error loading registrations:', error);
      setErrorMessage('Kunne ikke indlæse registreringer.');
    } finally {
      setIsLoading(false);
    }
  }, [currentUser]);

  const loadPendingApprovals = useCallback(async (): Promise<void> => {
    if (!currentUser?.email) return;

    try {
      setIsLoadingApprovals(true);
      const dataverseService = DataverseService.getInstance();
      const approvals = await dataverseService.getPendingApprovals(currentUser.email);
      setPendingApprovals(approvals);
    } catch (error) {
      console.error('Error loading pending approvals:', error);
      setErrorMessage('Kunne ikke indlæse godkendelser.');
    } finally {
      setIsLoadingApprovals(false);
    }
  }, [currentUser]);

  const loadAllRegistrations = useCallback(async (): Promise<void> => {
    if (!isAdmin) return;

    try {
      setIsLoadingAdmin(true);
      const dataverseService = DataverseService.getInstance();
      const [all, historyCount] = await Promise.all([
        dataverseService.getAllRegistrations(),
        dataverseService.getAccrualHistoryCount(),
      ]);
      setAllRegistrations(all);
      setAccrualHistoryCount(historyCount);
    } catch (error) {
      console.error('Error loading all registrations:', error);
      setErrorMessage('Kunne ikke indlæse alle registreringer.');
    } finally {
      setIsLoadingAdmin(false);
    }
  }, [isAdmin]);

  const handleSave = async (registration: IAbsenceRegistration): Promise<void> => {
    try {
      setIsSaving(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();

      if (editingRegistration?.DataverseId) {
        await dataverseService.updateRegistration(editingRegistration.DataverseId, registration);
      } else {
        await dataverseService.createRegistration({
          ...registration,
          Status: 'Kladde',
        });
      }

      await loadRegistrations();
      setEditingRegistration(undefined);
      setIsViewOnly(false);
      setActiveTab('list');
    } catch (error) {
      console.error('Error saving registration:', error);
      const errorMsg = error instanceof Error ? error.message : 'Ukendt fejl';
      setErrorMessage(`Kunne ikke gemme registrering: ${errorMsg}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSubmit = async (registration: IAbsenceRegistration): Promise<void> => {
    try {
      setIsSaving(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();

      if (editingRegistration?.DataverseId) {
        await dataverseService.updateRegistration(editingRegistration.DataverseId, {
          ...registration,
          Status: 'Afventer godkendelse',
        });
      } else {
        await dataverseService.createRegistration({
          ...registration,
          Status: 'Afventer godkendelse',
        });
      }

      await loadRegistrations();
      setEditingRegistration(undefined);
      setIsViewOnly(false);
      setActiveTab('list');
    } catch (error) {
      console.error('Error submitting registration:', error);
      const errorMsg = error instanceof Error ? error.message : 'Ukendt fejl';
      setErrorMessage(`Kunne ikke indsende registrering: ${errorMsg}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancel = (): void => {
    setEditingRegistration(undefined);
    setIsViewOnly(false);
    if (registrations.length > 0) {
      setActiveTab('list');
    }
  };

  const handleEdit = (registration: IAbsenceRegistration): void => {
    setEditingRegistration(registration);
    setIsViewOnly(false);
    setActiveTab('new');
  };

  const handleView = (registration: IAbsenceRegistration): void => {
    setEditingRegistration(registration);
    setIsViewOnly(true);
    setActiveTab('new');
  };

  const handleDelete = async (id: string): Promise<void> => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      await dataverseService.deleteRegistration(id);
      await loadRegistrations();
    } catch (error) {
      console.error('Error deleting registration:', error);
      setErrorMessage('Kunne ikke slette registrering. Prøv venligst igen.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSubmitForApproval = async (id: string): Promise<void> => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      await dataverseService.submitForApproval(id);
      await loadRegistrations();
    } catch (error) {
      console.error('Error submitting for approval:', error);
      setErrorMessage('Kunne ikke indsende til godkendelse. Prøv venligst igen.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteAllRegistrations = async (): Promise<void> => {
    try {
      setIsLoadingAdmin(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      await dataverseService.deleteAllRegistrations();
      await loadAllRegistrations();
      await loadRegistrations();
    } catch (error) {
      console.error('Error deleting all registrations:', error);
      setErrorMessage('Kunne ikke slette alle registreringer. Prøv venligst igen.');
    } finally {
      setIsLoadingAdmin(false);
    }
  };

  const handleAdminDeleteSingle = async (id: string): Promise<void> => {
    try {
      setIsLoadingAdmin(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      await dataverseService.deleteRegistration(id);
      await loadAllRegistrations();
      await loadRegistrations();
    } catch (error) {
      console.error('Error deleting registration:', error);
      setErrorMessage('Kunne ikke slette registrering. Prøv venligst igen.');
    } finally {
      setIsLoadingAdmin(false);
    }
  };

  const handleDeleteAllAccrualHistory = async (): Promise<number> => {
    try {
      setIsLoadingAdmin(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      const deletedCount = await dataverseService.deleteAllAccrualHistory();
      await loadAllRegistrations();
      return deletedCount;
    } catch (error) {
      console.error('Error deleting all accrual history:', error);
      setErrorMessage('Kunne ikke slette al optjeningshistorik. Prøv venligst igen.');
      throw error;
    } finally {
      setIsLoadingAdmin(false);
    }
  };

  const handleCreateAccrualHistory = async (
    accrual: Parameters<typeof DataverseService.prototype.createAccrualHistory>[0]
  ): Promise<void> => {
    try {
      setIsLoadingAdmin(true);
      setErrorMessage('');

      const dataverseService = DataverseService.getInstance();
      await dataverseService.createAccrualHistory(accrual);
      await loadAllRegistrations();
    } catch (error) {
      console.error('Error creating accrual history:', error);
      setErrorMessage('Kunne ikke oprette optjening. Prøv venligst igen.');
      throw error;
    } finally {
      setIsLoadingAdmin(false);
    }
  };

  const handleTabChange = (item?: PivotItem): void => {
    if (item) {
      setActiveTab(item.props.itemKey || 'new');
      if (item.props.itemKey === 'new' && !editingRegistration) {
        setEditingRegistration(undefined);
        setIsViewOnly(false);
      }
      if (item.props.itemKey === 'list') {
        setEditingRegistration(undefined);
        setIsViewOnly(false);
      }
      if (item.props.itemKey === 'approvals') {
        setEditingRegistration(undefined);
        setIsViewOnly(false);
      }
      if (item.props.itemKey === 'admin' && isAdmin) {
        setEditingRegistration(undefined);
        setIsViewOnly(false);
        loadAllRegistrations().catch(console.error);
      }
    }
  };

  if (isLoading && !currentUser) {
    return (
      <div className={styles.absenceRegistration}>
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label="Indlæser..." />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.absenceRegistration}>
      {/* Error message */}
      {errorMessage && (
        <div className={styles.errorBanner}>
          {errorMessage}
        </div>
      )}

      {/* Tabs */}
      <div className={styles.tabs}>
        {/* Help button */}
        <div style={{ position: 'absolute', right: 16, top: 8, zIndex: 10 }}>
          <TooltipHost content="Åbn dokumentation">
            <IconButton
              iconProps={{ iconName: 'BookAnswers' }}
              title="Dokumentation"
              ariaLabel="Åbn dokumentation"
              onClick={() => {
                const docUrl = `${context.pageContext.web.absoluteUrl}/SiteAssets/AbsenceRegistrationDocumentation.html`;
                window.open(docUrl, '_blank');
              }}
              styles={{
                root: {
                  color: '#004e6b',
                  backgroundColor: 'transparent',
                },
                rootHovered: {
                  color: '#006d9c',
                  backgroundColor: 'rgba(0, 78, 107, 0.1)',
                },
                icon: {
                  fontSize: 18,
                },
              }}
            />
          </TooltipHost>
        </div>
        <Pivot
          selectedKey={activeTab}
          onLinkClick={handleTabChange}
          headersOnly={false}
        >
          <PivotItem
            headerText="MINE REGISTRERINGER"
            itemKey="list"
            itemIcon="BulletedList"
          >
            <div className={styles.tabContent}>
              <MyRegistrations
                context={context}
                registrations={registrations}
                onEdit={handleEdit}
                onView={handleView}
                onDelete={handleDelete}
                onSubmitForApproval={handleSubmitForApproval}
                isLoading={isLoading}
                onRefresh={loadRegistrations}
                currentUserEmail={currentUser?.email}
              />
            </div>
          </PivotItem>

          <PivotItem
            headerText="NY REGISTRERING"
            itemKey="new"
            itemIcon="Calendar"
          >
            <div className={styles.tabContent}>
              {/* Status indicator - only shown on registration form */}
              <div className={styles.statusSection} style={{ justifyContent: 'flex-start', paddingLeft: 0 }}>
                <span className={`${styles.statusBadgeLarge} ${getStatusClass(editingRegistration?.Status || 'Kladde')}`}>
                  {editingRegistration?.Status || 'Kladde'}
                </span>
              </div>
              <AbsenceForm
                context={context}
                currentUser={currentUser}
                onSave={handleSave}
                onSubmit={handleSubmit}
                onCancel={handleCancel}
                editingRegistration={editingRegistration}
                isLoading={isSaving}
                readOnly={isViewOnly}
              />
            </div>
          </PivotItem>

          <PivotItem
            headerText="GODKENDELSER"
            itemKey="approvals"
            itemIcon="CheckList"
            itemCount={pendingApprovals.length > 0 ? pendingApprovals.length : undefined}
          >
            <div className={styles.tabContent}>
              <PendingApprovals
                context={context}
                pendingApprovals={pendingApprovals}
                onView={handleView}
                isLoading={isLoadingApprovals}
                onRefresh={loadPendingApprovals}
              />
            </div>
          </PivotItem>

          {isAdmin && (
            <PivotItem
              headerText="ADMIN"
              itemKey="admin"
              itemIcon="Settings"
            >
              <div className={styles.tabContent}>
                <AdminPanel
                  context={context}
                  allRegistrations={allRegistrations}
                  onDeleteAll={handleDeleteAllRegistrations}
                  onDeleteSingle={handleAdminDeleteSingle}
                  onDeleteAllAccrualHistory={handleDeleteAllAccrualHistory}
                  onCreateAccrualHistory={handleCreateAccrualHistory}
                  accrualHistoryCount={accrualHistoryCount}
                  isLoading={isLoadingAdmin}
                  onRefresh={loadAllRegistrations}
                />
              </div>
            </PivotItem>
          )}
        </Pivot>
      </div>
    </div>
  );
};

export default AbsenceRegistration;
