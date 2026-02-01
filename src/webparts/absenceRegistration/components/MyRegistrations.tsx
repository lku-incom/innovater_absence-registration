import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import styles from './AbsenceRegistration.module.scss';
import { IMyRegistrationsProps } from './IAbsenceRegistrationProps';
import { IAbsenceRegistration, RegistrationStatus } from '../models/IAbsenceRegistration';
import { ICalculatedHolidayBalance, IAccrualHistory, getHolidayYear } from '../models/IHolidayBalance';
import { DataverseService } from '../services/DataverseService';
import HolidayBalanceCard from './HolidayBalanceCard';
import { Separator } from '@fluentui/react/lib/Separator';

const MyRegistrations: React.FC<IMyRegistrationsProps> = (props) => {
  const {
    context,
    registrations,
    onEdit,
    onView,
    onDelete,
    onSubmitForApproval,
    isLoading,
    onRefresh,
    currentUserEmail,
  } = props;

  const [deleteDialogHidden, setDeleteDialogHidden] = React.useState(true);
  const [itemToDelete, setItemToDelete] = React.useState<IAbsenceRegistration | undefined>();

  // Holiday balance state
  const [holidayBalance, setHolidayBalance] = useState<ICalculatedHolidayBalance | undefined>(undefined);
  const [isLoadingBalance, setIsLoadingBalance] = useState<boolean>(false);
  const [balanceError, setBalanceError] = useState<string | undefined>(undefined);

  // Accrual history state
  const [accrualHistory, setAccrualHistory] = useState<IAccrualHistory[]>([]);
  const [isLoadingAccruals, setIsLoadingAccruals] = useState<boolean>(false);

  // Fetch holiday balance and accrual history (can be called on mount and on refresh)
  const fetchHolidayBalance = useCallback(async (): Promise<void> => {
    if (!currentUserEmail || !context) {
      setHolidayBalance(undefined);
      setAccrualHistory([]);
      return;
    }

    setIsLoadingBalance(true);
    setIsLoadingAccruals(true);
    setBalanceError(undefined);

    try {
      const dataverseService = DataverseService.getInstance(context);
      const holidayYear = getHolidayYear();

      // Fetch both balance and accrual history in parallel
      const [balance, history] = await Promise.all([
        dataverseService.calculateHolidayBalanceForUser(currentUserEmail, holidayYear),
        dataverseService.getAccrualHistoryForUser(currentUserEmail, holidayYear),
      ]);

      setHolidayBalance(balance);
      setAccrualHistory(history);
    } catch {
      setBalanceError('Kunne ikke hente feriesaldo');
    } finally {
      setIsLoadingBalance(false);
      setIsLoadingAccruals(false);
    }
  }, [currentUserEmail, context]);

  // Fetch holiday balance when current user email changes
  useEffect(() => {
    fetchHolidayBalance();
  }, [fetchHolidayBalance]);

  // Handle refresh - refresh both registrations and holiday balance
  const handleRefresh = useCallback((): void => {
    onRefresh();
    fetchHolidayBalance();
  }, [onRefresh, fetchHolidayBalance]);

  const getStatusClass = (status: RegistrationStatus): string => {
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
    const d = date instanceof Date ? date : new Date(date);
    return d.toLocaleDateString('da-DK', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  };

  const handleDeleteClick = (item: IAbsenceRegistration): void => {
    setItemToDelete(item);
    setDeleteDialogHidden(false);
  };

  const handleDeleteConfirm = async (): Promise<void> => {
    if (itemToDelete?.DataverseId) {
      await onDelete(itemToDelete.DataverseId);
    }
    setDeleteDialogHidden(true);
    setItemToDelete(undefined);
  };

  const handleDeleteCancel = (): void => {
    setDeleteDialogHidden(true);
    setItemToDelete(undefined);
  };

  const columns: IColumn[] = [
    {
      key: 'absenceType',
      name: 'Fraværstype',
      fieldName: 'AbsenceType',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: 'startDate',
      name: 'Dato start',
      fieldName: 'StartDate',
      minWidth: 90,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => formatDate(item.StartDate),
    },
    {
      key: 'endDate',
      name: 'Dato slut',
      fieldName: 'EndDate',
      minWidth: 90,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => formatDate(item.EndDate),
    },
    {
      key: 'numberOfDays',
      name: 'Antal dage',
      fieldName: 'NumberOfDays',
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => (
        <span className={`${styles.statusBadge} ${getStatusClass(item.Status)}`}>
          {item.Status}
        </span>
      ),
    },
    {
      key: 'approver',
      name: 'Godkender',
      fieldName: 'ApproverName',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: 'actions',
      name: 'Handlinger',
      minWidth: 120,
      maxWidth: 150,
      onRender: (item: IAbsenceRegistration) => (
        <div className={styles.actionButtons}>
          {item.Status === 'Kladde' && (
            <>
              <TooltipHost content="Rediger">
                <IconButton
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => onEdit(item)}
                  ariaLabel="Rediger"
                />
              </TooltipHost>
              <TooltipHost content="Send til godkendelse">
                <IconButton
                  iconProps={{ iconName: 'Send' }}
                  onClick={() => item.DataverseId && onSubmitForApproval(item.DataverseId)}
                  ariaLabel="Send til godkendelse"
                />
              </TooltipHost>
              <TooltipHost content="Slet">
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  onClick={() => handleDeleteClick(item)}
                  ariaLabel="Slet"
                />
              </TooltipHost>
            </>
          )}
          {item.Status === 'Afventer godkendelse' && (
            <>
              <TooltipHost content="Vis detaljer">
                <IconButton
                  iconProps={{ iconName: 'View' }}
                  onClick={() => onView(item)}
                  ariaLabel="Vis detaljer"
                />
              </TooltipHost>
              <TooltipHost content="Afventer godkendelse">
                <Icon iconName="Clock" style={{ color: '#856404' }} />
              </TooltipHost>
            </>
          )}
          {item.Status === 'Godkendt' && (
            <>
              <TooltipHost content="Vis detaljer">
                <IconButton
                  iconProps={{ iconName: 'View' }}
                  onClick={() => onView(item)}
                  ariaLabel="Vis detaljer"
                />
              </TooltipHost>
              <TooltipHost content="Godkendt">
                <Icon iconName="CheckMark" style={{ color: '#155724' }} />
              </TooltipHost>
            </>
          )}
          {item.Status === 'Afvist' && (
            <>
              <TooltipHost content="Afvist - klik for at redigere">
                <IconButton
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => onEdit(item)}
                  ariaLabel="Rediger"
                />
              </TooltipHost>
              <TooltipHost content="Slet">
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  onClick={() => handleDeleteClick(item)}
                  ariaLabel="Slet"
                />
              </TooltipHost>
            </>
          )}
        </div>
      ),
    },
  ];

  // Columns for accrual history table
  const accrualColumns: IColumn[] = [
    {
      key: 'accrualDate',
      name: 'Dato',
      minWidth: 90,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAccrualHistory) => formatDate(item.AccrualDate),
    },
    {
      key: 'accrualType',
      name: 'Type',
      fieldName: 'AccrualType',
      minWidth: 140,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: 'holidayYear',
      name: 'Ferieår',
      fieldName: 'HolidayYear',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: 'daysAccrued',
      name: 'Feriedage',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAccrualHistory) => (
        <span style={{ color: item.DaysAccrued >= 0 ? '#155724' : '#721c24' }}>
          {item.DaysAccrued >= 0 ? '+' : ''}{item.DaysAccrued.toFixed(1)}
        </span>
      ),
    },
    {
      key: 'feriefridageAccrued',
      name: 'Feriefridage',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: IAccrualHistory) => {
        const value = item.FeriefridageAccrued || 0;
        if (value === 0) return <span style={{ color: '#999' }}>-</span>;
        return (
          <span style={{ color: value >= 0 ? '#155724' : '#721c24' }}>
            {value >= 0 ? '+' : ''}{value.toFixed(1)}
          </span>
        );
      },
    },
    {
      key: 'notes',
      name: 'Noter',
      fieldName: 'Notes',
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IAccrualHistory) => (
        <span style={{ color: '#666', fontStyle: item.Notes ? 'normal' : 'italic' }}>
          {item.Notes || '-'}
        </span>
      ),
    },
  ];

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Indlæser registreringer..." />
      </div>
    );
  }

  // Render accrual history section (reusable)
  const renderAccrualHistorySection = (): React.ReactNode => (
    <>
      <Separator style={{ marginTop: 32, marginBottom: 16 }} />

      <h3 style={{ margin: '0 0 16px 0', fontSize: 16, fontWeight: 600, color: '#00565a' }}>
        <Icon iconName="Money" style={{ marginRight: 8 }} />
        Optjeningshistorik ({getHolidayYear()})
      </h3>

      {isLoadingAccruals ? (
        <div style={{ padding: 20, textAlign: 'center' }}>
          <Spinner size={SpinnerSize.medium} label="Indlæser optjeningshistorik..." />
        </div>
      ) : accrualHistory.length === 0 ? (
        <div style={{
          padding: 24,
          textAlign: 'center',
          backgroundColor: '#fafafa',
          borderRadius: 6,
          color: '#666',
        }}>
          <Icon iconName="Info" style={{ fontSize: 24, marginBottom: 8, display: 'block' }} />
          <span>Ingen optjeningshistorik fundet for dette ferieår.</span>
        </div>
      ) : (
        <DetailsList
          items={accrualHistory}
          columns={accrualColumns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          compact={true}
        />
      )}
    </>
  );

  if (registrations.length === 0) {
    return (
      <div>
        {/* Holiday Balance Card */}
        <HolidayBalanceCard
          balance={holidayBalance}
          isLoading={isLoadingBalance}
          error={balanceError}
          registrations={registrations}
          accrualHistory={accrualHistory}
          employeeName={holidayBalance?.EmployeeName}
        />

        <div className={styles.emptyState}>
          <Icon iconName="Calendar" className={styles.emptyIcon} />
          <p className={styles.emptyText}>
            Du har ingen fraværsregistreringer endnu.
          </p>
          <p className={styles.emptyText}>
            Opret en ny registrering via fanen "Ny registrering".
          </p>
        </div>

        {/* Still show accrual history even if no registrations */}
        {renderAccrualHistorySection()}
      </div>
    );
  }

  return (
    <div>
      {/* Holiday Balance Card */}
      <HolidayBalanceCard
        balance={holidayBalance}
        isLoading={isLoadingBalance}
        error={balanceError}
        registrations={registrations}
        accrualHistory={accrualHistory}
        employeeName={holidayBalance?.EmployeeName}
      />

      <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'flex-end' }}>
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          title="Opdater"
          ariaLabel="Opdater"
          onClick={handleRefresh}
        />
      </div>

      <h3 style={{ margin: '0 0 16px 0', fontSize: 16, fontWeight: 600, color: '#00565a' }}>
        <Icon iconName="Calendar" style={{ marginRight: 8 }} />
        Fraværsregistreringer
      </h3>

      <DetailsList
        items={registrations}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
      />

      {/* Accrual History Section */}
      {renderAccrualHistorySection()}

      <Dialog
        hidden={deleteDialogHidden}
        onDismiss={handleDeleteCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Slet registrering',
          subText: `Er du sikker på, at du vil slette denne fraværsregistrering (${itemToDelete?.AbsenceType} - ${formatDate(itemToDelete?.StartDate)} til ${formatDate(itemToDelete?.EndDate)})?`,
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={handleDeleteConfirm} text="Slet" />
          <DefaultButton onClick={handleDeleteCancel} text="Annuller" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default MyRegistrations;
