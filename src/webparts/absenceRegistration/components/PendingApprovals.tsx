import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import styles from './AbsenceRegistration.module.scss';
import { IPendingApprovalsProps } from './IAbsenceRegistrationProps';
import { IAbsenceRegistration } from '../models/IAbsenceRegistration';

const PendingApprovals: React.FC<IPendingApprovalsProps> = (props) => {
  const {
    pendingApprovals,
    onView,
    isLoading,
    onRefresh,
  } = props;

  const formatDate = (date: Date | undefined): string => {
    if (!date) return '-';
    const d = date instanceof Date ? date : new Date(date);
    return d.toLocaleDateString('da-DK', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  };

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
      key: 'notes',
      name: 'Noter',
      fieldName: 'Notes',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAbsenceRegistration) => (
        <span title={item.Notes || ''}>
          {item.Notes ? (item.Notes.length > 30 ? `${item.Notes.substring(0, 30)}...` : item.Notes) : '-'}
        </span>
      ),
    },
    {
      key: 'actions',
      name: '',
      minWidth: 50,
      maxWidth: 50,
      onRender: (item: IAbsenceRegistration) => (
        <div className={styles.actionButtons}>
          <TooltipHost content="Vis detaljer">
            <IconButton
              iconProps={{ iconName: 'View' }}
              onClick={() => onView(item)}
              ariaLabel="Vis detaljer"
            />
          </TooltipHost>
        </div>
      ),
    },
  ];

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Indlæser godkendelser..." />
      </div>
    );
  }

  if (pendingApprovals.length === 0) {
    return (
      <div className={styles.emptyState}>
        <Icon iconName="Inbox" className={styles.emptyIcon} />
        <p className={styles.emptyText}>
          Du har ingen afventende godkendelser.
        </p>
        <p className={styles.emptyText}>
          Når medarbejdere sender fraværsanmodninger til dig, vises de her.
        </p>
      </div>
    );
  }

  return (
    <div>
      <MessageBar messageBarType={MessageBarType.info} style={{ marginBottom: 16 }}>
        <strong>Godkendelse via Teams/e-mail:</strong> Fraværsanmodninger godkendes via Microsoft Teams eller e-mail.
        Du modtager en notifikation, når en medarbejder indsender en anmodning.
        <br />
        <a
          href="https://make.powerautomate.com/approvals/received"
          target="_blank"
          rel="noopener noreferrer"
          style={{ marginTop: 8, display: 'inline-block' }}
        >
          Åbn godkendelser i Power Automate →
        </a>
      </MessageBar>

      <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span><strong>{pendingApprovals.length}</strong> afventende godkendelse{pendingApprovals.length !== 1 ? 'r' : ''}</span>
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          title="Opdater"
          ariaLabel="Opdater"
          onClick={onRefresh}
        />
      </div>

      <DetailsList
        items={pendingApprovals}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
      />
    </div>
  );
};

export default PendingApprovals;
