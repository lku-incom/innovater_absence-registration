import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './AbsenceRegistration.module.scss';
import { ICalculatedHolidayBalance, IAccrualHistory } from '../models/IHolidayBalance';
import { IAbsenceRegistration } from '../models/IAbsenceRegistration';
import { ExcelExportService } from '../services/ExcelExportService';

export interface IHolidayBalanceCardProps {
  balance: ICalculatedHolidayBalance | undefined;
  isLoading: boolean;
  error?: string;
  // Optional data for Excel export
  registrations?: IAbsenceRegistration[];
  accrualHistory?: IAccrualHistory[];
  employeeName?: string;
}

const HolidayBalanceCard: React.FC<IHolidayBalanceCardProps> = ({
  balance,
  isLoading,
  error,
  registrations = [],
  accrualHistory = [],
  employeeName = '',
}) => {
  // Handle Excel export
  const handleExportToExcel = (): void => {
    const name = employeeName || balance?.EmployeeName || 'Medarbejder';
    ExcelExportService.exportUserHolidayData(name, balance, registrations, accrualHistory);
  };

  // Check if export is available (has data to export)
  const canExport = balance || registrations.length > 0 || accrualHistory.length > 0;
  if (isLoading) {
    return (
      <div className={styles.balanceCard}>
        <div className={styles.balanceCardHeader}>
          <Icon iconName="Calendar" className={styles.balanceIcon} />
          <span>Feriesaldo</span>
        </div>
        <div className={styles.balanceCardLoading}>
          <Spinner size={SpinnerSize.small} label="Indlæser saldo..." />
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.balanceCard}>
        <div className={styles.balanceCardHeader}>
          <Icon iconName="Calendar" className={styles.balanceIcon} />
          <span>Feriesaldo</span>
        </div>
        <div className={styles.balanceCardError}>
          <Icon iconName="Warning" />
          <span>{error}</span>
        </div>
      </div>
    );
  }

  if (!balance) {
    return (
      <div className={styles.balanceCard}>
        <div className={styles.balanceCardHeader}>
          <Icon iconName="Calendar" className={styles.balanceIcon} />
          <span>Feriesaldo</span>
        </div>
        <div className={styles.balanceCardEmpty}>
          <Icon iconName="Info" />
          <span>Ingen feriesaldo fundet. Kontakt din administrator for at få oprettet din startsaldo.</span>
        </div>
      </div>
    );
  }

  // Format numbers to show max 1 decimal
  const formatDays = (days: number): string => {
    return days.toFixed(1);
  };

  return (
    <div className={styles.balanceCard}>
      <div className={styles.balanceCardHeader}>
        <Icon iconName="Calendar" className={styles.balanceIcon} />
        <span>Feriesaldo {balance.HolidayYear}</span>
        <TooltipHost
          content={`Ferieår: 1. september ${balance.HolidayYear.split('-')[0]} til 31. august ${balance.HolidayYear.split('-')[1]}`}
        >
          <Icon iconName="Info" className={styles.infoIconSmall} />
        </TooltipHost>
        {canExport && (
          <TooltipHost content="Download feriedata til Excel">
            <IconButton
              iconProps={{ iconName: 'ExcelDocument' }}
              onClick={handleExportToExcel}
              styles={{
                root: {
                  marginLeft: 'auto',
                  color: 'white',
                  height: 28,
                  width: 28,
                },
                rootHovered: {
                  backgroundColor: 'rgba(255, 255, 255, 0.2)',
                  color: 'white',
                },
                icon: {
                  fontSize: 16,
                },
              }}
              ariaLabel="Download til Excel"
            />
          </TooltipHost>
        )}
      </div>

      <div className={styles.balanceCardContent}>
        {/* Feriedage (statutory vacation) */}
        <div className={styles.balanceSection}>
          <div className={styles.balanceSectionTitle}>
            Feriedage
            <TooltipHost content="Lovpligtig ferie (25 dage/år)">
              <Icon iconName="Info" className={styles.infoIconSmall} />
            </TooltipHost>
          </div>
          <div className={styles.balanceRow}>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Til rådighed</span>
              <span className={`${styles.balanceValue} ${balance.AvailableDays >= 0 ? styles.primary : styles.negative}`}>
                {formatDays(balance.AvailableDays)}
              </span>
            </div>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Optjent</span>
              <span className={styles.balanceValue}>{formatDays(balance.TotalAccruedDays)}</span>
            </div>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Brugt</span>
              <span className={styles.balanceValue}>{formatDays(balance.UsedDays)}</span>
            </div>
            {balance.PendingDays > 0 && (
              <div className={styles.balanceItem}>
                <span className={styles.balanceLabel}>Afventer</span>
                <span className={`${styles.balanceValue} ${styles.pending}`}>
                  {formatDays(balance.PendingDays)}
                </span>
              </div>
            )}
          </div>
        </div>

        {/* Feriefridage (extra days) */}
        <div className={styles.balanceSection}>
          <div className={styles.balanceSectionTitle}>
            Feriefridage
            <TooltipHost content="Kontraktbaserede fridage (typisk 5/år)">
              <Icon iconName="Info" className={styles.infoIconSmall} />
            </TooltipHost>
          </div>
          <div className={styles.balanceRow}>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Til rådighed</span>
              <span className={`${styles.balanceValue} ${balance.AvailableFeriefridage >= 0 ? styles.primary : styles.negative}`}>
                {formatDays(balance.AvailableFeriefridage)}
              </span>
            </div>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Optjent</span>
              <span className={styles.balanceValue}>{formatDays(balance.TotalAccruedFeriefridage)}</span>
            </div>
            <div className={styles.balanceItem}>
              <span className={styles.balanceLabel}>Brugt</span>
              <span className={styles.balanceValue}>{formatDays(balance.UsedFeriefridage)}</span>
            </div>
            {balance.PendingFeriefridage > 0 && (
              <div className={styles.balanceItem}>
                <span className={styles.balanceLabel}>Afventer</span>
                <span className={`${styles.balanceValue} ${styles.pending}`}>
                  {formatDays(balance.PendingFeriefridage)}
                </span>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default HolidayBalanceCard;
