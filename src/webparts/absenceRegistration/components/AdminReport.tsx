import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import styles from './AbsenceRegistration.module.scss';
import { DataverseService } from '../services/DataverseService';
import { ICalculatedHolidayBalance, getHolidayYear } from '../models/IHolidayBalance';

export interface IAdminReportProps {
  context: WebPartContext;
  onClose: () => void;
}

interface IEmployeeBalance extends ICalculatedHolidayBalance {
  // Visual helper fields
  feriedagePercentUsed: number;
  feriefridagePercentUsed: number;
}

interface IAggregatedStats {
  totalEmployees: number;
  totalAccruedDays: number;
  totalUsedDays: number;
  totalPendingDays: number;
  totalAvailableDays: number;
  totalAccruedFeriefridage: number;
  totalUsedFeriefridage: number;
  totalPendingFeriefridage: number;
  totalAvailableFeriefridage: number;
}

const AdminReport: React.FC<IAdminReportProps> = ({ context, onClose }) => {
  const [employeeBalances, setEmployeeBalances] = useState<IEmployeeBalance[]>([]);
  const [stats, setStats] = useState<IAggregatedStats | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const holidayYear = getHolidayYear();

  // Fetch all employee balances
  const fetchData = useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);

    try {
      const dataverseService = DataverseService.getInstance(context);
      const balances = await dataverseService.getAllEmployeeBalances(holidayYear);

      // Calculate percentages for visual bars
      const enrichedBalances: IEmployeeBalance[] = balances.map((b) => ({
        ...b,
        feriedagePercentUsed:
          b.TotalAccruedDays > 0
            ? Math.min(100, ((b.UsedDays + b.PendingDays) / b.TotalAccruedDays) * 100)
            : 0,
        feriefridagePercentUsed:
          b.TotalAccruedFeriefridage > 0
            ? Math.min(100, ((b.UsedFeriefridage + b.PendingFeriefridage) / b.TotalAccruedFeriefridage) * 100)
            : 0,
      }));

      // Sort by employee name
      enrichedBalances.sort((a, b) => a.EmployeeName.localeCompare(b.EmployeeName, 'da'));

      setEmployeeBalances(enrichedBalances);

      // Calculate aggregated stats
      const aggregated: IAggregatedStats = {
        totalEmployees: balances.length,
        totalAccruedDays: balances.reduce((sum, b) => sum + b.TotalAccruedDays, 0),
        totalUsedDays: balances.reduce((sum, b) => sum + b.UsedDays, 0),
        totalPendingDays: balances.reduce((sum, b) => sum + b.PendingDays, 0),
        totalAvailableDays: balances.reduce((sum, b) => sum + b.AvailableDays, 0),
        totalAccruedFeriefridage: balances.reduce((sum, b) => sum + b.TotalAccruedFeriefridage, 0),
        totalUsedFeriefridage: balances.reduce((sum, b) => sum + b.UsedFeriefridage, 0),
        totalPendingFeriefridage: balances.reduce((sum, b) => sum + b.PendingFeriefridage, 0),
        totalAvailableFeriefridage: balances.reduce((sum, b) => sum + b.AvailableFeriefridage, 0),
      };
      setStats(aggregated);
    } catch (err) {
      console.error('Error fetching employee balances:', err);
      setError('Kunne ikke hente medarbejderdata. Prøv igen senere.');
    } finally {
      setIsLoading(false);
    }
  }, [context, holidayYear]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  // Format number to 1 decimal
  const formatNumber = (num: number): string => num.toFixed(1);

  // Render balance bar
  const renderBalanceBar = (
    used: number,
    pending: number,
    total: number,
    type: 'feriedage' | 'feriefridage'
  ): React.ReactNode => {
    if (total === 0) {
      return <span style={{ color: '#999', fontStyle: 'italic', fontSize: 12 }}>Ingen data</span>;
    }

    const usedPercent = Math.min(100, (used / total) * 100);
    const pendingPercent = Math.min(100 - usedPercent, (pending / total) * 100);
    const available = total - used - pending;

    const usedColor = type === 'feriedage' ? '#004e6b' : '#5a6e3a';
    const pendingColor = type === 'feriedage' ? '#f5a623' : '#c9a227';

    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <div
          style={{
            width: 80,
            height: 12,
            backgroundColor: '#e0e0e0',
            borderRadius: 6,
            overflow: 'hidden',
            display: 'flex',
          }}
        >
          <div
            style={{
              width: `${usedPercent}%`,
              backgroundColor: usedColor,
              height: '100%',
            }}
          />
          <div
            style={{
              width: `${pendingPercent}%`,
              backgroundColor: pendingColor,
              height: '100%',
            }}
          />
        </div>
        <span style={{ fontSize: 12, color: available >= 0 ? '#333' : '#c0392b', fontWeight: available < 0 ? 600 : 400 }}>
          {formatNumber(available)}
        </span>
      </div>
    );
  };

  // Column definitions
  const columns: IColumn[] = [
    {
      key: 'employeeName',
      name: 'Medarbejder',
      fieldName: 'EmployeeName',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'feriedage',
      name: 'Feriedage',
      minWidth: 140,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: IEmployeeBalance) =>
        renderBalanceBar(item.UsedDays, item.PendingDays, item.TotalAccruedDays, 'feriedage'),
    },
    {
      key: 'feriedageDetail',
      name: 'Optjent / Brugt',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IEmployeeBalance) => (
        <span style={{ fontSize: 12, color: '#666' }}>
          {formatNumber(item.TotalAccruedDays)} / {formatNumber(item.UsedDays)}
          {item.PendingDays > 0 && (
            <span style={{ color: '#f5a623' }}> (+{formatNumber(item.PendingDays)})</span>
          )}
        </span>
      ),
    },
    {
      key: 'feriefridage',
      name: 'Feriefridage',
      minWidth: 140,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: IEmployeeBalance) =>
        renderBalanceBar(
          item.UsedFeriefridage,
          item.PendingFeriefridage,
          item.TotalAccruedFeriefridage,
          'feriefridage'
        ),
    },
    {
      key: 'feriefridageDetail',
      name: 'Optjent / Brugt',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IEmployeeBalance) => (
        <span style={{ fontSize: 12, color: '#666' }}>
          {formatNumber(item.TotalAccruedFeriefridage)} / {formatNumber(item.UsedFeriefridage)}
          {item.PendingFeriefridage > 0 && (
            <span style={{ color: '#c9a227' }}> (+{formatNumber(item.PendingFeriefridage)})</span>
          )}
        </span>
      ),
    },
  ];

  if (isLoading) {
    return (
      <div style={{ padding: 40, textAlign: 'center' }}>
        <Spinner size={SpinnerSize.large} label="Indlæser rapport..." />
      </div>
    );
  }

  return (
    <div style={{ padding: 0 }}>
      {/* Header */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          marginBottom: 24,
          paddingBottom: 16,
          borderBottom: '1px solid #e0e0e0',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <Icon iconName="ReportDocument" style={{ fontSize: 24, color: '#004e6b' }} />
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 600, color: '#333' }}>
              Ferieoversigt
            </h2>
            <span style={{ fontSize: 13, color: '#666' }}>Ferieår {holidayYear}</span>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <TooltipHost content="Opdater data">
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={fetchData}
              ariaLabel="Opdater"
            />
          </TooltipHost>
          <DefaultButton text="Luk" onClick={onClose} iconProps={{ iconName: 'Cancel' }} />
        </div>
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 16 }}>
          {error}
        </MessageBar>
      )}

      {/* Summary Cards */}
      {stats && (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
            gap: 16,
            marginBottom: 24,
          }}
        >
          {/* Total Employees */}
          <div
            style={{
              background: 'linear-gradient(135deg, #004e6b 0%, #006d8f 100%)',
              borderRadius: 8,
              padding: 16,
              color: 'white',
            }}
          >
            <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 4 }}>Medarbejdere</div>
            <div style={{ fontSize: 28, fontWeight: 600 }}>{stats.totalEmployees}</div>
          </div>

          {/* Feriedage Summary */}
          <div
            style={{
              background: '#f5f5f5',
              borderRadius: 8,
              padding: 16,
              borderLeft: '4px solid #004e6b',
            }}
          >
            <div style={{ fontSize: 12, color: '#666', marginBottom: 8 }}>Feriedage (total)</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
              <div>
                <span style={{ fontSize: 24, fontWeight: 600, color: '#004e6b' }}>
                  {formatNumber(stats.totalAvailableDays)}
                </span>
                <span style={{ fontSize: 12, color: '#666', marginLeft: 4 }}>til rådighed</span>
              </div>
            </div>
            <div style={{ fontSize: 11, color: '#999', marginTop: 8 }}>
              Optjent: {formatNumber(stats.totalAccruedDays)} | Brugt: {formatNumber(stats.totalUsedDays)}
              {stats.totalPendingDays > 0 && ` | Afventer: ${formatNumber(stats.totalPendingDays)}`}
            </div>
          </div>

          {/* Feriefridage Summary */}
          <div
            style={{
              background: '#f5f5f5',
              borderRadius: 8,
              padding: 16,
              borderLeft: '4px solid #5a6e3a',
            }}
          >
            <div style={{ fontSize: 12, color: '#666', marginBottom: 8 }}>Feriefridage (total)</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
              <div>
                <span style={{ fontSize: 24, fontWeight: 600, color: '#5a6e3a' }}>
                  {formatNumber(stats.totalAvailableFeriefridage)}
                </span>
                <span style={{ fontSize: 12, color: '#666', marginLeft: 4 }}>til rådighed</span>
              </div>
            </div>
            <div style={{ fontSize: 11, color: '#999', marginTop: 8 }}>
              Optjent: {formatNumber(stats.totalAccruedFeriefridage)} | Brugt:{' '}
              {formatNumber(stats.totalUsedFeriefridage)}
              {stats.totalPendingFeriefridage > 0 &&
                ` | Afventer: ${formatNumber(stats.totalPendingFeriefridage)}`}
            </div>
          </div>

          {/* Usage Percentage */}
          <div
            style={{
              background: '#f5f5f5',
              borderRadius: 8,
              padding: 16,
            }}
          >
            <div style={{ fontSize: 12, color: '#666', marginBottom: 8 }}>Forbrugsrate</div>
            <div style={{ display: 'flex', gap: 16 }}>
              <div>
                <div style={{ fontSize: 20, fontWeight: 600, color: '#004e6b' }}>
                  {stats.totalAccruedDays > 0
                    ? Math.round((stats.totalUsedDays / stats.totalAccruedDays) * 100)
                    : 0}
                  %
                </div>
                <div style={{ fontSize: 11, color: '#666' }}>Feriedage</div>
              </div>
              <div>
                <div style={{ fontSize: 20, fontWeight: 600, color: '#5a6e3a' }}>
                  {stats.totalAccruedFeriefridage > 0
                    ? Math.round((stats.totalUsedFeriefridage / stats.totalAccruedFeriefridage) * 100)
                    : 0}
                  %
                </div>
                <div style={{ fontSize: 11, color: '#666' }}>Feriefridage</div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Legend */}
      <div
        style={{
          display: 'flex',
          gap: 24,
          marginBottom: 16,
          padding: '8px 12px',
          backgroundColor: '#fafafa',
          borderRadius: 4,
          fontSize: 12,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <div style={{ width: 12, height: 12, backgroundColor: '#004e6b', borderRadius: 2 }} />
          <span>Brugt (feriedage)</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <div style={{ width: 12, height: 12, backgroundColor: '#f5a623', borderRadius: 2 }} />
          <span>Afventer godkendelse</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <div style={{ width: 12, height: 12, backgroundColor: '#5a6e3a', borderRadius: 2 }} />
          <span>Brugt (feriefridage)</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <div style={{ width: 12, height: 12, backgroundColor: '#e0e0e0', borderRadius: 2 }} />
          <span>Til rådighed</span>
        </div>
      </div>

      {/* Employee Table */}
      <h3 style={{ margin: '0 0 12px 0', fontSize: 14, fontWeight: 600, color: '#004e6b' }}>
        <Icon iconName="People" style={{ marginRight: 8 }} />
        Medarbejderoversigt ({employeeBalances.length})
      </h3>

      {employeeBalances.length === 0 ? (
        <div
          style={{
            padding: 40,
            textAlign: 'center',
            backgroundColor: '#fafafa',
            borderRadius: 8,
            color: '#666',
          }}
        >
          <Icon iconName="Info" style={{ fontSize: 32, marginBottom: 12, display: 'block' }} />
          <span>Ingen medarbejderdata fundet for dette ferieår.</span>
        </div>
      ) : (
        <DetailsList
          items={employeeBalances}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          compact={true}
        />
      )}
    </div>
  );
};

export default AdminReport;
