import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { TextField } from '@fluentui/react/lib/TextField';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ComboBox, IComboBoxOption, IComboBox } from '@fluentui/react/lib/ComboBox';
import styles from './AbsenceRegistration.module.scss';
import { IAbsenceFormProps } from './IAbsenceRegistrationProps';
import {
  IAbsenceRegistration,
  AbsenceType,
  AbsenceTypeOptions,
  IUserInfo,
} from '../models/IAbsenceRegistration';
import { DanishHolidayService } from '../services/DanishHolidayService';
import { GraphService } from '../services/GraphService';

// Danish day and month names for date picker
const DayPickerStrings = {
  months: [
    'Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni',
    'Juli', 'August', 'September', 'Oktober', 'November', 'December',
  ],
  shortMonths: [
    'Jan', 'Feb', 'Mar', 'Apr', 'Maj', 'Jun',
    'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dec',
  ],
  days: ['Søndag', 'Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag'],
  shortDays: ['Sø', 'Ma', 'Ti', 'On', 'To', 'Fr', 'Lø'],
  goToToday: 'Gå til i dag',
  prevMonthAriaLabel: 'Forrige måned',
  nextMonthAriaLabel: 'Næste måned',
  prevYearAriaLabel: 'Forrige år',
  nextYearAriaLabel: 'Næste år',
  closeButtonAriaLabel: 'Luk',
  monthPickerHeaderAriaLabel: '{0}, vælg for at ændre år',
  yearPickerHeaderAriaLabel: '{0}, vælg for at ændre måned',
};

const AbsenceForm: React.FC<IAbsenceFormProps> = (props) => {
  const {
    context,
    currentUser,
    onSave,
    onSubmit,
    onCancel,
    editingRegistration,
    isLoading,
    readOnly = false,
  } = props;

  const [startDate, setStartDate] = useState<Date | undefined>(
    editingRegistration?.StartDate
  );
  const [endDate, setEndDate] = useState<Date | undefined>(
    editingRegistration?.EndDate
  );
  const [absenceType, setAbsenceType] = useState<AbsenceType | undefined>(
    editingRegistration?.AbsenceType
  );
  const [notes, setNotes] = useState<string>(editingRegistration?.Notes || '');
  const [numberOfDays, setNumberOfDays] = useState<number>(
    editingRegistration?.NumberOfDays || 0
  );
  const [validationError, setValidationError] = useState<string>('');
  const [selectedApprover, setSelectedApprover] = useState<IUserInfo | undefined>(undefined);
  const [approverOptions, setApproverOptions] = useState<IComboBoxOption[]>([]);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const searchTimeoutRef = useRef<number | undefined>(undefined);

  // Get user initials for avatar
  const getInitials = (name: string): string => {
    if (!name) return '?';
    const parts = name.split(' ');
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
  };

  // Check if manager info is missing
  const hasMissingInfo = !currentUser?.manager;

  // Reset form when editingRegistration changes
  useEffect(() => {
    if (editingRegistration) {
      setStartDate(editingRegistration.StartDate);
      setEndDate(editingRegistration.EndDate);
      setAbsenceType(editingRegistration.AbsenceType);
      setNotes(editingRegistration.Notes || '');
      setNumberOfDays(editingRegistration.NumberOfDays);
      // Set approver from existing registration if no manager
      if (editingRegistration.ApproverEmail && !currentUser?.manager) {
        setSelectedApprover({
          id: 0,
          email: editingRegistration.ApproverEmail,
          displayName: editingRegistration.ApproverName || '',
        });
        setApproverOptions([{
          key: editingRegistration.ApproverEmail,
          text: `${editingRegistration.ApproverName} (${editingRegistration.ApproverEmail})`,
        }]);
      }
    } else {
      setStartDate(undefined);
      setEndDate(undefined);
      setAbsenceType(undefined);
      setNotes('');
      setNumberOfDays(0);
      setSelectedApprover(undefined);
      setApproverOptions([]);
    }
    setValidationError('');
  }, [editingRegistration, currentUser?.manager]);

  // Calculate working days when dates change
  useEffect(() => {
    if (startDate && endDate) {
      const days = DanishHolidayService.calculateWorkingDays(startDate, endDate);
      setNumberOfDays(days);
    } else {
      setNumberOfDays(0);
    }
  }, [startDate, endDate]);

  const handleStartDateChange = useCallback((date: Date | null | undefined): void => {
    setStartDate(date || undefined);
    setValidationError('');
    if (date && endDate && date > endDate) {
      setEndDate(undefined);
    }
  }, [endDate]);

  const handleEndDateChange = useCallback((date: Date | null | undefined): void => {
    setEndDate(date || undefined);
    setValidationError('');
  }, []);

  const handleAbsenceTypeChange = useCallback(
    (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
      if (option) {
        setAbsenceType(option.key as AbsenceType);
        setValidationError('');
      }
    },
    []
  );

  const handleNotesChange = useCallback(
    (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
      setNotes(newValue || '');
    },
    []
  );

  // Search users for approver selection
  const handleApproverSearch = useCallback(
    (text: string): void => {
      // Clear any existing timeout
      if (searchTimeoutRef.current) {
        window.clearTimeout(searchTimeoutRef.current);
      }

      if (!text || text.length < 2) {
        setApproverOptions([]);
        return;
      }

      setIsSearching(true);

      // Debounce the search
      searchTimeoutRef.current = window.setTimeout(async () => {
        try {
          const graphService = GraphService.getInstance();
          const users = await graphService.searchUsers(text);

          const options: IComboBoxOption[] = users.map((user) => ({
            key: user.email,
            text: `${user.displayName} (${user.email})`,
            data: user,
          }));

          setApproverOptions(options);
        } catch (error) {
          console.error('Error searching users:', error);
          setApproverOptions([]);
        } finally {
          setIsSearching(false);
        }
      }, 300);
    },
    []
  );

  const handleApproverChange = useCallback(
    (_event: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
      if (option) {
        const user = option.data as IUserInfo;
        setSelectedApprover(user);
        setValidationError('');
      }
    },
    []
  );

  const validateForm = (): boolean => {
    if (!startDate) {
      setValidationError('Vælg venligst en startdato');
      return false;
    }
    if (!endDate) {
      setValidationError('Vælg venligst en slutdato');
      return false;
    }
    if (startDate > endDate) {
      setValidationError('Slutdato skal være efter startdato');
      return false;
    }
    if (!absenceType) {
      setValidationError('Vælg venligst en fraværstype');
      return false;
    }
    if (numberOfDays === 0) {
      setValidationError('Der er ingen arbejdsdage i den valgte periode');
      return false;
    }
    // Validate approver if no manager from Azure AD
    if (!currentUser?.manager) {
      if (!selectedApprover?.email) {
        setValidationError('Vælg venligst en godkender');
        return false;
      }
    }
    return true;
  };

  const buildRegistration = (): IAbsenceRegistration => {
    // Use manager from Azure AD if available, otherwise use selected approver
    const approverName = currentUser?.manager?.displayName || selectedApprover?.displayName || '';
    const approverEmail = currentUser?.manager?.email || selectedApprover?.email || '';

    return {
      Id: editingRegistration?.Id,
      DataverseId: editingRegistration?.DataverseId,
      Title: `${currentUser?.displayName} - ${absenceType}`,
      EmployeeId: currentUser?.id || 0,
      EmployeeEmail: currentUser?.email || '',
      EmployeeName: currentUser?.displayName || '',
      Department: currentUser?.department || '',
      ApproverId: 0,
      ApproverName: approverName,
      ApproverEmail: approverEmail,
      StartDate: startDate!,
      EndDate: endDate!,
      NumberOfDays: numberOfDays,
      AbsenceType: absenceType!,
      Notes: notes,
      Status: editingRegistration?.Status || 'Kladde',
    };
  };

  const handleSaveClick = async (): Promise<void> => {
    if (!validateForm()) return;
    await onSave(buildRegistration());
  };

  const handleSubmitClick = async (): Promise<void> => {
    if (!validateForm()) return;
    await onSubmit(buildRegistration());
  };

  const absenceTypeOptions: IChoiceGroupOption[] = AbsenceTypeOptions.map((option) => ({
    key: option.key,
    text: option.text,
  }));

  const minEndDate = startDate || new Date();

  return (
    <div className={styles.formSection}>
      {/* Buttons at top */}
      <div className={styles.buttonRow} style={{ marginBottom: '24px' }}>
        {!readOnly && (
          <>
            <DefaultButton
              text="Gem som kladde"
              onClick={handleSaveClick}
              disabled={isLoading}
              iconProps={{ iconName: 'Save' }}
            />
            <DefaultButton
              text="Send til godkendelse"
              onClick={handleSubmitClick}
              disabled={isLoading}
              primary
              iconProps={{ iconName: 'Send' }}
            />
          </>
        )}
        <DefaultButton
          text="Luk"
          onClick={onCancel}
          disabled={isLoading}
          iconProps={{ iconName: 'Cancel' }}
        />
        {isLoading && <Spinner size={SpinnerSize.small} />}
      </div>

      {validationError && (
        <div className={styles.errorBanner}>{validationError}</div>
      )}

      {/* Employee with avatar */}
      <div className={styles.formRow}>
        <div className={styles.formField}>
          <label>
            Medarbejder <span className={styles.required}>*</span>
          </label>
          <div className={styles.employeeDisplay}>
            <div className={styles.avatar}>
              <span className={styles.initials}>
                {getInitials(currentUser?.displayName || '')}
              </span>
            </div>
            <span className={styles.name}>
              {currentUser?.displayName || 'Indlæser...'}
            </span>
          </div>
        </div>
      </div>

      {/* Approver/Manager */}
      <div className={styles.formRow}>
        <div className={styles.formField} style={{ flex: 1 }}>
          <label>Godkender {!currentUser?.manager && <span className={styles.required}>*</span>}</label>
          {currentUser?.manager ? (
            <div className={styles.fieldValue}>
              {currentUser.manager.displayName}
            </div>
          ) : (
            <ComboBox
              placeholder="Søg efter godkender..."
              options={approverOptions}
              selectedKey={selectedApprover?.email}
              onChange={handleApproverChange}
              onInputValueChange={handleApproverSearch}
              allowFreeform
              autoComplete="off"
              disabled={isLoading || readOnly}
              useComboBoxAsMenuWidth
              calloutProps={{ doNotLayer: true }}
              text={selectedApprover ? `${selectedApprover.displayName} (${selectedApprover.email})` : undefined}
            />
          )}
          {isSearching && <Spinner size={SpinnerSize.xSmall} style={{ marginTop: 4 }} />}
        </div>
      </div>

      {/* Date selection */}
      <div className={styles.formRow}>
        <div className={styles.formField}>
          <label>
            Dato start <span className={styles.required}>*</span>
          </label>
          <DatePicker
            firstDayOfWeek={DayOfWeek.Monday}
            strings={DayPickerStrings}
            placeholder="DD-MM-YYYY"
            value={startDate}
            onSelectDate={handleStartDateChange}
            formatDate={(date) => date?.toLocaleDateString('da-DK', {
              day: '2-digit',
              month: '2-digit',
              year: 'numeric'
            }) || ''}
            minDate={new Date()}
            disabled={isLoading || readOnly}
          />
        </div>

        <div className={styles.formField}>
          <label>
            Dato slut <span className={styles.required}>*</span>
          </label>
          <DatePicker
            firstDayOfWeek={DayOfWeek.Monday}
            strings={DayPickerStrings}
            placeholder="DD-MM-YYYY"
            value={endDate}
            onSelectDate={handleEndDateChange}
            formatDate={(date) => date?.toLocaleDateString('da-DK', {
              day: '2-digit',
              month: '2-digit',
              year: 'numeric'
            }) || ''}
            minDate={minEndDate}
            disabled={isLoading || readOnly}
          />
        </div>
      </div>

      {/* Number of days */}
      <div className={styles.formRow}>
        <div className={styles.formField}>
          <label>
            Antal dage{' '}
            <TooltipHost content="Antal arbejdsdage ekskl. weekender og danske helligdage">
              <Icon iconName="Info" className={styles.infoIcon} />
            </TooltipHost>
          </label>
          <div className={styles.daysDisplay}>
            <span className={styles.daysValue}>{numberOfDays}</span>
          </div>
        </div>
      </div>

      {/* Absence type */}
      <div className={`${styles.formRow} ${styles.fullWidth}`}>
        <div className={styles.formField}>
          <label>
            Fraværstype <span className={styles.required}>*</span>
          </label>
          <ChoiceGroup
            className={styles.absenceTypeGroup}
            options={absenceTypeOptions}
            selectedKey={absenceType}
            onChange={handleAbsenceTypeChange}
            disabled={isLoading || readOnly}
          />
        </div>
      </div>

      {/* Notes */}
      <div className={`${styles.formRow} ${styles.fullWidth}`}>
        <div className={styles.formField}>
          <label>Noter</label>
          <TextField
            className={styles.notesField}
            multiline
            rows={4}
            value={notes}
            onChange={handleNotesChange}
            disabled={isLoading || readOnly}
            readOnly={readOnly}
          />
        </div>
      </div>

      {/* Info message */}
      <div className={styles.infoMessage}>
        En ferieregistrering skal godkendes internt ved at vælge "Send til godkendelse" under Handlinger.
        Du modtager automatisk en e-mail bekræftelse ved godkendelse og vil kunne se din opdaterede
        feriebalance. Kun godkendte registreringer medtages i ferie- og feriefridage balancen.
      </div>
    </div>
  );
};

export default AbsenceForm;
