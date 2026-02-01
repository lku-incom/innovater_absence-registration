# Power Automate Flows for Holiday Balance Automation

This document describes the Power Automate flows needed to automatically maintain the holiday balance tables in Dataverse according to Danish Holiday Law (Ferieloven).

## Dataverse Table Names

| Table (Logical Name) | API Entity Set Name | Description |
|---------------------|---------------------|-------------|
| `cr153_AbsenceRegistration` | `cr153_absenceregistrations` | Absence registrations |
| `cr_holidaybalance` | `cr_holidaybalances` | Holiday balances |
| `cr_accrualhistory` | `cr_accrualhistories` | Accrual audit history |

**Note:** Power Automate uses the logical names. Web API calls use the entity set names (plural).

## Overview

| Flow Name | Trigger | Schedule | Purpose |
|-----------|---------|----------|---------|
| Monthly Accrual | Scheduled | 1st of month, 06:00 | Add 2.08 feriedage + 0.42 feriefridage |
| Absence Approval Sync | Dataverse trigger | On approval | Update UsedDays when absence is approved |
| New Holiday Year Setup | Scheduled | September 1, 00:01 | Create new year balances, process transfers |
| Year-End Cleanup | Scheduled | January 1, 00:01 | Archive old records, forfeit unused mandatory days |
| Transfer Deadline Reminder | Scheduled | December 1, 09:00 | Notify employees about transfer deadline |

---

## Flow 1: Monthly Accrual Flow

**Name:** `HolidayBalance-MonthlyAccrual`
**Trigger:** Recurrence - Monthly on day 1 at 06:00 CET
**Purpose:** Accrue holiday days for all active employees

### Flow Steps

```
TRIGGER: Recurrence
├── Frequency: Month
├── Interval: 1
├── Start time: 2025-02-01T06:00:00Z
├── Time zone: (UTC+01:00) Copenhagen
└── On these days: 1

ACTION 1: Initialize Variables
├── varCurrentDate (DateTime) = utcNow()
├── varHolidayYear (String) = [calculated - see expression below]
├── varMonthlyAccrual (Decimal) = 2.08
└── varMonthlyFeriefridage (Decimal) = 0.42

ACTION 2: Calculate Holiday Year
├── Expression: if(greater(int(formatDateTime(utcNow(),'MM')), 8),
│               concat(formatDateTime(utcNow(),'yyyy'), '-', string(add(int(formatDateTime(utcNow(),'yyyy')), 1))),
│               concat(string(sub(int(formatDateTime(utcNow(),'yyyy')), 1)), '-', formatDateTime(utcNow(),'yyyy')))
└── Store in: varHolidayYear

ACTION 3: List Holiday Balance Records
├── Table: cr_holidaybalances
├── Filter rows: cr_holidayyear eq '@{varHolidayYear}' and cr_isactive eq true
└── Select columns: cr_holidaybalanceid,cr_employeeemail,cr_employeename,cr_accrueddays,cr_feriefridageaccrued,cr_availabledays,cr_feriefridageavailable,cr_lastaccrualdate

ACTION 4: Apply to Each (Balance Record)
│
├── CONDITION: Check if already accrued this month
│   ├── If: formatDateTime(items('Apply_to_each')?['cr_lastaccrualdate'],'yyyy-MM') equals formatDateTime(utcNow(),'yyyy-MM')
│   ├── Yes: Skip (do nothing)
│   └── No: Continue to update
│
├── ACTION 4.1: Calculate New Values
│   ├── newAccruedDays = add(items('Apply_to_each')?['cr_accrueddays'], varMonthlyAccrual)
│   ├── newFeriefridageAccrued = add(items('Apply_to_each')?['cr_feriefridageaccrued'], varMonthlyFeriefridage)
│   ├── newAvailableDays = [recalculate based on formula]
│   └── newFeriefridageAvailable = [recalculate based on formula]
│
├── ACTION 4.2: Update Holiday Balance Record
│   ├── Table: cr_holidaybalances
│   ├── Row ID: items('Apply_to_each')?['cr_holidaybalanceid']
│   └── Update fields:
│       ├── cr_accrueddays: @{newAccruedDays}
│       ├── cr_feriefridageaccrued: @{newFeriefridageAccrued}
│       ├── cr_availabledays: @{newAvailableDays}
│       ├── cr_feriefridageavailable: @{newFeriefridageAvailable}
│       └── cr_lastaccrualdate: @{utcNow()}
│
└── ACTION 4.3: Create Accrual History Record (Optional)
    ├── Table: cr_accrualhistories
    └── Fields:
        ├── cr_name: @{items('Apply_to_each')?['cr_employeename']} - @{formatDateTime(utcNow(),'MMMM yyyy')}
        ├── cr_employeeemail: @{items('Apply_to_each')?['cr_employeeemail']}
        ├── cr_employeename: @{items('Apply_to_each')?['cr_employeename']}
        ├── cr_holidayyear: @{varHolidayYear}
        ├── cr_accrualdate: @{utcNow()}
        ├── cr_accrualmonth: @{int(formatDateTime(utcNow(),'MM'))}
        ├── cr_accrualyear: @{int(formatDateTime(utcNow(),'yyyy'))}
        ├── cr_daysaccrued: @{varMonthlyAccrual}
        ├── cr_feriefridageaccrued: @{varMonthlyFeriefridage}
        └── cr_accrualtype: 100000000 (Månedlig optjening)

ACTION 5: Send Summary Email (Optional)
├── To: HR Administrator
├── Subject: Holiday Balance Monthly Accrual Completed
└── Body: Processed X records for @{varHolidayYear}
```

### Key Expressions

**Calculate Holiday Year:**
```
if(
  greater(int(formatDateTime(utcNow(),'MM')), 8),
  concat(formatDateTime(utcNow(),'yyyy'), '-', string(add(int(formatDateTime(utcNow(),'yyyy')), 1))),
  concat(string(sub(int(formatDateTime(utcNow(),'yyyy')), 1)), '-', formatDateTime(utcNow(),'yyyy'))
)
```

**Calculate Available Days:**
```
add(
  add(
    outputs('newAccruedDays'),
    coalesce(items('Apply_to_each')?['cr_transferredindays'], 0)
  ),
  sub(
    sub(0, coalesce(items('Apply_to_each')?['cr_useddays'], 0)),
    coalesce(items('Apply_to_each')?['cr_pendingdays'], 0)
  )
)
```

---

## Flow 2: Absence Approval Sync

**Name:** `HolidayBalance-AbsenceApprovalSync`
**Trigger:** When a row is modified (cr153_AbsenceRegistration)
**Purpose:** Update UsedDays when an absence is approved

### Absence Type Option Set Values
| Value | Type |
|-------|------|
| 100000000 | Ferie |
| 100000001 | Sygdom |
| 100000002 | Barselsorlov |
| 100000003 | Feriefridage |
| 100000004 | Flex/afspadsering |
| 100000005 | Andet fravær |

### Status Option Set Values
| Value | Status |
|-------|--------|
| 100000000 | Kladde |
| 100000001 | Afventer godkendelse |
| 100000002 | Godkendt |
| 100000003 | Afvist |

### Flow Steps

```
TRIGGER: When a row is modified
├── Table: cr153_AbsenceRegistration (logical name)
├── Scope: Organization
└── Filter: cr153_status eq 100000002 (Godkendt)

CONDITION 1: Check if status changed to Godkendt
├── Expression: triggerBody()?['cr153_status'] equals 100000002
└── AND: triggerOutputs()?['body/_cr153_status_previousvalue'] not equals 100000002

If Yes:
│
├── ACTION 1: Get Absence Details
│   ├── NumberOfDays: triggerBody()?['cr153_numberofdays']
│   ├── AbsenceType: triggerBody()?['cr153_absencetype']
│   ├── EmployeeEmail: triggerBody()?['cr153_employeeemail']
│   └── StartDate: triggerBody()?['cr153_startdate']
│
├── ACTION 2: Calculate Holiday Year from Start Date
│   └── Expression: [same as above based on cr153_startdate]
│
├── ACTION 3: Get Holiday Balance Record
│   ├── Table: cr_holidaybalance (logical name)
│   └── Filter: cr_employeeemail eq '@{triggerBody()?['cr153_employeeemail']}' and cr_holidayyear eq '@{varHolidayYear}'
│
├── CONDITION 2: Check Absence Type (only Ferie and Feriefridage affect balance)
│   │
│   ├── If AbsenceType equals 100000000 (Ferie):
│   │   ├── ACTION 3.1: Update Feriedage
│   │   │   ├── newUsedDays = add(first(outputs('List_records')?['body/value'])?['cr_useddays'], triggerBody()?['cr153_numberofdays'])
│   │   │   ├── newPendingDays = sub(first(outputs('List_records')?['body/value'])?['cr_pendingdays'], triggerBody()?['cr153_numberofdays'])
│   │   │   └── newAvailableDays = [recalculate]
│   │   └── ACTION 3.2: Update Holiday Balance
│   │       └── Update cr_useddays, cr_pendingdays, cr_availabledays
│   │
│   └── If AbsenceType equals 100000003 (Feriefridage):
│       ├── ACTION 3.3: Update Feriefridage
│       │   ├── newFeriefridageUsed = add(first(outputs('List_records')?['body/value'])?['cr_feriefridageused'], triggerBody()?['cr153_numberofdays'])
│       │   └── newFeriefridageAvailable = [recalculate]
│       └── ACTION 3.4: Update Holiday Balance
│           └── Update cr_feriefridageused, cr_feriefridageavailable
│
│   Note: Sygdom (100000001), Barselsorlov (100000002), Flex (100000004),
│         Andet fravær (100000005) do NOT affect holiday balance
│
└── ACTION 4: Add Note to Absence Record (Optional)
    └── cr153_notes: append "Balance updated on @{utcNow()}"
```

### Additional Trigger: Pending Days Update

When an absence is submitted for approval (status = 100000001), increase PendingDays:

```
TRIGGER: When a row is modified
├── Table: cr153_AbsenceRegistration
├── Filter: cr153_status eq 100000001 (Afventer godkendelse)

CONDITION: Only for Ferie (100000000) or Feriefridage (100000003)

ACTION: Update PendingDays
├── Get current balance from cr_holidaybalance
├── If Ferie: newPendingDays = add(currentPendingDays, numberOfDays)
├── Update cr_pendingdays
└── Recalculate cr_availabledays
```

---

## Flow 3: New Holiday Year Setup

**Name:** `HolidayBalance-NewYearSetup`
**Trigger:** Recurrence - Yearly on September 1 at 00:01 CET
**Purpose:** Create new holiday year balances and process transfers

### Flow Steps

```
TRIGGER: Recurrence
├── Frequency: Year
├── Interval: 1
├── Start time: 2025-09-01T00:01:00Z
└── Time zone: (UTC+01:00) Copenhagen

ACTION 1: Initialize Variables
├── varPreviousYear = concat(string(sub(int(formatDateTime(utcNow(),'yyyy')), 1)), '-', formatDateTime(utcNow(),'yyyy'))
├── varNewYear = concat(formatDateTime(utcNow(),'yyyy'), '-', string(add(int(formatDateTime(utcNow(),'yyyy')), 1)))
└── varEmployeesList = []

ACTION 2: Get All Active Employees from Previous Year
├── Table: cr_holidaybalances
├── Filter: cr_holidayyear eq '@{varPreviousYear}' and cr_isactive eq true
└── Select: all fields

ACTION 3: Apply to Each (Employee Balance)
│
├── ACTION 3.1: Check for Existing New Year Record
│   ├── Table: cr_holidaybalances
│   └── Filter: cr_employeeemail eq '@{items('Apply_to_each')?['cr_employeeemail']}' and cr_holidayyear eq '@{varNewYear}'
│
├── CONDITION: New Year Record Exists?
│   │
│   ├── If No: Create New Balance Record
│   │   │
│   │   ├── ACTION 3.2: Calculate Transfer Amount
│   │   │   ├── If HasTransferAgreement = true:
│   │   │   │   └── transferIn = items('Apply_to_each')?['cr_transferredoutdays']
│   │   │   └── Else:
│   │   │       └── transferIn = 0
│   │   │
│   │   └── ACTION 3.3: Create New Holiday Balance
│   │       ├── Table: cr_holidaybalances
│   │       └── Fields:
│   │           ├── cr_name: @{items('Apply_to_each')?['cr_employeename']} - @{varNewYear}
│   │           ├── cr_employeeemail: @{items('Apply_to_each')?['cr_employeeemail']}
│   │           ├── cr_employeename: @{items('Apply_to_each')?['cr_employeename']}
│   │           ├── cr_holidayyear: @{varNewYear}
│   │           ├── cr_accrueddays: 0
│   │           ├── cr_useddays: 0
│   │           ├── cr_pendingdays: 0
│   │           ├── cr_availabledays: @{transferIn}
│   │           ├── cr_transferredindays: @{transferIn}
│   │           ├── cr_transferredoutdays: 0
│   │           ├── cr_hastransferagreement: false
│   │           ├── cr_feriefridageaccrued: 0
│   │           ├── cr_feriefridageused: 0
│   │           ├── cr_feriefridageavailable: @{feriefridageTransferIn}
│   │           ├── cr_feriefridagetransferredin: @{feriefridageTransferIn}
│   │           ├── cr_feriefridagetransferredout: 0
│   │           └── cr_isactive: true
│   │
│   └── If Yes: Update Existing (merge transfer if needed)
│       └── ACTION 3.4: Update Transfer In Days
│
├── ACTION 3.5: Create Transfer Accrual History (if transfer > 0)
│   ├── Table: cr_accrualhistories
│   └── Fields:
│       ├── cr_accrualtype: 100000001 (Årsstart overførsel)
│       ├── cr_daysaccrued: @{transferIn}
│       └── cr_notes: Overført fra @{varPreviousYear}
│
└── ACTION 3.6: Mark Previous Year as Inactive
    ├── Table: cr_holidaybalances
    ├── Row ID: items('Apply_to_each')?['cr_holidaybalanceid']
    └── Update: cr_isactive = false

ACTION 4: Send Summary Report
├── To: HR Administrator
├── Subject: New Holiday Year @{varNewYear} - Setup Complete
└── Body: Created X new balance records, processed Y transfers
```

---

## Flow 4: Year-End Cleanup

**Name:** `HolidayBalance-YearEndCleanup`
**Trigger:** Recurrence - Yearly on January 1 at 00:01 CET
**Purpose:** Process forfeitures and archive completed years

### Flow Steps

```
TRIGGER: Recurrence
├── Frequency: Year
├── Interval: 1
├── Start time: 2026-01-01T00:01:00Z
└── Time zone: (UTC+01:00) Copenhagen

ACTION 1: Initialize Variables
├── varCompletedYear = concat(string(sub(int(formatDateTime(utcNow(),'yyyy')), 2)), '-', string(sub(int(formatDateTime(utcNow(),'yyyy')), 1)))
│   // For Jan 1, 2026: This gives "2024-2025" (the year that ended Dec 31, 2025)
└── varReportData = []

ACTION 2: Get All Balances for Completed Year
├── Table: cr_holidaybalances
└── Filter: cr_holidayyear eq '@{varCompletedYear}'

ACTION 3: Apply to Each (Balance)
│
├── ACTION 3.1: Calculate Forfeited Days
│   ├── MANDATORY_DAYS = 20
│   ├── usedDays = items('Apply_to_each')?['cr_useddays']
│   ├── mandatoryRemaining = max(0, sub(MANDATORY_DAYS, usedDays))
│   ├── totalAvailable = [calculate available]
│   └── forfeitedDays = min(mandatoryRemaining, totalAvailable)
│
├── CONDITION: Any days forfeited?
│   │
│   └── If forfeitedDays > 0:
│       │
│       ├── ACTION 3.2: Log Forfeiture
│       │   ├── Table: cr_accrualhistories
│       │   └── Fields:
│       │       ├── cr_accrualtype: 100000004 (Årsslut forfeiture)
│       │       ├── cr_daysaccrued: @{negate(forfeitedDays)}
│       │       └── cr_notes: Forbrugt ikke inden deadline - bortfalder jf. Ferieloven
│       │
│       └── ACTION 3.3: Send Notification to Employee
│           ├── To: items('Apply_to_each')?['cr_employeeemail']
│           ├── Subject: Feriedage bortfaldet
│           └── Body: @{forfeitedDays} feriedage er bortfaldet...
│
├── ACTION 3.4: Calculate 5th Week Payout (if no transfer agreement)
│   ├── If HasTransferAgreement = false AND daysAbove20 > 0:
│   │   └── Log payout record
│   └── Else: Skip
│
└── ACTION 3.5: Add to Report
    └── Append employee summary to varReportData

ACTION 4: Generate Year-End Report
├── Create Excel/PDF report
└── Send to HR Administrator
```

---

## Flow 5: Transfer Deadline Reminder

**Name:** `HolidayBalance-TransferReminder`
**Trigger:** Recurrence - Yearly on December 1 at 09:00 CET
**Purpose:** Remind employees about transfer agreement deadline

### Flow Steps

```
TRIGGER: Recurrence
├── Frequency: Year
├── Start time: 2025-12-01T09:00:00Z
└── Time zone: (UTC+01:00) Copenhagen

ACTION 1: Calculate Current Holiday Year
└── varHolidayYear = [expression]

ACTION 2: Get Balances Without Transfer Agreement
├── Table: cr_holidaybalances
├── Filter: cr_holidayyear eq '@{varHolidayYear}'
│           and cr_isactive eq true
│           and cr_hastransferagreement eq false
└── Select: cr_employeeemail, cr_employeename, cr_accrueddays, cr_useddays, cr_pendingdays

ACTION 3: Apply to Each (Balance)
│
├── ACTION 3.1: Calculate Transferable Days
│   ├── totalAvailable = accruedDays + transferredIn - usedDays - pendingDays
│   ├── daysAbove20 = max(0, totalAvailable - 20)
│   └── maxTransferable = min(daysAbove20, 5)
│
├── CONDITION: Has Transferable Days?
│   │
│   └── If maxTransferable > 0:
│       │
│       └── ACTION 3.2: Send Reminder Email
│           ├── To: items('Apply_to_each')?['cr_employeeemail']
│           ├── Subject: Husk: Frist for ferieoverførsel er 31. december
│           └── Body:
│               Kære @{items('Apply_to_each')?['cr_employeename']},
│
│               Du har mulighed for at overføre op til @{maxTransferable} feriedage
│               til næste ferieår.
│
│               Ifølge Ferieloven skal en skriftlig aftale om overførsel af den
│               5. ferieuge indgås senest den 31. december.
│
│               Din nuværende feriesaldo:
│               - Optjent: @{accruedDays} dage
│               - Brugt: @{usedDays} dage
│               - Tilgængelig: @{totalAvailable} dage
│               - Kan overføres: @{maxTransferable} dage
│
│               Kontakt venligst HR for at indgå en overførselsaftale.
│
│               Med venlig hilsen,
│               HR-afdelingen

ACTION 4: Send Summary to HR
├── To: HR Administrator
├── Subject: Ferieoverførsel - Påmindelser sendt
└── Body: Sendt @{length(outputs('Apply_to_each'))} påmindelser
```

---

## Dataverse Tables Required

### cr_holidaybalances (Holiday Balance)

| Field | Type | Description |
|-------|------|-------------|
| cr_holidaybalanceid | GUID | Primary key |
| cr_name | String | Auto: "Name - Year" |
| cr_employeeemail | String | Employee email (key) |
| cr_employeename | String | Employee display name |
| cr_holidayyear | String | e.g., "2025-2026" |
| cr_accrueddays | Decimal | Total accrued |
| cr_useddays | Decimal | Approved absences |
| cr_pendingdays | Decimal | Pending approval |
| cr_availabledays | Decimal | Calculated |
| cr_transferredindays | Decimal | From previous year |
| cr_transferredoutdays | Decimal | To next year |
| cr_hastransferagreement | Boolean | Written agreement |
| cr_transferagreementdate | DateTime | Agreement date |
| cr_feriefridageaccrued | Decimal | Extra days accrued |
| cr_feriefridageused | Decimal | Extra days used |
| cr_feriefridageavailable | Decimal | Calculated |
| cr_feriefridagetransferredin | Decimal | From previous |
| cr_feriefridagetransferredout | Decimal | To next |
| cr_lastaccrualdate | DateTime | Last accrual |
| cr_isactive | Boolean | Current year flag |

### cr_accrualhistories (Accrual History)

| Field | Type | Description |
|-------|------|-------------|
| cr_accrualhistoryid | GUID | Primary key |
| cr_name | String | Auto: "Name - Month Year" |
| cr_employeeemail | String | Employee email |
| cr_employeename | String | Employee name |
| cr_holidayyear | String | Holiday year |
| cr_accrualdate | DateTime | When accrued |
| cr_accrualmonth | Integer | Month (1-12) |
| cr_accrualyear | Integer | Year |
| cr_daysaccrued | Decimal | Days (can be negative) |
| cr_feriefridageaccrued | Decimal | Extra days |
| cr_balanceafteraccrual | Decimal | Running total |
| cr_accrualtype | OptionSet | Type of accrual |
| cr_notes | String | Notes/comments |

### cr_accrualtype Option Set Values

| Value | Label (Danish) | Label (English) |
|-------|----------------|-----------------|
| 100000000 | Månedlig optjening | Monthly Accrual |
| 100000001 | Årsstart overførsel | Year Start Transfer In |
| 100000002 | Manuel justering | Manual Adjustment |
| 100000003 | Startsaldo | Initial Balance |
| 100000004 | Årsslut overførsel (ud) | Year End Transfer Out |
| 100000005 | Feriefridage optjening | Feriefridage Accrual |
| 100000006 | Bortfald | Forfeiture |
| 100000007 | Udbetaling | Payout |

---

## Implementation Checklist

- [ ] Create cr_accrualhistories table in Dataverse
- [ ] Add cr_accrualtype option set
- [ ] Import Flow 1: Monthly Accrual
- [ ] Import Flow 2: Absence Approval Sync
- [ ] Import Flow 3: New Holiday Year Setup
- [ ] Import Flow 4: Year-End Cleanup
- [ ] Import Flow 5: Transfer Reminder
- [ ] Test each flow in sandbox environment
- [ ] Configure email templates
- [ ] Set up error notifications
- [ ] Document manual override procedures

---

## Error Handling

Each flow should include:

1. **Try-Catch Scope** around main operations
2. **Configure Run After** to handle failures
3. **Send Email on Failure** to administrator
4. **Log Errors** to a custom error log table

Example error handling pattern:
```
SCOPE: Try
├── [Main flow actions]
│
SCOPE: Catch (Configure run after: has failed, has timed out)
├── ACTION: Send Error Email
│   ├── To: admin@company.dk
│   ├── Subject: Flow Error: @{workflow()?['name']}
│   └── Body: Error details: @{result('Try')}
└── ACTION: Terminate with error status
```

---

## Testing Checklist

### Flow 1: Monthly Accrual
- [ ] Run manually and verify accrual amounts
- [ ] Verify duplicate prevention (same month)
- [ ] Check accrual history records created
- [ ] Verify calculations match Danish law

### Flow 2: Absence Approval
- [ ] Submit absence → verify pending days increase
- [ ] Approve absence → verify used days increase, pending decrease
- [ ] Reject absence → verify pending days decrease
- [ ] Test both Ferie and Feriefridage types

### Flow 3: New Year Setup
- [ ] Run on test date (Sept 1)
- [ ] Verify new records created for all employees
- [ ] Verify transfers processed correctly
- [ ] Verify old year marked inactive

### Flow 4: Year-End Cleanup
- [ ] Test forfeiture calculation
- [ ] Verify notifications sent
- [ ] Check report generation

### Flow 5: Transfer Reminder
- [ ] Verify only eligible employees notified
- [ ] Check email content accuracy
- [ ] Test summary to HR
