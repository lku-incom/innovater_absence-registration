# Monthly Holiday Accrual Flow - Power Automate

This Power Automate flow automatically adds monthly holiday accrual for all employees according to Danish holiday law (Ferieloven).

## What It Does

The flow runs **automatically on the 1st of every month** at 06:00 CET and:

1. **Every month**: Adds **2.08 feriedage** (vacation days) for each employee
   - Based on: 25 days/year ÷ 12 months = 2.08 days/month

2. **Every January**: Additionally adds **5 feriefridage** (extra vacation days)
   - These are contract-based extra days (not covered by Ferieloven)
   - Given once per calendar year

## Holiday Year Calculation

The Danish holiday year runs from **September 1st to August 31st**:
- September - December → current year-next year (e.g., "2025-2026")
- January - August → previous year-current year (e.g., "2024-2025")

## Dataverse Table

Records are created in the `cr_accrualhistories` table with:

| Field | Value |
|-------|-------|
| cr_name | "{Employee Name} - Månedlig optjening - {Month Year}" |
| cr_employeeemail | Employee's email from Azure AD |
| cr_employeename | Employee's display name |
| cr_holidayyear | Calculated holiday year (e.g., "2024-2025") |
| cr_accrualdate | Current date |
| cr_accrualmonth | Current month (1-12) |
| cr_accrualyear | Current year |
| cr_daysaccrued | 2.08 |
| cr_feriefridageaccrued | 5 (January only) or 0 |
| cr_accrualtype | 100000000 (Månedlig optjening) |
| cr_notes | Description of the accrual |

## Setup Instructions

### 1. Import the Flow

1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **My flows** → **Import** → **Import Package (Legacy)**
3. Upload `MonthlyHolidayAccrualFlow.json`
4. Configure the connections:
   - **Dataverse** (shared_commondataserviceforapps)
   - **Office 365 Users** (shared_office365users)
   - **Office 365 Outlook** (shared_office365)

### 2. Configure Connections

Create or select existing connections for:

| Connector | Purpose |
|-----------|---------|
| Microsoft Dataverse | Write accrual records |
| Office 365 Users | Get list of employees from Azure AD |
| Office 365 Outlook | Send summary email to HR |

### 3. Customize Settings

**Update the recipient email** in "Send summary email" action:
- Default: `hr@innovater.dk`
- Change to your HR team's email address

**Filter employees** (optional):
- Edit the "Get all users from Azure AD" action
- Add filters like department, company, or custom attributes
- Example: `department eq 'Denmark' and accountEnabled eq true`

### 4. Test the Flow

1. Click **Test** → **Manually**
2. The flow will run immediately (not wait for schedule)
3. Verify records are created in Dataverse
4. Check the summary email

## Accrual Type Values

| Type | Value | Description |
|------|-------|-------------|
| Månedlig optjening | 100000000 | Monthly accrual (used by this flow) |
| Årsstart overførsel | 100000001 | Year start carryover |
| Manuel justering | 100000002 | Manual adjustment |
| Startsaldo | 100000003 | Initial balance |

## Troubleshooting

### No records created
- Check Azure AD filter - users must have `accountEnabled eq true` and valid email
- External/guest users (containing `#EXT#`) are automatically skipped

### Wrong holiday year
- Holiday year is calculated based on run date
- Sept-Dec = current-next year
- Jan-Aug = previous-current year

### Missing feriefridage
- Feriefridage are only added in January
- If you need to add them retroactively, use the Admin Panel in the web part

## Manual Run

To run the flow outside the schedule:
1. Open the flow in Power Automate
2. Click **Run** → **Run flow**
3. Or use the **Test** function

## Disabling the Flow

To temporarily disable:
1. Open the flow
2. Click **Turn off** in the command bar

## Related Files

- Flow definition: `power-automate/MonthlyHolidayAccrualFlow.json`
- Data model: `src/webparts/absenceRegistration/models/IHolidayBalance.ts`
- Service: `src/webparts/absenceRegistration/services/DataverseService.ts`
