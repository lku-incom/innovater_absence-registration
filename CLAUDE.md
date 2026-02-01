# Absence Registration Project Context

## Project Overview
This is a **SharePoint Framework (SPFx) web part** for employee absence/vacation registration, built for Innovater (Danish company).

## Tech Stack
- **SPFx 1.18.0** with **React 17** and **TypeScript 4.5**
- **FluentUI React 8.110** for UI components
- **PnPjs 3.18** for SharePoint/Graph API
- **Dataverse** (Power Platform) for data storage
- **date-fns** for date manipulation

## Key Architecture

### Components (`src/webparts/absenceRegistration/components/`)
- `AbsenceRegistration.tsx` - Main container with tab navigation
- `AbsenceForm.tsx` - Registration form with date pickers and validation
- `MyRegistrations.tsx` - List view with status badges and actions
- `PendingApprovals.tsx` - Approval list for managers
- `AdminPanel.tsx` - Admin functions (delete all, view all registrations)

### Services (`src/webparts/absenceRegistration/services/`)
- `DataverseService.ts` - CRUD operations for Dataverse table `cr153_absenceregistrations`
- `GraphService.ts` - Microsoft Graph for user profiles and manager lookup
- `DanishHolidayService.ts` - Working day calculations excluding Danish holidays
- `SharePointService.ts` - SharePoint list operations (alternative storage)

### Models (`src/webparts/absenceRegistration/models/`)
- `IAbsenceRegistration.ts` - Core interfaces and types

## Dataverse Schema

**Table:** `cr153_absenceregistrations`

| Field | Type | Description |
|-------|------|-------------|
| `cr153_absenceregistrationid` | GUID | Primary key |
| `cr153_employeename` | string | Employee display name |
| `cr153_employeeemail` | string | Employee email |
| `cr153_approvername` | string | Approver/manager name |
| `cr153_approveremail` | string | Approver email |
| `cr153_startdate` | datetime | Absence start date |
| `cr153_enddate` | datetime | Absence end date |
| `cr153_numberofdays` | number | Calculated working days |
| `cr153_absencetype` | optionset | Type of absence |
| `cr153_status` | optionset | Registration status |
| `cr153_notes` | string | Optional notes |
| `cr153_approvercomments` | string | Approver's comments |
| `cr153_approvaldate` | datetime | When approved/rejected |

### Status Values
| Value | Danish | English |
|-------|--------|---------|
| 100000000 | Kladde | Draft |
| 100000001 | Afventer godkendelse | Pending Approval |
| 100000002 | Godkendt | Approved |
| 100000003 | Afvist | Rejected |

### Absence Type Values
| Value | Danish |
|-------|--------|
| 100000000 | Ferie |
| 100000001 | Sygdom |
| 100000002 | Barselsorlov |
| 100000003 | Feriefridage |
| 100000004 | Flex/afspadsering |
| 100000005 | Andet fravær |

## Workflow
1. **Kladde** (Draft) - Employee creates registration
2. **Afventer godkendelse** (Pending) - Submitted for approval
3. **Godkendt** (Approved) or **Afvist** (Rejected) - Manager decision

## Environment
- **SharePoint Tenant:** `innovaterdk.sharepoint.com`
- **Dev Site:** `innovaterdk.sharepoint.com/sites/projektstyring`
- **Dataverse:** `orgab6f6874.crm4.dynamics.com`

## Build Commands
```bash
npm install          # Install dependencies
gulp serve           # Start dev server
gulp bundle --ship   # Production build
gulp package-solution --ship  # Create .sppkg package
```

## Power Automate Integration
A Power Automate flow handles approval notifications:
- Triggers when status changes to "Afventer godkendelse"
- Sends approval request via Microsoft Approvals
- Updates Dataverse with result
- Sends email confirmation to employee

Flow definition: `power-automate/AbsenceApprovalFlow.json`

## Admin Features
The admin tab is only visible to users in the `ADMIN_EMAILS` array in `AbsenceRegistration.tsx`:
- Currently configured for: `rp@innovater.dk`
- Features: View all registrations, delete individual records, delete all records
- Add more admin emails by modifying the array in the component

## Important Notes
- Danish localization throughout (UI labels, date formats, holidays)
- Working day calculation excludes weekends and Danish public holidays
- Store Bededag removed from holidays (abolished 2024)
- Dates use noon (12:00) internally to avoid timezone issues
