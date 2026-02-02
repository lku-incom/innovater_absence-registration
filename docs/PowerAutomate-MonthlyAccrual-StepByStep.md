# Monthly Holiday Accrual Flow - Step-by-Step Setup

Follow these steps to create the flow in Power Automate.

## Step 1: Create the Flow

1. Go to **make.powerautomate.com**
2. Click **+ Create** → **Scheduled cloud flow**
3. Name: `Monthly Holiday Accrual`
4. Configure schedule:
   - Start: Tomorrow at 06:00
   - Repeat every: **1 Month**
5. Click **Create**

## Step 2: Configure the Recurrence Trigger

1. Click on the **Recurrence** trigger
2. Click **Show advanced options**
3. Set:
   - **Time zone**: (UTC+01:00) Brussels, Copenhagen, Madrid, Paris
   - **At these hours**: 6
   - **At these minutes**: 0
   - **On these days**: (leave empty for monthly)

## Step 3: Add Variables

Add these actions in order (click **+ New step** for each):

### 3.1 Initialize variable - CurrentMonth
- Search: `Initialize variable`
- Name: `CurrentMonth`
- Type: `Integer`
- Value: `@{int(formatDateTime(utcNow(), 'MM'))}`

### 3.2 Initialize variable - CurrentYear
- Search: `Initialize variable`
- Name: `CurrentYear`
- Type: `Integer`
- Value: `@{int(formatDateTime(utcNow(), 'yyyy'))}`

### 3.3 Initialize variable - HolidayYear
- Search: `Initialize variable`
- Name: `HolidayYear`
- Type: `String`
- Value (paste exactly):
```
@{if(greaterOrEquals(variables('CurrentMonth'), 9), concat(string(variables('CurrentYear')), '-', string(add(variables('CurrentYear'), 1))), concat(string(sub(variables('CurrentYear'), 1)), '-', string(variables('CurrentYear'))))}
```

### 3.4 Initialize variable - IsJanuary
- Search: `Initialize variable`
- Name: `IsJanuary`
- Type: `Boolean`
- Value: `@{equals(variables('CurrentMonth'), 1)}`

### 3.5 Initialize variable - ProcessedCount
- Search: `Initialize variable`
- Name: `ProcessedCount`
- Type: `Integer`
- Value: `0`

## Step 4: Get Users from Security Group (Recommended)

Using a security group allows you to control exactly which employees receive holiday accruals.

### 4.1 Create Security Group in Azure AD (One-time setup)
1. Go to **https://entra.microsoft.com**
2. Navigate to **Groups** → **All groups** → **New group**
3. Configure:
   - Group type: **Security**
   - Group name: `Holiday Accrual Eligible`
   - Group description: `Employees eligible for automatic holiday accrual`
4. Click **Create**
5. **Copy the Object Id** (you'll need this for the flow)

### 4.2 Manage Group Members
You can manage group members in two ways:
- **In Azure/Entra**: Groups → Select group → Members → Add members
- **In the Absence Registration App**: Admin Panel → "Ferieberettigede medarbejdere" section

### 4.3 Add the Flow Action
1. Click **+ New step**
2. Search: `Office 365 Groups`
3. Select: **List group members**
4. Configure:
   - **Group Id**: Paste your security group's Object Id

**Alternative: Get ALL users (if not using security group):**
1. Search: `Office 365 Users`
2. Select: **Search for users (V2)**
3. Configure:
   - **Search term**: (leave empty to get all)
   - **Top**: `999`

## Step 5: Loop Through Users

1. Click **+ New step**
2. Search: `Apply to each`
3. Select an output: `value` (from the user search)

Inside the loop, add:

### 5.1 Condition - Check Valid User
1. Inside "Apply to each", click **Add an action**
2. Search: `Condition`
3. Configure the condition:
   - Click in the left box, select **Expression**, paste: `empty(items('Apply_to_each')?['mail'])`
   - Operator: `is equal to`
   - Right value: `false`

### 5.2 In the "If yes" branch - Create Dataverse Record

1. Click **Add an action** in the "If yes" branch
2. Search: `Microsoft Dataverse`
3. Select: **Add a new row**
4. Configure:
   - **Table name**: `Accrual Histories` (or search for `cr_accrualhistories`)
   - Fill in fields:

| Field | Value |
|-------|-------|
| Name | `@{items('Apply_to_each')?['displayName']} - Månedlig optjening - @{formatDateTime(utcNow(), 'MMM yyyy')}` |
| Employee Email | `@{items('Apply_to_each')?['mail']}` |
| Employee Name | `@{items('Apply_to_each')?['displayName']}` |
| Holiday Year | `@{variables('HolidayYear')}` |
| Accrual Date | `@{utcNow()}` |
| Accrual Month | `@{variables('CurrentMonth')}` |
| Accrual Year | `@{variables('CurrentYear')}` |
| Days Accrued | `2.08` |
| Feriefridage Accrued | `@{if(variables('IsJanuary'), 5, 0)}` |
| Accrual Type | `Månedlig optjening` (select from dropdown, value 100000000) |
| Notes | `@{if(variables('IsJanuary'), 'Månedlig ferieoptjening (2.08 dage) + Årlig feriefridage (5 dage)', 'Månedlig ferieoptjening (2.08 dage)')}` |

### 5.3 Increment Counter
1. After the Dataverse action, click **Add an action**
2. Search: `Increment variable`
3. Name: `ProcessedCount`
4. Value: `1`

## Step 6: Send Summary Email

1. **Outside** the Apply to each loop, click **+ New step**
2. Search: `Office 365 Outlook`
3. Select: **Send an email (V2)**
4. Configure:
   - **To**: `hr@innovater.dk` (change to your HR email)
   - **Subject**: `Månedlig ferieoptjening kørt - @{formatDateTime(utcNow(), 'MMMM yyyy')}`
   - **Body** (switch to HTML/Code view):

```html
<h2>Månedlig ferieoptjening er gennemført</h2>
<p><strong>Dato:</strong> @{formatDateTime(utcNow(), 'dd-MM-yyyy HH:mm')}</p>
<p><strong>Ferieår:</strong> @{variables('HolidayYear')}</p>
<p><strong>Måned:</strong> @{formatDateTime(utcNow(), 'MMMM yyyy')}</p>
<hr/>
<h3>Resultat</h3>
<ul>
<li><strong>Antal medarbejdere behandlet:</strong> @{variables('ProcessedCount')}</li>
<li><strong>Feriedage tilføjet per medarbejder:</strong> 2.08 dage</li>
@{if(variables('IsJanuary'), '<li><strong>Feriefridage tilføjet per medarbejder:</strong> 5 dage (årlig tildeling)</li>', '')}
</ul>
<p><em>Denne e-mail er automatisk genereret af Power Automate.</em></p>
```

## Step 7: Configure Error Handling (Optional)

1. Click the **...** menu on the "Apply to each" action
2. Select **Configure run after**
3. Check: "is successful", "has failed", "is skipped", "has timed out"

## Step 8: Save and Test

1. Click **Save** (top right)
2. Click **Test** → **Manually** → **Test**
3. Wait for the flow to complete
4. Check:
   - Dataverse table for new records
   - Your email for the summary

## Flow Overview

When complete, your flow should look like this:

```
┌─────────────────────────────┐
│ Recurrence (Monthly)        │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Initialize CurrentMonth     │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Initialize CurrentYear      │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Initialize HolidayYear      │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Initialize IsJanuary        │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Initialize ProcessedCount   │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ List group members          │
│ (Security Group)            │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Apply to each               │
│  ┌────────────────────────┐ │
│  │ Condition (has email?) │ │
│  │  ├─ Yes: Add Dataverse │ │
│  │  │       Increment var │ │
│  │  └─ No: (skip)         │ │
│  └────────────────────────┘ │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│ Send summary email          │
└─────────────────────────────┘
```

## Troubleshooting

### "Table not found" error
- Make sure your Dataverse environment is connected
- The table name might be `cr_accrualhistories` (logical name) or shown as `Accrual Histories` (display name)

### Users not found
- Check if your account has permissions to read Azure AD users
- Try using Microsoft Graph connector instead of Office 365 Users

### Wrong holiday year
- The formula calculates: Sept-Dec = current-next, Jan-Aug = previous-current
- Test by changing your system date or the formula temporarily

## Next Steps

After testing:
1. Turn off the flow temporarily if you don't want it to run yet
2. Adjust the schedule if needed (different time, different day of month)
3. Add filters to the user query if you only want certain departments
