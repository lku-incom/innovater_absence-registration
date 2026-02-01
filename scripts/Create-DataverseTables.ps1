# Create-DataverseTables.ps1
# Creates Holiday Balance and Accrual History tables in Dataverse
# Requires: Azure AD authentication to Dataverse

param(
    [string]$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
)

# Get access token using interactive login
Write-Host "Authenticating to Dataverse..." -ForegroundColor Cyan

# Use MSAL.PS module if available, otherwise use device code flow
$resource = "$DataverseUrl/"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d" # Power Platform CLI client ID

# Try to get token using Azure CLI if available
$token = $null
try {
    $azAccount = az account get-access-token --resource $DataverseUrl 2>$null | ConvertFrom-Json
    if ($azAccount) {
        $token = $azAccount.accessToken
        Write-Host "Using Azure CLI authentication" -ForegroundColor Green
    }
} catch {
    Write-Host "Azure CLI not available, will use device code flow" -ForegroundColor Yellow
}

if (-not $token) {
    # Use device code flow
    Write-Host "Please authenticate using the browser window that will open..." -ForegroundColor Yellow

    # Device code flow
    $deviceCodeUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode"
    $tokenUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
    $scope = "$DataverseUrl/.default"

    $deviceCodeResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -Body @{
        client_id = $clientId
        scope = $scope
    }

    Write-Host $deviceCodeResponse.message -ForegroundColor Cyan

    # Poll for token
    $pollInterval = $deviceCodeResponse.interval
    $expiresIn = $deviceCodeResponse.expires_in
    $startTime = Get-Date

    while ((Get-Date) -lt $startTime.AddSeconds($expiresIn)) {
        Start-Sleep -Seconds $pollInterval
        try {
            $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body @{
                grant_type = "urn:ietf:params:oauth:grant-type:device_code"
                client_id = $clientId
                device_code = $deviceCodeResponse.device_code
            }
            $token = $tokenResponse.access_token
            Write-Host "Authentication successful!" -ForegroundColor Green
            break
        } catch {
            if ($_.Exception.Response.StatusCode -ne 400) {
                throw
            }
            # Authorization pending, continue polling
        }
    }
}

if (-not $token) {
    Write-Error "Failed to obtain access token"
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
    "Prefer" = "return=representation"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Function to create entity/table
function New-DataverseEntity {
    param(
        [string]$SchemaName,
        [string]$DisplayName,
        [string]$PluralName,
        [string]$Description
    )

    Write-Host "Creating entity: $DisplayName..." -ForegroundColor Cyan

    $entityDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.EntityMetadata"
        "SchemaName" = $SchemaName
        "DisplayName" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayName
                    "LanguageCode" = 1033
                }
            )
        }
        "DisplayCollectionName" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $PluralName
                    "LanguageCode" = 1033
                }
            )
        }
        "Description" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $Description
                    "LanguageCode" = 1033
                }
            )
        }
        "OwnershipType" = "UserOwned"
        "HasActivities" = $false
        "HasNotes" = $false
        "IsActivity" = $false
        "PrimaryNameAttribute" = "cr_name"
        "Attributes" = @(
            @{
                "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
                "SchemaName" = "cr_name"
                "RequiredLevel" = @{
                    "Value" = "ApplicationRequired"
                }
                "MaxLength" = 200
                "DisplayName" = @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                    "LocalizedLabels" = @(
                        @{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            "Label" = "Name"
                            "LanguageCode" = 1033
                        }
                    )
                }
                "IsPrimaryName" = $true
            }
        )
    }

    try {
        $response = Invoke-RestMethod -Uri "$apiUrl/EntityDefinitions" -Method POST -Headers $headers -Body ($entityDef | ConvertTo-Json -Depth 20)
        Write-Host "Entity created successfully: $SchemaName" -ForegroundColor Green
        return $response.MetadataId
    } catch {
        if ($_.Exception.Response.StatusCode -eq 409) {
            Write-Host "Entity already exists: $SchemaName" -ForegroundColor Yellow
            # Get existing entity ID
            $existing = Invoke-RestMethod -Uri "$apiUrl/EntityDefinitions(LogicalName='$($SchemaName.ToLower())')" -Method GET -Headers $headers
            return $existing.MetadataId
        }
        Write-Error "Failed to create entity: $_"
        throw
    }
}

# Function to add attribute to entity
function Add-DataverseAttribute {
    param(
        [string]$EntityLogicalName,
        [hashtable]$AttributeDef
    )

    $attrName = $AttributeDef.SchemaName
    Write-Host "  Adding attribute: $attrName..." -ForegroundColor Gray

    try {
        $response = Invoke-RestMethod -Uri "$apiUrl/EntityDefinitions(LogicalName='$EntityLogicalName')/Attributes" -Method POST -Headers $headers -Body ($AttributeDef | ConvertTo-Json -Depth 20)
        Write-Host "  Attribute added: $attrName" -ForegroundColor Green
    } catch {
        if ($_.Exception.Response.StatusCode -eq 409) {
            Write-Host "  Attribute already exists: $attrName" -ForegroundColor Yellow
        } else {
            Write-Warning "  Failed to add attribute $attrName : $_"
        }
    }
}

# ============================================
# Create Holiday Balance Table
# ============================================
Write-Host "`n=== Creating Holiday Balance Table ===" -ForegroundColor Magenta

$holidayBalanceId = New-DataverseEntity -SchemaName "cr_holidaybalance" `
    -DisplayName "Holiday Balance" `
    -PluralName "Holiday Balances" `
    -Description "Tracks employee holiday balance per holiday year according to Danish holiday law"

# Add attributes to Holiday Balance
$holidayBalanceAttrs = @(
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_employeeemail"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 100
        "FormatName" = @{ "Value" = "Email" }
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Employee Email"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_employeename"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 100
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Employee Name"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_holidayyear"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 20
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Holiday Year"; "LanguageCode" = 1033 }) }
        "Description" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Holiday year period (Sept-Aug), e.g., 2024-2025"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_accrueddays"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 50
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Accrued Days"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_useddays"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 50
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Used Days"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_pendingdays"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 50
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Pending Days"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_availabledays"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = -50
        "MaxValue" = 100
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Available Days"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_carriedoverdays"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 25
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Carried Over Days"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_feriefridageaccrued"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 20
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Feriefridage Accrued"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_feriefridageused"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 20
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Feriefridage Used"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_feriefridageavailable"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = -20
        "MaxValue" = 20
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Feriefridage Available"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
        "SchemaName" = "cr_lastaccrualdate"
        "RequiredLevel" = @{ "Value" = "None" }
        "Format" = "DateOnly"
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Last Accrual Date"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
        "SchemaName" = "cr_employmentstartdate"
        "RequiredLevel" = @{ "Value" = "None" }
        "Format" = "DateOnly"
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Employment Start Date"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata"
        "SchemaName" = "cr_isactive"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "DefaultValue" = $true
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Is Active"; "LanguageCode" = 1033 }) }
        "OptionSet" = @{
            "TrueOption" = @{ "Value" = 1; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Yes"; "LanguageCode" = 1033 }) } }
            "FalseOption" = @{ "Value" = 0; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "No"; "LanguageCode" = 1033 }) } }
        }
    }
)

foreach ($attr in $holidayBalanceAttrs) {
    Add-DataverseAttribute -EntityLogicalName "cr_holidaybalance" -AttributeDef $attr
}

# ============================================
# Create Accrual History Table
# ============================================
Write-Host "`n=== Creating Accrual History Table ===" -ForegroundColor Magenta

$accrualHistoryId = New-DataverseEntity -SchemaName "cr_accrualhistory" `
    -DisplayName "Accrual History" `
    -PluralName "Accrual History Records" `
    -Description "Audit trail for monthly holiday accruals"

# Add attributes to Accrual History
$accrualHistoryAttrs = @(
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_employeeemail"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 100
        "FormatName" = @{ "Value" = "Email" }
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Employee Email"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_employeename"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 100
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Employee Name"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_holidayyear"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MaxLength" = 20
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Holiday Year"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
        "SchemaName" = "cr_accrualdate"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "Format" = "DateOnly"
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Accrual Date"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.IntegerAttributeMetadata"
        "SchemaName" = "cr_accrualmonth"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MinValue" = 1
        "MaxValue" = 12
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Accrual Month"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.IntegerAttributeMetadata"
        "SchemaName" = "cr_accrualyear"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "MinValue" = 2020
        "MaxValue" = 2100
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Accrual Year"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_daysaccrued"
        "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 10
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Days Accrued"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_feriefridageaccrued"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 5
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Feriefridage Accrued"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = "cr_balanceafteraccrual"
        "RequiredLevel" = @{ "Value" = "None" }
        "Precision" = 2
        "MinValue" = 0
        "MaxValue" = 100
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Balance After Accrual"; "LanguageCode" = 1033 }) }
    },
    @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        "SchemaName" = "cr_notes"
        "RequiredLevel" = @{ "Value" = "None" }
        "MaxLength" = 1000
        "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Notes"; "LanguageCode" = 1033 }) }
    }
)

foreach ($attr in $accrualHistoryAttrs) {
    Add-DataverseAttribute -EntityLogicalName "cr_accrualhistory" -AttributeDef $attr
}

# Create Accrual Type choice/picklist
Write-Host "`n  Adding Accrual Type picklist..." -ForegroundColor Gray
$accrualTypeAttr = @{
    "@odata.type" = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
    "SchemaName" = "cr_accrualtype"
    "RequiredLevel" = @{ "Value" = "ApplicationRequired" }
    "DisplayName" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Accrual Type"; "LanguageCode" = 1033 }) }
    "OptionSet" = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.OptionSetMetadata"
        "IsGlobal" = $false
        "OptionSetType" = "Picklist"
        "Options" = @(
            @{ "Value" = 100000000; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Monthly Accrual"; "LanguageCode" = 1033 }) } }
            @{ "Value" = 100000001; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Year Start Carryover"; "LanguageCode" = 1033 }) } }
            @{ "Value" = 100000002; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Manual Adjustment"; "LanguageCode" = 1033 }) } }
            @{ "Value" = 100000003; "Label" = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; "LocalizedLabels" = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; "Label" = "Initial Balance"; "LanguageCode" = 1033 }) } }
        )
    }
}
Add-DataverseAttribute -EntityLogicalName "cr_accrualhistory" -AttributeDef $accrualTypeAttr

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "Table creation completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "`nTables created:"
Write-Host "  - Holiday Balance (cr_holidaybalance)"
Write-Host "  - Accrual History (cr_accrualhistory)"
Write-Host "`nNote: You may need to publish customizations in Power Apps to see the tables."
