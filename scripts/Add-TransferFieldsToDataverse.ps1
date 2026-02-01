# Add Transfer Tracking Fields to Dataverse Holiday Balance Table
# Based on Danish Holiday Law (Ferieloven) transfer rules

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

# Authenticate using device code flow
Write-Host "=== Authenticating to Dataverse ===" -ForegroundColor Magenta
$deviceCodeUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode"
$tokenUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
$scope = "$DataverseUrl/.default"

$deviceCodeResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -Body @{
    client_id = $clientId
    scope = $scope
}

Write-Host $deviceCodeResponse.message -ForegroundColor Yellow

$pollInterval = $deviceCodeResponse.interval
$expiresIn = $deviceCodeResponse.expires_in
$startTime = Get-Date
$token = $null

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
        if ($_.Exception.Response.StatusCode -ne 400) { throw }
    }
}

if (-not $token) {
    Write-Error "Failed to authenticate"
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Get the Holiday Balance entity metadata
Write-Host "`n=== Getting Holiday Balance Entity Metadata ===" -ForegroundColor Cyan
$entityName = "cr_holidaybalance"

# Function to create a decimal column
function Add-DecimalColumn {
    param(
        [string]$SchemaName,
        [string]$DisplayName,
        [string]$DisplayNameDa,
        [string]$Description,
        [string]$DescriptionDa,
        [decimal]$MinValue = 0,
        [decimal]$MaxValue = 50,
        [int]$Precision = 2
    )

    $columnDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
        "SchemaName" = $SchemaName
        "RequiredLevel" = @{
            "Value" = "None"
            "CanBeChanged" = $true
        }
        "DisplayName" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayName
                    "LanguageCode" = 1033
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayNameDa
                    "LanguageCode" = 1030
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
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DescriptionDa
                    "LanguageCode" = 1030
                }
            )
        }
        "MinValue" = $MinValue
        "MaxValue" = $MaxValue
        "Precision" = $Precision
    }

    return $columnDef
}

# Function to create a boolean column
function Add-BooleanColumn {
    param(
        [string]$SchemaName,
        [string]$DisplayName,
        [string]$DisplayNameDa,
        [string]$Description,
        [string]$DescriptionDa
    )

    $columnDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata"
        "SchemaName" = $SchemaName
        "RequiredLevel" = @{
            "Value" = "None"
            "CanBeChanged" = $true
        }
        "DisplayName" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayName
                    "LanguageCode" = 1033
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayNameDa
                    "LanguageCode" = 1030
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
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DescriptionDa
                    "LanguageCode" = 1030
                }
            )
        }
        "OptionSet" = @{
            "TrueOption" = @{
                "Value" = 1
                "Label" = @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                    "LocalizedLabels" = @(
                        @{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            "Label" = "Yes"
                            "LanguageCode" = 1033
                        },
                        @{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            "Label" = "Ja"
                            "LanguageCode" = 1030
                        }
                    )
                }
            }
            "FalseOption" = @{
                "Value" = 0
                "Label" = @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                    "LocalizedLabels" = @(
                        @{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            "Label" = "No"
                            "LanguageCode" = 1033
                        },
                        @{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            "Label" = "Nej"
                            "LanguageCode" = 1030
                        }
                    )
                }
            }
        }
        "DefaultValue" = $false
    }

    return $columnDef
}

# Function to create a date column
function Add-DateColumn {
    param(
        [string]$SchemaName,
        [string]$DisplayName,
        [string]$DisplayNameDa,
        [string]$Description,
        [string]$DescriptionDa
    )

    $columnDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
        "SchemaName" = $SchemaName
        "RequiredLevel" = @{
            "Value" = "None"
            "CanBeChanged" = $true
        }
        "DisplayName" = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            "LocalizedLabels" = @(
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayName
                    "LanguageCode" = 1033
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DisplayNameDa
                    "LanguageCode" = 1030
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
                },
                @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    "Label" = $DescriptionDa
                    "LanguageCode" = 1030
                }
            )
        }
        "Format" = "DateOnly"
        "DateTimeBehavior" = @{
            "Value" = "UserLocal"
        }
    }

    return $columnDef
}

# Define the new columns to add
$columnsToAdd = @(
    @{
        Type = "Decimal"
        SchemaName = "cr_transferredindays"
        DisplayName = "Transferred In Days"
        DisplayNameDa = "Overforte dage (ind)"
        Description = "Vacation days transferred in from previous holiday year (requires written agreement)"
        DescriptionDa = "Feriedage overfort fra forrige ferieaar (kraever skriftlig aftale)"
        MinValue = 0
        MaxValue = 50
    },
    @{
        Type = "Decimal"
        SchemaName = "cr_transferredoutdays"
        DisplayName = "Transferred Out Days"
        DisplayNameDa = "Overforte dage (ud)"
        Description = "Vacation days to be transferred to next holiday year (max 5 per year)"
        DescriptionDa = "Feriedage der skal overfores til naeste ferieaar (max 5 pr. aar)"
        MinValue = 0
        MaxValue = 50
    },
    @{
        Type = "Boolean"
        SchemaName = "cr_hastransferagreement"
        DisplayName = "Has Transfer Agreement"
        DisplayNameDa = "Har overforelsesaftale"
        Description = "Whether a written agreement exists for transferring vacation days"
        DescriptionDa = "Om der er en skriftlig aftale om overforsel af feriedage"
    },
    @{
        Type = "Date"
        SchemaName = "cr_transferagreementdate"
        DisplayName = "Transfer Agreement Date"
        DisplayNameDa = "Overforelsesaftaledato"
        Description = "Date when transfer agreement was made (must be by December 31)"
        DescriptionDa = "Dato hvor overforelsesaftalen blev indgaaet (skal vaere inden 31. december)"
    },
    @{
        Type = "Decimal"
        SchemaName = "cr_feriefridagetransferredin"
        DisplayName = "Feriefridage Transferred In"
        DisplayNameDa = "Feriefridage overfort (ind)"
        Description = "Contract-based extra days transferred from previous year (company policy)"
        DescriptionDa = "Kontraktbaserede ekstra dage overfort fra forrige aar (firmapolitik)"
        MinValue = 0
        MaxValue = 20
    },
    @{
        Type = "Decimal"
        SchemaName = "cr_feriefridagetransferredout"
        DisplayName = "Feriefridage Transferred Out"
        DisplayNameDa = "Feriefridage overfort (ud)"
        Description = "Contract-based extra days to transfer to next year (company policy)"
        DescriptionDa = "Kontraktbaserede ekstra dage til overforsel til naeste aar (firmapolitik)"
        MinValue = 0
        MaxValue = 20
    }
)

Write-Host "`n=== Adding Transfer Fields to Holiday Balance Table ===" -ForegroundColor Magenta

$successCount = 0
$errorCount = 0

foreach ($column in $columnsToAdd) {
    Write-Host "`nAdding column: $($column.SchemaName)..." -ForegroundColor Cyan

    try {
        switch ($column.Type) {
            "Decimal" {
                $columnDef = Add-DecimalColumn `
                    -SchemaName $column.SchemaName `
                    -DisplayName $column.DisplayName `
                    -DisplayNameDa $column.DisplayNameDa `
                    -Description $column.Description `
                    -DescriptionDa $column.DescriptionDa `
                    -MinValue $column.MinValue `
                    -MaxValue $column.MaxValue
            }
            "Boolean" {
                $columnDef = Add-BooleanColumn `
                    -SchemaName $column.SchemaName `
                    -DisplayName $column.DisplayName `
                    -DisplayNameDa $column.DisplayNameDa `
                    -Description $column.Description `
                    -DescriptionDa $column.DescriptionDa
            }
            "Date" {
                $columnDef = Add-DateColumn `
                    -SchemaName $column.SchemaName `
                    -DisplayName $column.DisplayName `
                    -DisplayNameDa $column.DisplayNameDa `
                    -Description $column.Description `
                    -DescriptionDa $column.DescriptionDa
            }
        }

        $body = $columnDef | ConvertTo-Json -Depth 10
        $uri = "$apiUrl/EntityDefinitions(LogicalName='$entityName')/Attributes"

        $response = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body
        Write-Host "  SUCCESS: $($column.DisplayName)" -ForegroundColor Green
        $successCount++
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*already exists*" -or $errorMessage -like "*duplicate*") {
            Write-Host "  SKIPPED: Column already exists" -ForegroundColor Yellow
        } else {
            Write-Host "  ERROR: $errorMessage" -ForegroundColor Red
            $errorCount++
        }
    }
}

# Now add new options to the Accrual Type choice field
Write-Host "`n=== Adding New Accrual Type Options ===" -ForegroundColor Magenta

$accrualTypeOptions = @(
    @{
        Value = 100000004
        LabelEn = "Year End Transfer Out"
        LabelDa = "Aarsslut overforsel (ud)"
    },
    @{
        Value = 100000005
        LabelEn = "Feriefridage Accrual"
        LabelDa = "Feriefridage optjening"
    }
)

foreach ($option in $accrualTypeOptions) {
    Write-Host "`nAdding option: $($option.LabelEn) (value: $($option.Value))..." -ForegroundColor Cyan

    try {
        $optionBody = @{
            "Value" = $option.Value
            "Label" = @{
                "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                "LocalizedLabels" = @(
                    @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                        "Label" = $option.LabelEn
                        "LanguageCode" = 1033
                    },
                    @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                        "Label" = $option.LabelDa
                        "LanguageCode" = 1030
                    }
                )
            }
        } | ConvertTo-Json -Depth 10

        # Insert option into the local option set
        $uri = "$apiUrl/InsertOptionValue"
        $insertBody = @{
            "AttributeLogicalName" = "cr_accrualtype"
            "EntityLogicalName" = "cr_accrualhistory"
            "Value" = $option.Value
            "Label" = @{
                "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                "LocalizedLabels" = @(
                    @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                        "Label" = $option.LabelEn
                        "LanguageCode" = 1033
                    },
                    @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                        "Label" = $option.LabelDa
                        "LanguageCode" = 1030
                    }
                )
            }
        } | ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $insertBody
        Write-Host "  SUCCESS: Added option $($option.LabelEn)" -ForegroundColor Green
        $successCount++
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*already exists*" -or $errorMessage -like "*duplicate*" -or $errorMessage -like "*The value*is out of range*") {
            Write-Host "  SKIPPED: Option may already exist" -ForegroundColor Yellow
        } else {
            Write-Host "  ERROR: $errorMessage" -ForegroundColor Red
            $errorCount++
        }
    }
}

# Publish customizations
Write-Host "`n=== Publishing Customizations ===" -ForegroundColor Magenta
try {
    $publishBody = @{
        "ParameterXml" = "<importexportxml><entities><entity>cr_holidaybalance</entity><entity>cr_accrualhistory</entity></entities></importexportxml>"
    } | ConvertTo-Json

    $response = Invoke-RestMethod -Uri "$apiUrl/PublishXml" -Method POST -Headers $headers -Body $publishBody
    Write-Host "Customizations published successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Could not publish customizations. You may need to publish manually in make.powerapps.com" -ForegroundColor Yellow
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "Schema Update Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "`nSummary:"
Write-Host "  - Columns added: $successCount"
Write-Host "  - Errors: $errorCount"
Write-Host "`nNew fields added to Holiday Balance table:"
Write-Host "  - cr_transferredindays (Transferred In Days)"
Write-Host "  - cr_transferredoutdays (Transferred Out Days)"
Write-Host "  - cr_hastransferagreement (Has Transfer Agreement)"
Write-Host "  - cr_transferagreementdate (Transfer Agreement Date)"
Write-Host "  - cr_feriefridagetransferredin (Feriefridage Transferred In)"
Write-Host "  - cr_feriefridagetransferredout (Feriefridage Transferred Out)"
Write-Host "`nNew Accrual Type options added:"
Write-Host "  - Year End Transfer Out (100000004)"
Write-Host "  - Feriefridage Accrual (100000005)"
