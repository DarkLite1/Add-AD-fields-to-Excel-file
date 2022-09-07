#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, Toolbox.ActiveDirectory

<#
    .SYNOPSIS
        Add AD properties to an Excel worksheet.

    .DESCRIPTION
        Complement user data in an Excel sheet with data from active directory

        This script is useful for retrieving data from the active directory for 
        all the users found in the rows of the Excel sheet. The match between 
        the row in Excel and the active directory is based on the key value 
        pair in the variable 'Match'. 

        A new Excel file is created containing the requested AD fields with the 
        prefix 'ad' (Ex. 'adName', 'adEnabled', 'adDisplayName', ...). This file
        is then sent to the e-mail addresses defined in 'MailTo'.

    .PARAMETER ExcelFile
        The path to the Excel file.

    .PARAMETER Match
        One or multiple key value pairs that are used to create the AD search 
        filter to find the matching user account.

        Ex:
        Match      = @{
            # ExcelField  = AdField
            'Logon name'  = 'SamAccountName' 
        }

    .PARAMETER AdProperties
        The data fields that need to be retrieved from the active directory 
        to complement the Excel file.

    .PARAMETER MailTo
        List of e-mail addresses to where the e-mail will be sent with the new 
        Excel file in attachment.

    .EXAMPLE
        $params = @{
            ScriptName = 'Add AD fields to Excel file'
            ExcelFile  = './test.xlsx'
            Match      = @{
                # ExcelField = AdField
                'Last Name'  = 'SurName' 
                'First name' = 'GivenName' 
            }
            AdProperties = @(
                'Mail', 'Name', 'SamAccountName', 'Enabled',
                'UserPrincipalName', 'Office'
            )
            MailTo = @( 'bob@contoso.com' )
        }
        & './Add AD fields to Excel file.ps1' @params
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [String]$ExcelFile,
    [Parameter(Mandatory)]
    [HashTable]$Match,
    [Parameter(Mandatory)]
    [String[]]$AdProperties,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\Add AD fields to Excel file\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $excelFileContent = Import-Excel -Path $ExcelFile -EA Stop

        $adUsersCount = 0
        foreach ($row in $excelFileContent) {
            #region Create AD search filter
            $filter = ($Match.GetEnumerator() | ForEach-Object {
                    "{0} -eq ""{1}""" -f $_.Value, $(
                        $row.$($_.Key)
                    )
                }) -join ' -and '
            Write-Verbose "AD Search filter '$filter'"
            #endregion

            #region Get AD user account
            $params = @{
                Properties = $AdProperties
                Filter     = $filter
            }
            $adUser = Get-ADUser @params
            #endregion

            #region Verbose
            if ($adUser) {
                $adUsersCount++
                Write-Verbose "AD user account '$($adUser.SamAccountName)'"
            }
            else {
                Write-Warning 'No AD user account found'
            }
            #endregion

            #region Add AD properties
            foreach ($property in $AdProperties) {
                $addMemberParams = @{
                    NotePropertyName  = "ad$property" 
                    NotePropertyValue = $adUser.$property
                }
                $row | Add-Member @addMemberParams
            }
            #endregion
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        $excelParams = @{
            Path         = "$LogFile.xlsx"
            AutoSize     = $true
            FreezeTopRow = $true
        }

        #region Export data to Excel
        if ($excelFileContent) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Data'

            $params = @{
                Property = '*', 
                @{
                    name       = 'adOu'
                    expression = { 
                        if ($_.adCanonicalName ) { 
                            ConvertTo-OuNameHC $_.adCanonicalName 
                        }
                    }
                }
            }
            if ($AdProperties -contains 'manager') {
                $params.Property += @{
                    name       = 'adManager'
                    expression = { 
                        if ($_.adManager) { Get-ADDisplayNameHC $_.adManager }
                    } 
                }
            }
            $excelFileContent | Select-Object @params |
            Export-Excel @excelParams
        }
        #endregion

        #region Export errors to Excel
        if ($Error) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Errors'

            $Error.Exception.Message | Select-Object -Unique | 
            Export-Excel @excelParams
        }
        #endregion

        #region Send mail
        $Message = "
            <p>The Excel sheet contains <b>{0} rows</b> for which <b>{1} matching AD user accounts</b> were found.</p>
            <p>The following AD fields were added:
            {2}</p>
            <p><i>* Check the attachment for details</i></p>" -f 
        $($excelFileContent.count),
        $adUsersCount,
        $($AdProperties | ConvertTo-HtmlListHC)

        $mailParams = @{
            To          = $MailTo
            Bcc         = $ScriptAdmin
            Subject     = "{0} AD users, {1} rows" -f
            $adUsersCount, $($excelFileContent.count)
            Message     = $Message
            Attachments = $excelParams.Path
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = "$LogFile - Mail.html"
        }
        Get-ScriptRuntimeHC -Stop  
        Send-MailHC @mailParams
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $)"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}