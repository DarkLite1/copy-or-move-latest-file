#Requires -Version 5.1
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

<#
    .SYNOPSIS
        Copy or move the most recently edited file to another folder.

    .DESCRIPTION
        Copy or move the most recently edited file found in the source folder
        to the destination folder. When 'FileExtension' is given, only files
        matching the extension will be searched. When 'DestinationFileName' is
        given, the file will be renamed in the destination folder. When
        'FileNameStartsWith' is given, only files that start with that specific
        string are checked in the source folder.

    .PARAMETER Action
        When set to 'Copy' the source file will be left alone and when set to
        'Move' the source file will be deleted after copying.

    .PARAMETER SourceFolder
        Folder where to search for the most recently edited file.

    .PARAMETER DestinationFolder
        Folder to where to copy the most recently edited file.

    .PARAMETER FileExtension
        The file extension to look for in the source folder. If none is given
        the latest file is selected regardless of extension (Ex. '.csv').

    .PARAMETER FileNameStartsWith
        Only files in the source folder where the file name starts with a
        specific string will be searched, other files will be excluded.

    .PARAMETER DestinationFileName
        New name for the copied file without its extension (Ex. 'Copied').

    .PARAMETER OverWrite
        If the file is already there it will be overwritten when OverWrite is
        set to true. When OverWrite is set to false an error email will be sent
        to the users in MailTo.

    .PARAMETER MailTo
        Inform users by sending an email after running the script on success
        and failure. When MailTo is empty no email will be sent.

    .NOTES
        A copied file can be renamed in the destination folder but its
        extension cannot be renamed. Renaming an extension in the destination
        folder is not supported.

    .EXAMPLE
        $params = @{
            ScriptName        = 'Copy latest file'
            Action            = 'Copy'
            SourceFolder      = 'C:\FolderA'
            DestinationFolder = 'C:\FolderB'
            MailTo            = 'gmail@chuckNorris.com'
        }
        . $scriptPath @params

        Copy the most recent file from FolderA to FolderB, regardless its file
        extension or file name, and send a summary mail to Chuck Norris.

        In case the file is already in the destination folder it will not be
        overwritten because 'OverWrite' is not used.

    .EXAMPLE
        $params = @{
            ScriptName        = 'Move latest txt file'
            Action            = 'Move'
            SourceFolder      = 'C:\FolderA'
            DestinationFolder = 'C:\FolderB'
            FileExtension     = '.txt'
            OverWrite         = $true
        }
        . $scriptPath @params

        Move the most recent file from FolderA to FolderB that has file
        extension '.txt', regardless of its name, and overwrite the file in the
        destination folder when it already exists.

        No mail is sent because 'MailTo' is not used.

    .EXAMPLE
        $params = @{
            ScriptName          = 'Copy latest csv file'
            Action              = 'Copy'
            SourceFolder        = 'C:\FolderA'
            DestinationFolder   = 'C:\FolderB'
            FileExtension       = '.csv'
            DestinationFileName = 'copied'
        }
        . $scriptPath @params

        Copy the most recent file from FolderA to FolderB that has file
        extension '.csv' and rename the file in the destination folder to
        'copied.csv'.

        Because 'OverWrite' is not used an existing file in the destination
        folder will not be overwritten. No mails are sent because 'MailTo'
        is not used either.

    .EXAMPLE
        $params = @{
            ScriptName          = 'Copy latest zip file that starts with backup'
            Action              = 'Copy'
            SourceFolder        = 'C:\FolderA'
            DestinationFolder   = 'C:\FolderB'
            DestinationFileName = 'latest backup'
            FileExtension       = '.zip'
            FileNameStartsWith  = 'backup'
            OverWrite           = $true
            MailTo              = 'gmail@chuckNorris.com'
        }
        . $scriptPath @params

        Copy the most recent file from FolderA to FolderB that has file
        extension '.zip' and a file name that starts with 'backup'. Rename
        the file in the destination folder to 'latest backup.zip' and overwrite
        the file when it already exists. Send a summary mail to Chuck Norris.
#>

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Copy or move latest file\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $Error.Clear()

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            if (-not ($MailTo = $file.MailTo)) {
                throw "Property 'MailTo' not found"
            }

            if (-not ($SourceFolder = $file.SourceFolder)) {
                throw "Property 'SourceFolder' not found"
            }

            if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
                throw "Source folder '$SourceFolder' not found"
            }

            if (-not ($DestinationFolder = $file.DestinationFolder)) {
                throw "Property 'DestinationFolder' not found"
            }

            if (-not ($Action = $file.Action)) {
                throw "Property 'Action' not found"
            }

            if (-not (Test-Path -LiteralPath $DestinationFolder -PathType Container)) {
                throw "Destination folder '$DestinationFolder' not found"
            }

            if ($file.Action -notMatch '^Copy$|^Move$') {
                throw "Property 'Action' with value '$($file.Action)' is not valid. Accepted values are 'Copy' or 'Move'."
            }

            try {
                $OverWrite = [Boolean]::Parse($file.OverWrite)
            }
            catch {
                throw "Property 'Overwrite' is not a boolean value"
            }

            $DestinationFileName = $file.DestinationFileName
            $FileExtension = $file.SearchFor.FileExtension
            $FileNameStartsWith = $file.SearchFor.FileNameStartsWith
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        $mailParams = @{}
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
        #region Get latest source file
        Try {
            $getParams = @{
                LiteralPath = $SourceFolder
                Filter      = '{0}*{1}' -f $FileNameStartsWith, $FileExtension
                File        = $true
                Force       = $true
                ErrorAction = 'Stop'
            }
            $fileToMove = Get-ChildItem @getParams |
            Sort-Object LastWriteTime | Select-Object -Last 1
        }
        Catch {
            throw "Failed checking the source folder '$SourceFolder': $_"
        }
        #endregion

        if ($fileToMove) {
            #region Create destination copy parameters
            $joinParam = @{
                Path      = $DestinationFolder
                ChildPath = if ($DestinationFileName) {
                    $DestinationFileName + $fileToMove.Extension
                }
                else { $fileToMove.Name }
            }
            $copyParams = @{
                LiteralPath = $fileToMove.FullName
                Destination = Join-Path @joinParam
                Force       = $true
                ErrorAction = 'Stop'
            }
            #endregion

            #region Test if destination file already exists
            if (
                (-not $OverWrite) -and
                (Test-Path -LiteralPath $copyParams.Destination -PathType Leaf)
            ) {
                throw "The file '$($copyParams.Destination)' already exists in the destination folder, use 'OverWrite' if you want to over write this file."
            }
            #endregion

            #region Copy source file to destination folder
            Try {
                Copy-Item @copyParams
            }
            Catch {
                throw "Failed to copy file '$($fileToMove.FullName)' to the destination folder '$DestinationFolder': $_"
            }
            #endregion

            #region Remove source file
            if (
                ($Action -eq 'Move') -and
                (Test-Path -LiteralPath $fileToMove.FullName -PathType Leaf)
            ) {
                Remove-Item -LiteralPath $fileToMove.FullName -Force
            }
            #endregion
        }
        else {
            Write-Warning "No file found in source folder '$SourceFolder'"
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
        #region Send email
        $mailParams.Message = (
            $(
                '<p>'
                $(if ($Action -eq 'Move') {
                        '<b>Move</b> the most recently edited file'
                    }
                    else {
                        '<b>Copy</b> the most recently edited file'
                    })
                $(if ($FileExtension) {
                        " with <b>extension '$FileExtension'</b>"
                    }
                )
                $(if ($FileNameStartsWith) {
                        " where the file name <b>starts with '$FileNameStartsWith'</b>"
                    }
                )
                $(" from the <a href=`"$SourceFolder`">source folder</a> to the <a href=`"$DestinationFolder`">destination folder</a>")
                $(if ($OverWrite) {
                        " and <b>overwrite</b> the destination file when it exists already"
                    }
                )
                '.<p>'
            ) | Where-Object { $_ }
        ) -join ''

        if ($fileToMove) {
            $mailParams.Message += "
                <p>$(if ($Action -eq 'Move') { 'Moved' } else { 'Copied' }) file:</p>
                <table>
                    <tr>
                        <th>Source</th>
                        <td>
                            <a href=`"$($fileToMove.FullName)`">$($fileToMove.Name)</a><br>
                            LastWriteTime: $(($fileToMove.LastWriteTime).ToString('dd/MM/yyyy HH:mm:ss'))
                        </td>
                    </tr>
                    <tr>
                        <th>Destination</th>
                        <td><a href=`"$($copyParams.Destination)`">$($joinParam.ChildPath)</a></td>
                    </tr>
                </table>
            "
            $mailParams.subject = 'File {0}' -f $(
                if ($Action -eq 'Move') { 'moved' } else { 'copied' }
            )
            $mailParams.Priority = 'Normal'
        }
        else {
            $mailParams.Message += '<p>No file found in the source folder matching the search criteria.</p>'
            $mailParams.subject = 'No file {0}' -f $(
                if ($Action -eq 'Move') { 'moved' } else { 'copied' }
            )
            $mailParams.Priority = 'High'
        }


        $mailParams += @{
            To        = $ScriptAdmin
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }
        if ($MailTo) {
            $mailParams.To = $MailTo
            $mailParams.Bcc = $ScriptAdmin
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}