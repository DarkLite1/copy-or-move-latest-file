<# 
    .SYNOPSIS   
        Copy or move the most recently edited file to another folder.

    .DESCRIPTION
        Copy or move the most recently edited file found in the source folder 
        to the destination folder. When FileExtension is given, only files 
        matching the extension will be searched. When DestinationFileName is 
        given, the file will be renamed in the destination folder.

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
        Only a file that starts with a specific string will be copied or moved
        to the destination folder. Files that do not match the string will be 
        excluded.

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
        }
        . $scriptPath @params

        Copy the most recent file from FolderA to FolderB

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
        extension '.txt' and over write the file in the destination when it 
        exists already

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
        'copied.csv'. Because 'OverWrite' is not used an existing file on the 
        destination will not be over written. You can however add this switch 
        if you want that. 
#>
                
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$SourceFolder,
    [Parameter(Mandatory)]
    [String]$DestinationFolder,
    [Parameter(Mandatory)]
    [ValidateSet('Copy', 'Move')]
    [String]$Action,
    [String]$DestinationFileName,
    [String]$FileExtension,
    [String]$FileNameStartsWith,
    [Switch]$OverWrite,
    [String[]]$MailTo,
    [String]$LogFolder = $env:POWERSHELL_LOG_FOLDER,
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
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
                LogFolder    = New-Item -Path "$LogFolder\File and folder" -ItemType 'Directory' -Force -ErrorAction 'Stop'
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

        $mailParams = @{}
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

Process {
    Try {
        #region Create source search parameters
        $getParams = @{
            LiteralPath = $SourceFolder
            File        = $true
            Force       = $true
            ErrorAction = 'Stop'
        }
        if ($FileExtension) {
            $getParams.Filter = '*{0}' -f $FileExtension
        }
        #endregion

        #region Get latest source file
        Try {
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
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}
End {
    Try {
        #region Send email
        $mailParams.Message = (
            $( 
                '<p>'
                $(if ($Action = 'Move') {
                        '<b>Move</b> the most recently edited file'
                    }
                    else {
                        '<b>Copy</b> the most recently edited file'
                    })
                $(if ($FileExtension) { 
                        " with <b>extension '$FileExtension'</b>" 
                    }
                )
                $(" from the <a href=`"$SourceFolder`">source folder</a> to the <a href=`"$DestinationFolder`">destination folder</a>")
                $(if ($OverWrite) { 
                        " and <b>over write the destination file</b> when it exists already" 
                    }
                )
                '.<p>'
            ) | Where-Object { $_ }
        ) -join ''
        
        if ($fileToMove) {
            $mailParams.Message += "
                <p>File details:</p>
                <table>
                    <tr>
                        <th>Destination file name</th>
                        <td><a href=`"$($copyParams.Destination)`">$($copyParams.Destination)</a></td>
                    </tr>
                    <tr>
                        <th>Source file name</th>
                        <td><a href=`"$($fileToMove.FullName)`">$($fileToMove.FullName)</a></td>
                    </tr>
                    <tr>
                        <th>Source file LastWriteTime</th>
                        <td>$(($fileToMove.LastWriteTime).ToString('dd/MM/yyyy HH:mm:ss'))</td>
                    </tr>
                </table>
            "
            $mailParams.subject = 'File {0}' -f $(
                if ($Action = 'Move') { 'moved' } else { 'copied' }
            )
            $mailParams.Priority = 'Normal'
        }
        else {
            $mailParams.Message += '<p>No file found in the source folder matching the search criteria.</p>'
            $mailParams.subject = 'No file {0}' -f $(
                if ($Action = 'Move') { 'moved' } else { 'copied' }
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
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}