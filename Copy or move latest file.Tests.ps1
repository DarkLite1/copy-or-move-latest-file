#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName          = 'Test (Brecht)'
        Action              = 'Copy'
        SourceFolder        = (New-Item 'TestDrive:/A' -ItemType Directory).FullName 
        DestinationFolder   = (New-Item 'TestDrive:/B' -ItemType Directory).FullName 
        DestinationFileName = 'copiedFile'
        FileExtension       = 'csv'
        OverWrite           = $true
        MailTo              = @('bob@contoso.com')
        LogFolder           = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'Action', 'SourceFolder', 'DestinationFolder' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'throw a terminating error when' {
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xx:://x'

        { .$testScript @testNewParams } | Should -Throw "Failed creating the log folder 'xx:://x'*"
    }
    It 'the source folder cannot be found' {
        $testNewParams = $testParams.clone()
        $testNewParams.SourceFolder = 'TestDrive:/x'

        { .$testScript @testNewParams } | Should -Throw "Failed checking the source folder '*/x'*"
    } 
}
Describe 'copy only the latest file found in the source folder' {
    BeforeAll {
        $testSourceFiles = @(
            '1.txt', '2.txt', '3.txt', 
            'fruitKiwi.zip', 'fruitApple.zip', 'noFruit.zip',
            '1.csv', '2.csv', 
            '1.xlsx'
        ) | ForEach-Object {
            Start-Sleep -Milliseconds 1
            (New-Item (Join-Path $testParams.SourceFolder $_) -ItemType File).FullName
        }
    }
    BeforeEach {
        Get-ChildItem $testParams.DestinationFolder | Remove-Item
    }
    It 'regardless its extension' {
        $testNewParams = $testParams.clone()
        $testNewParams.Remove('FileExtension')
        $testNewParams.Remove('DestinationFileName')
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be '1.xlsx'
    }
    It 'regardless its extension with a new name' {
        $testNewParams = $testParams.clone()
        $testNewParams.Remove('FileExtension')
        $testNewParams.DestinationFileName = 'A'
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'A.xlsx'
    }
    It 'for a specific extension' {
        $testNewParams = $testParams.clone()
        $testNewParams.FileExtension = '.csv'
        $testNewParams.Remove('DestinationFileName')
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be '2.csv'
    }
    It 'for a specific extension with a new name' {
        $testNewParams = $testParams.clone()
        $testNewParams.FileExtension = '.csv'
        $testNewParams.DestinationFileName = 'A'
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'A.csv'
    }
    It 'that begins with a specific string' {
        $testNewParams = $testParams.clone()
        $testNewParams.Remove('DestinationFileName')
        $testNewParams.FileExtension = '.zip'
        $testNewParams.FileNameStartsWith = 'fruit'
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'fruitApple.zip'
    }
    It 'copy nothing when no match is found' {
        $testNewParams = $testParams.clone()
        $testNewParams.Remove('DestinationFileName')
        $testNewParams.FileExtension = '.zip'
        $testNewParams.FileNameStartsWith = 'notFound'
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        $actual | Should -BeNullOrEmpty
    }
    It "and remove the source file when action is 'Move'" {
        $testNewParams = $testParams.clone()
        $testNewParams.FileExtension = '.csv'
        $testNewParams.DestinationFileName = 'A'
        $testNewParams.Action = 'Move'
        . $testScript @testNewParams
        
        $testSourceFiles = @(
            '1.txt', '2.txt', '3.txt', 
            '1.csv', # '2.csv', 
            '1.xlsx'
        ) | ForEach-Object {
            Join-Path $testParams.SourceFolder $_ | Should -Exist
        }
        Join-Path $testParams.SourceFolder '2.csv' | Should -Not -Exist
    } 
}
Describe 'when the file name already exists in the destination folder' {
    BeforeAll {
        $testSourceFile = New-Item (
            Join-Path $testParams.SourceFolder '1.txt') -ItemType File
        'A' | Out-File -LiteralPath  $testSourceFile
    }
    BeforeEach {
        $testDestinationFile = New-Item (
            Join-Path $testParams.DestinationFolder '1.txt') -ItemType File -Force
        'B' | Out-File -LiteralPath  $testDestinationFile
    }
    It 'it is over written when OverWrite is true' {
        $testNewParams = $testParams.clone()
        $testNewParams.OverWrite = $true
        $testNewParams.Remove('FileExtension')
        $testNewParams.Remove('DestinationFileName')
        . $testScript @testNewParams
        $actual = Get-ChildItem $testParams.DestinationFolder
        Get-Content -Path $actual.FullName | Should -BeExactly 'A'
    } 
    It 'it is not over written when OverWrite is false and an error is thrown' {
        $testNewParams = $testParams.clone()
        $testNewParams.OverWrite = $false
        $testNewParams.Remove('FileExtension')
        $testNewParams.Remove('DestinationFileName')
        { . $testScript @testNewParams } | 
        Should -Throw "The file '$($testDestinationFile.FullName)' already exists in the destination folder, use 'OverWrite'*"
        $actual = Get-ChildItem $testParams.DestinationFolder
        Get-Content -Path $actual.FullName | Should -BeExactly 'B'
    }
}
Describe 'send a summary mail' {
    It 'to users and the admin when MailTo is used' {
        New-Item (Join-Path $testParams.SourceFolder '1.csv') -ItemType File

        $testNewParams = $testParams.clone()
        $testNewParams.FileExtension = '.csv'
        $testNewParams.DestinationFileName = 'A'
        $testNewParams.Action = 'Move'
        $testNewParams.MailTo = 'bob@contoso.com'
        . $testScript @testNewParams
    
        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq 'File moved') -and
            ($Message -like "*<b>Move</b> the most recently edited file with <b>extension '.csv'</b> from the <a href=`"$($testNewParams.SourceFolder)`">source folder</a> to the <a href=`"$($testNewParams.DestinationFolder)`">destination folder</a> and <b>over write the destination file</b> when it exists already.*
        *<th>Destination file</th>*
        *<td><a href=`"$($testNewParams.DestinationFolder + '\A.csv')`">$($testNewParams.DestinationFolder + '\A.csv')</a></td>*
        *<th>Source file</th>*
        *<td><a href=`"$($testNewParams.SourceFolder + '\1.csv')`">1.csv</a></td>*
        *<th>Source file LastWriteTime</th>*"
            )
        }
    }
    It 'only to the admin when MailTo is not used' {
        New-Item (Join-Path $testParams.SourceFolder '1.csv') -ItemType File

        $testNewParams = $testParams.clone()
        $testNewParams.FileExtension = '.csv'
        $testNewParams.DestinationFileName = 'A'
        $testNewParams.Action = 'Move'
        $testNewParams.Remove('MailTo')
        . $testScript @testNewParams
    
        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq $ScriptAdmin) -and
            (-not $Bcc) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq 'File moved') -and
            ($Message -like "*<b>Move</b> the most recently edited file with <b>extension '.csv'</b> from the <a href=`"$($testNewParams.SourceFolder)`">source folder</a> to the <a href=`"$($testNewParams.DestinationFolder)`">destination folder</a> and <b>over write the destination file</b> when it exists already.*
        *<th>Destination file</th>*
        *<td><a href=`"$($testNewParams.DestinationFolder + '\A.csv')`">$($testNewParams.DestinationFolder + '\A.csv')</a></td>*
        *<th>Source file</th>*
        *<td><a href=`"$($testNewParams.SourceFolder + '\1.csv')`">1.csv</a></td>*
        *<th>Source file LastWriteTime</th>*"
            )
        }
    }
}