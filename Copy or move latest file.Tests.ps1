#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $testParams.ScriptAdmin) -and
        ($Priority -eq 'High') -and
        ($Subject -eq 'FAILURE')
    }

    $testInputFile = @{
        Action              = 'Copy'
        SourceFolder        = (New-Item 'TestDrive:/A' -ItemType Directory).FullName
        DestinationFolder   = (New-Item 'TestDrive:/B' -ItemType Directory).FullName
        DestinationFileName = 'copiedFile'
        SearchFor           = @{
            FileExtension      = 'csv'
            FileNameStartsWith = $null
        }
        OverWrite           = $true
        MailTo              = @('bob@contoso.com')
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.Clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.Clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'Action', 'SourceFolder', 'DestinationFolder', 'MailTo'
            ) {
                $testNewInputFile = $testInputFile.Clone()
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property '$_' not found*")
                }
            }
            It 'OverWrite is not a boolean value' {
                $testNewInputFile = $testInputFile.Clone()
                $testNewInputFile.OverWrite = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Overwrite' is not a boolean value*")
                }
            }
        }
    }
    Context 'the folder is not found' {
        It 'SourceFolder' {
            $testNewInputFile = $testInputFile.Clone()
            $testNewInputFile.SourceFolder = 'c:\upDoesNotExist.ps1'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Source folder 'c:\upDoesNotExist.ps1' not found*")
            }
        }
        It 'DestinationFolder' {
            $testNewInputFile = $testInputFile.Clone()
            $testNewInputFile.DestinationFolder = 'c:\downDoesNotExist.ps1'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Destination folder 'c:\downDoesNotExist.ps1' not found*")
            }
        }
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
            (New-Item (Join-Path $testInputFile.SourceFolder $_) -ItemType File).FullName
        }
    }
    BeforeEach {
        Get-ChildItem $testInputFile.DestinationFolder | Remove-Item
    }
    It 'regardless its extension' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = $null
        $testNewInputFile.Remove('DestinationFileName')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be '1.xlsx'
    }
    It 'regardless its extension with a new name' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = $null
        $testNewInputFile.DestinationFileName = 'A'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'A.xlsx'
    }
    It 'for a specific extension' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = '.csv'
        $testNewInputFile.Remove('DestinationFileName')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be '2.csv'
    }
    It 'for a specific extension with a new name' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = '.csv'
        $testNewInputFile.DestinationFileName = 'A'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'A.csv'
    }
    It 'that begins with a specific string' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.Remove('DestinationFileName')
        $testNewInputFile.SearchFor.FileExtension = '.zip'
        $testNewInputFile.SearchFor.FileNameStartsWith = 'fruit'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -HaveCount 1
        $actual.Name | Should -Be 'fruitApple.zip'
    }
    It 'copy nothing when no match is found' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.Remove('DestinationFileName')
        $testNewInputFile.SearchFor.FileExtension = '.zip'
        $testNewInputFile.SearchFor.FileNameStartsWith = 'notFound'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        $actual | Should -BeNullOrEmpty
    }
    It "and remove the source file when action is 'Move'" {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = '.csv'
        $testNewInputFile.SearchFor.FileNameStartsWith = $null
        $testNewInputFile.DestinationFileName = 'A'
        $testNewInputFile.Action = 'Move'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        @(
            '1.txt', '2.txt', '3.txt',
            '1.csv', # '2.csv',
            '1.xlsx'
        ) | ForEach-Object {
            Join-Path $testNewInputFile.SourceFolder $_ | Should -Exist
        }
        Join-Path $testNewInputFile.SourceFolder '2.csv' | Should -Not -Exist
    }
}
Describe 'when the file name already exists in the destination folder' {
    BeforeAll {
        $testSourceFile = New-Item (
            Join-Path $testInputFile.SourceFolder '1.txt') -ItemType File
        'A' | Out-File -LiteralPath  $testSourceFile
    }
    BeforeEach {
        $testDestinationFile = New-Item (
            Join-Path $testInputFile.DestinationFolder '1.txt') -ItemType File -Force
        'B' | Out-File -LiteralPath  $testDestinationFile
    }
    It 'it is over written when OverWrite is true' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.OverWrite = $true
        $testNewInputFile.SearchFor.Remove('FileExtension')
        $testNewInputFile.Remove('DestinationFileName')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        Get-Content -Path $actual.FullName | Should -BeExactly 'A'
    }
    It 'it is not overwritten when OverWrite is false and an error is thrown' {
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.OverWrite = $false
        $testNewInputFile.SearchFor.Remove('FileExtension')
        $testNewInputFile.Remove('DestinationFileName')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*The file '$($testDestinationFile.FullName)' already exists in the destination folder, use 'OverWrite'*")
        }

        $actual = Get-ChildItem $testInputFile.DestinationFolder
        Get-Content -Path $actual.FullName | Should -BeExactly 'B'
    }
}
Describe 'send a summary mail when' {
    BeforeAll {
        New-Item (Join-Path $testInputFile.SourceFolder '1.csv') -ItemType File
        $testNewInputFile = $testInputFile.clone()
        $testNewInputFile.SearchFor.FileExtension = '.csv'
        $testNewInputFile.DestinationFileName = 'A'
        $testNewInputFile.MailTo = 'bob@contoso.com'
    }
    It 'a file is copied' {
        $testNewInputFile.Action = 'Copy'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq 'File copied') -and
            ($Message -like "*<b>Copy</b> the most recently edited file with <b>extension '.csv'</b> from the <a href=`"$($testInputFile.SourceFolder)`">source folder</a> to the <a href=`"$($testInputFile.DestinationFolder)`">destination folder</a> and <b>overwrite</b> the destination file when it exists already.*
            *<th>Source</th>*
            *<td>*
            *<a href=`"$($testInputFile.SourceFolder + '\1.csv')`">1.csv</a><br>*
            *LastWriteTime: *</td>*
            *<th>Destination</th>*
            *<td><a href=`"$($testInputFile.DestinationFolder + '\A.csv')`">A.csv</a></td>*"
            )
        }
    }
    It 'a file is moved' {
        $testNewInputFile.Action = 'Move'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        . $testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq 'File moved') -and
            ($Message -like "*<b>Move</b> the most recently edited file with <b>extension '.csv'</b> from the <a href=`"$($testInputFile.SourceFolder)`">source folder</a> to the <a href=`"$($testInputFile.DestinationFolder)`">destination folder</a> and <b>overwrite</b> the destination file when it exists already.*
            *<th>Source</th>*
            *<td>*
            *<a href=`"$($testInputFile.SourceFolder + '\1.csv')`">1.csv</a><br>*
            *LastWriteTime: *</td>*
            *<th>Destination</th>*
            *<td><a href=`"$($testInputFile.DestinationFolder + '\A.csv')`">A.csv</a></td>*"
            )
        }
    }
    It 'no file is found in the source folder' {
        $testNewInputFile.Action = 'Move'

        . $testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'No file moved') -and
            ($Message -like "*<b>Move</b> the most recently edited file with <b>extension '.csv'</b> from the <a href=`"$($testInputFile.SourceFolder)`">source folder</a> to the <a href=`"$($testInputFile.DestinationFolder)`">destination folder</a> and <b>overwrite</b> the destination file when it exists already.*No file found in the source folder matching the search criteria*"
            )
        }
    }
}