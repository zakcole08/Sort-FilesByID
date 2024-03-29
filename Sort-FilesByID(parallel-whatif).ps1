Clear-Host
Write-Host -foregroundcolor Blue " _____            _    ______ _ _            ______         ___________ 
/  ___|          | |   |  ___(_) |           | ___ \       |_   _|  _  \
\ `--.   ___  _ __| |_  | |_   _| | ___  ___  | |_/ /_   _    | | | | | |
 `--. \ / _ \| '__| __| |  _| | | |/ _ \/ __| | ___ \ | | |   | | | | | |
/\__/ / (_) | |  | |_  | |   | | |  __/\__ \ | |_/ / |_| |  _| |_| |/ / 
\____/ \___/|_|   \__| \_|   |_|_|\___||___/ \____/ \__, |  \___/|___/  
                                                     __/ |              
                                                    |___/               
"
Write-Host "This script will move files from the 'Source Folder' containing an ID (first number in the filename) to the Directory/Sub-directory with the corresponding ID in the 'Destination Folder'.
----------------------------------------------------------------------------
Source Folder: The folder containing the documents you would like to move
Destination Folder: The target location to move the files into. (Should contain directories/sub-directories with ID numbers)
Filename: [Folder]\[Filename.docx]"
Write-Host ""

$sourcefolder = Read-Host -Prompt 'Enter Source Folder'
$destfolder = Read-Host -Prompt 'Enter Destination Folder'
$destfilename = Read-Host -Prompt 'Enter Filename (e.g. letter1.docx)'

$destDirectories = Get-ChildItem -Path "$destfolder" -Recurse -Directory

$destHashtable = @{}
foreach ($dir in $destDirectories)
{
    $dirID = $dir.Name -replace '.*?(\d+).*', '$1'
    $destHashtable[$dirID] = $dir.FullName
}

$files = Get-ChildItem -Path "$sourcefolder" -Recurse -File

$jobs = @()

foreach ($file in $files)
{
    $fileName = $file.Name
    $fileID = $fileName -replace '.*?(\d+).*', '$1'

    $matchingFolder = $destHashtable[$fileID]

    if ($matchingFolder)
    {
        $destinationPath = Join-Path $matchingFolder $destfilename
        $job = Start-Job -ScriptBlock
        {
            param ($file, $destinationPath)
            Move-Item -Path $file.FullName -Destination $destinationPath -WhatIf
            Write-Host -ForegroundColor Green "$($file.FullName) successfully moved to $destinationPath"
            Write-Host ""
        } -ArgumentList $file, $destinationPath
        $jobs += $job
    }
    else
    {
        Write-Host -ForegroundColor Red "Could not move $($file.FullName) to matching folder"
        Write-Host ""
    }
}

Wait-Job -Job $jobs | Out-Null

$jobs | Receive-Job

$jobs | Remove-Job

Read-Host -Prompt 'Done... Press Enter To Exit...'
