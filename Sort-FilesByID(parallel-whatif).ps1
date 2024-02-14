$sourcefolder = Read-Host -Prompt 'Enter Source Folder'
$destfolder = Read-Host -Prompt 'Enter Destination Folder'
$destfilename = Read-Host -Prompt 'Enter Filename (e.g. letter1.docx)'

$destDirectories = Get-ChildItem -Path "$destfolder" -Recurse -Directory

$destHashtable = @{}
foreach ($dir in $destDirectories) {
    $dirID = $dir.Name -replace '.*?(\d+).*', '$1'
    $destHashtable[$dirID] = $dir.FullName
}

$files = Get-ChildItem -Path "$sourcefolder" -Recurse -File

$jobs = @()

foreach ($file in $files) {
    $fileName = $file.Name
    $fileID = $fileName -replace '.*?(\d+).*', '$1'

    $matchingFolder = $destHashtable[$fileID]

    if ($matchingFolder) {
        $destinationPath = Join-Path $matchingFolder $destfilename
        $job = Start-Job -ScriptBlock {
            param ($file, $destinationPath)
            Move-Item -Path $file.FullName -Destination $destinationPath -WhatIf
            Write-Host -ForegroundColor Green "$($file.FullName) successfully moved to $destinationPath"
            Write-Host ""
        } -ArgumentList $file, $destinationPath
        $jobs += $job
    }
    else {
        Write-Host -ForegroundColor Red "Could not move $($file.FullName) to matching folder"
        Write-Host ""
    }
}

Wait-Job -Job $jobs | Out-Null

$jobs | Receive-Job

$jobs | Remove-Job

Read-Host -Prompt 'Done... Press Enter To Exit...'
