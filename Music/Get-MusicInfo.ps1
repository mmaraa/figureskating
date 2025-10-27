<#
.SYNOPSIS
    List all MP3 files in a folder (and optionally subfolders), displaying their length and title, grouped by folder.

.DESCRIPTION
    This script scans the given folder for MP3 files and prints tables for "Vapaaohjelma" and "Lyhytohjelma" files, grouped by their parent folder.

.PARAMETER Path
    The root folder to search for MP3 files.

.PARAMETER Recurse
    If specified, search recursively in all subfolders.

.EXAMPLE
    .\Get-MusicInfo.ps1 -Path "C:\Competitions\Music" -Recurse

    .\Get-MusicInfo.ps1 -Path ".\Music\SM-Seniorit\"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [switch]$Recurse
)

if ($Recurse) {
    $mp3Files = Get-ChildItem -Path $Path -Filter *.mp3 -Recurse | ForEach-Object {
        $shell = New-Object -ComObject Shell.Application
        $folderObj = $shell.Namespace($_.DirectoryName)
        $fileObj = $folderObj.ParseName($_.Name)

        # Extract file length (duration)
        $duration = $folderObj.GetDetailsOf($fileObj, 27)  # 27 is usually the duration column for MP3s

        # Find the index of the Title column for this folder (if available)
        $titleIndex = (0..300 | Where-Object { $folderObj.GetDetailsOf($null, $_) -eq 'Title' }) | Select-Object -First 1
        if ($null -ne $titleIndex) {
            $title = $folderObj.GetDetailsOf($fileObj, $titleIndex)
        }
        else {
            # Fallback: try common column header names (localized systems may differ)
            $possibleNames = @('Title', 'Ämne', 'Titel')
            $found = $null
            foreach ($i in 0..300) {
                $hdr = $folderObj.GetDetailsOf($null, $i)
                if ($possibleNames -contains $hdr) { $found = $i; break }
            }
            if ($null -ne $found) { $title = $folderObj.GetDetailsOf($fileObj, $found) } else { $title = '' }
        }

        # Add Folder property for grouping
        [PSCustomObject]@{
            Name     = $_.Name
            Title    = $title
            Length   = $duration
            Folder   = Split-Path $_.DirectoryName -Leaf
        }
    }
} else {
    $mp3Files = Get-ChildItem -Path $Path -Filter *.mp3 | ForEach-Object {
        $shell = New-Object -ComObject Shell.Application
        $folderObj = $shell.Namespace($_.DirectoryName)
        $fileObj = $folderObj.ParseName($_.Name)

        # Extract file length (duration)
        $duration = $folderObj.GetDetailsOf($fileObj, 27)  # 27 is usually the duration column for MP3s

        # Find the index of the Title column for this folder (if available)
        $titleIndex = (0..300 | Where-Object { $folderObj.GetDetailsOf($null, $_) -eq 'Title' }) | Select-Object -First 1
        if ($null -ne $titleIndex) {
            $title = $folderObj.GetDetailsOf($fileObj, $titleIndex)
        }
        else {
            # Fallback: try common column header names (localized systems may differ)
            $possibleNames = @('Title', 'Ämne', 'Titel')
            $found = $null
            foreach ($i in 0..300) {
                $hdr = $folderObj.GetDetailsOf($null, $i)
                if ($possibleNames -contains $hdr) { $found = $i; break }
            }
            if ($null -ne $found) { $title = $folderObj.GetDetailsOf($fileObj, $found) } else { $title = '' }
        }

        [PSCustomObject]@{
            Name     = $_.Name
            Title    = $title
            Length   = $duration
            Folder   = Split-Path $_.DirectoryName -Leaf
        }
    }
}

# Group by Folder and print tables for each
$mp3Files | Group-Object Folder | ForEach-Object {
    $folderName = $_.Name
    $groupFiles = $_.Group

    # Vapaaohjelma
    $vapaa = $groupFiles | Where-Object { $_.Name -like "*Vapaaohjelma*" } | Sort-Object Name
    if ($vapaa.Count -gt 0) {
        Write-Host "`n$folderName - Vapaaohjelma" -ForegroundColor Cyan
        $vapaa | Format-Table Name, Title, Length -AutoSize
    }

    # Lyhytohjelma
    $lyhyt = $groupFiles | Where-Object { $_.Name -like "*Lyhytohjelma*" } | Sort-Object Name
    if ($lyhyt.Count -gt 0) {
        Write-Host "`n$folderName - Lyhytohjelma" -ForegroundColor Yellow
        $lyhyt | Format-Table Name, Title, Length -AutoSize
    }
}
