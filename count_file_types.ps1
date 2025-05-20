param (
    [string]$path
)

$foldersExcluded = @("\AppData", "\ProgramFiles (x86)")

$extensionTypes = @{
    "Images" = @("jpg", "png", "bmp", "gif", "tiff")
    "Documents" = @("docx", "doc", "txt", "pdf", "rtf")
    "Spreadsheets" = @("xls", "xlsx")
    "Audio" = @("mp3", "wav", "flac", "aac")
    "Videos" = @("mp4", "avi", "mov", "wmv", "mkv")
}

$counts = @{}

Get-ChildItem -Path $path -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.PSIsContainer) { return } # Skip directories
    foreach ($folder in $foldersExcluded) {
        if ($_.DirectoryName.Contains($folder)) { return }
    }
    $ext = $_.Extension -replace "^\.","" # Remove the dot from extension
    foreach ($extensionType in $extensionTypes.Keys) {
        if ($ext -in $extensionTypes[$extensionType]) {
            $counts[$extensionType] = ($counts[$extensionType] + 1)
            #ii $_.FullName
        }
    }
}

Write-Output $counts
Out-File -InputObject $counts -FilePath ".\file_counts.txt"
