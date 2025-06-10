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

$countByType = @{}
$countByExtension = @{}
$blah = @{}

$shell = New-Object -com shell.application

Get-ChildItem -Path $path -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.PSIsContainer) { return } # Skip directories
    
    # The following skips excluded folders listed $folderExcluded
    foreach ($folder in $foldersExcluded) {
        if ($_.DirectoryName.Contains($folder)) { return }
    }

    $rootFolder = $shell.namespace($_.Directory.FullName)
    $fileObj = $rootFolder.Items().Item($_.Name)
    for ($i = 0; $i -le 294; $i++) {     #294 ...i don't know how high that should be to get all meta data
        $name = $rootFolder.getDetailsOf($null, $i)
        if ($name) {
            $value = $rootFolder.getDetailsOf($fileObj, $i)
            if ($value) {
                echo $name
                if ($name -eq "Perceived type") {
                    $blah[$value] = ($blah[$value] + 1)
                }
                Add-Member -InputObject $fileObj "NoteProperty" -Name $name -Value $value -Force
            }
        }
    }

    #$ext = $_.Extension -replace "^\.","" # Remove the dot from extension
    #foreach ($extensionType in $extensionTypes.Keys) {
    #    $countByExtension[$ext] = ($countByExtension[$ext] + 1)
    #    if ($ext -in $extensionTypes[$extensionType]) {
    #        $countByType[$extensionType] = ($countByType[$extensionType] + 1)
    #        #ii $_.FullName
    #    }
    #}
}

#Write-Output "`nBY TYPE:"
#Write-Output $countByType | Format-Table 

Write-Output "`nBY PERCEIVED TYPE:"
Write-Output $blah | Format-Table 

#Write-Output "`nBY EXTENSION:"
#Write-Output $countByExtension | Format-Table

#Out-File -InputObject $countByType -FilePath ".\file_counts.txt"