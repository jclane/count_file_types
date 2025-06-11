param (
    [string]$path
)

$foldersExcluded = @("\AppData", "\ProgramFiles (x86)")

$countByExtension = @{}
$countByPerceivedType = @{}

$shell = New-Object -ComObject shell.application

Get-ChildItem -Path $path -Recurse -Force -File -ErrorAction "SilentlyContinue" | ForEach-Object {
    if ($_.PSIsContainer) { return } # Skip directories
    
    # The following skips excluded folders listed $folderExcluded
    foreach ($folder in $foldersExcluded) {
        if ($_.DirectoryName.Contains($folder)) { return }
    }

    $ext = $_.Extension -replace "^\.","" # Remove the dot from extension
    $countByExtension[$ext] = ($countByExtension[$ext] + 1)
    
    $folderObj = $shell.namespace($_.Directory.FullName)
    $fileObj = $folderObj.Items().Item($_.Name)
    for ($i = 0; $i -le 294; $i++) {     #294 ...i don't know how high that should be to get all meta data
        $name = $folderObj.getDetailsOf($null, $i)
        if ($name) {
			$value = $folderObj.getDetailsOf($fileObj, $i)
            if ($value) {
                if ($name -eq "Perceived type") {
                    $countByPerceivedType[$value] = ($countByPerceivedType[$value] + 1)
                }
                Add-Member -InputObject $fileObj "NoteProperty" -Name $name -Value $value -Force
            }
        }
    }
}

Write-Output "`nBY PERCEIVED TYPE:"
Write-Output $countByPerceivedType | Format-Table @{Label="Perceived Type"; Expression={$_.Name}}, @{Label="Count"; Expression={$_.Value}} -Autosize

Write-Output "`nBY EXTENSION:"
Write-Output $countByExtension | Format-Table @{Label="Extension"; Expression={$_.Name}}, @{Label="Count"; Expression={$_.Value}} -Autosize
