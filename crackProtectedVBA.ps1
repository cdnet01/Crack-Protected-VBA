# variables
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$current_dir = Get-Location
$path = $args[0]
$cracked_output = "CRACKED_$path_$timestamp.xlsm"

# Define ANSI escape codes for colors
$GREEN = [System.ConsoleColor]::Green
$YELLOW = [System.ConsoleColor]::Yellow
$RED = [System.ConsoleColor]::Red
$NC = [System.Console]::ResetColor()

# Check if a filename was provided as a command-line argument
if ($args.Count -eq 0) {
    Write-Host "Usage: $PSCommandPath <filename>"
    exit 1
}

# Check if the file exists
if (-not (Test-Path $path -PathType Leaf)) {
    Write-Host ("[-] File '$path' not found.") -ForegroundColor $RED
    exit 1
}
Write-Host ("[+] Processing file: $path") -ForegroundColor $GREEN

# Copy to staging directory
Write-Host "[+] copying $path to staging directory..." -ForegroundColor $GREEN
$staging_dir = Join-Path $current_dir "staging_$timestamp"
New-Item -ItemType Directory -Path $staging_dir -Force | Out-Null
Copy-Item $path -Destination (Join-Path $staging_dir "tobe-cracked.zip") -Force

# extract VBA file contents
Write-Host "[+] extracting .xlsm file..." -ForegroundColor $GREEN
Expand-Archive -Path (Join-Path $staging_dir "tobe-cracked.zip") -DestinationPath $staging_dir | Out-Null
Remove-Item (Join-Path $staging_dir "tobe-cracked.zip")

# swap DPB hex bit to DPX using PowerShell binary reading and writing
Write-Host "[+] breaking password protection mechanism..." -ForegroundColor $GREEN
$binFile = Join-Path $staging_dir "xl\vbaProject.bin"
$byteArray = Get-Content $binFile -Raw -Encoding Byte
$byteString = $byteArray.ForEach('ToString', 'X') -join ' '
$byteString = $byteString -replace '\b44 50 42\b(.*)', '44 50 58$1'
[byte[]] $newByteArray = -split $byteString -replace '^', '0x'
Set-Content $binFile -Encoding Byte -Value $newByteArray

Write-Host "Binary data written back to $binFile" -ForegroundColor Green

# Repackage VBA contents
Write-Host "[+] Packaging files back into .xlsm" -ForegroundColor $GREEN
Set-Location $staging_dir
Compress-Archive -Path * -DestinationPath (Join-Path $current_dir "cracked.zip") -Force | Out-Null
Move-Item (Join-Path $current_dir "cracked.zip") (Join-Path $current_dir $cracked_output)

# Remove artifacts
Write-Host "[+] Removing staging directory" -ForegroundColor $GREEN
Set-Location $current_dir
Remove-Item $staging_dir -Recurse -Force

Write-Host "[+] Done! Cracked file is $($current_dir)\$cracked_output" -ForegroundColor $GREEN
