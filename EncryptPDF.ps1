# pdftk must be installed and added to PATH environment variables
# It can be downloaded and installed at https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/
# After installation, add its \bin directory to PATH environment variables
# Typically it should be at "C:\Program Files (x86)\PDFtk\bin"
if (!(Get-Command "pdftk" -errorAction SilentlyContinue)) {
    Write-Warning "Please install pdftk and try again."
    exit
}

Write-Warning "Enter an owner password. Owner password must be different from PDF passwords in Passwords.csv"
$owner_pw = Read-Host "Password"
if(-not $owner_pw) {
    Write-Warning "You are not setting an owner password. People can remove password in protected PDF files when owner password is not set."
}

# Get all PDF files in current directory    
$path = Get-Location
$pdfFiles = Get-ChildItem -Path ($path.Path + "\*") -include *.pdf

# Import the password list in CSV file
$passwords = Import-CSV .\Passwords.csv

# Create output directory if not exists
$outputPath = Join-Path -Path $path -ChildPath "\output"
New-Item -ItemType Directory -Force -Path $outputPath > $null

foreach ($file in $pdfFiles) {
    $filepath = Join-Path -Path $outputPath -ChildPath $file.Name
    $matchPassword = $passwords | Where-Object -Property Name -eq $file.BaseName | Select-Object -First 1
    
    if ($null -eq $matchPassword -or $null -eq $matchPassword.Password) {
        Write-Warning "Found no matching password for file ""$($file.Name)"" in Paswords.csv"
    }
    else {
        $cmd = "pdftk ""$file"" output ""$filePath"" user_pw $($matchPassword.Password)"
        if ($owner_pw) {
            $cmd = $cmd + " owner_pw $owner_pw"
        }
        "Writing to output ""$filePath"""
        Invoke-Expression $cmd
    }
}
