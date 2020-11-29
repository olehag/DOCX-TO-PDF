# if adminaccess is needed. uncomment line below.
#if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

$ErrorActionPreference = "continue"

#Fill in location.
$docxPATH = "C:\New folder"
$pdfPATH = "C:\New folder\pdf"

#Test Output folder.
if (!(Test-Path $pdfPATH))
{
    #Creates output folder.
    mkdir $pdfPATH
}


#Opens DOCX and saves as PDF.
$Word = New-Object -ComObject "Word.Application"
Get-ChildItem -Path $docxPATH -File *.docx | ForEach-Object {
    $NewName = $_.FullName -replace 'docx','pdf'
    ($Word.Documents.Open($_.FullName)).SaveAs([ref]$NewName,[ref]17) 
    $Word.Application.ActiveDocument.Close()
} 

Get-ChildItem -Path $docxPATH -File *.pdf | Move-item -Destination $pdfPATH -Force
