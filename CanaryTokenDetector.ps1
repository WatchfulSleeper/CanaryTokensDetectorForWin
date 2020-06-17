#
# Name: CanaryTokenDetector.ps1
# Author: Swarley
#
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

$docxPath = New-Object System.Windows.Forms.OpenFileDialog
$docxPath.filter = "Docx (*.docx)| *.docx"
[void]$docxPath.ShowDialog()
$docxPath = $docxPath.FileName

Copy-Item -Path $docxPath -Destination "Work.zip" -ErrorAction SilentlyContinue
Expand-Archive -Path "Work.zip" -DestinationPath "Work" -ErrorAction SilentlyContinue

$DocxSubFiles = Get-ChildItem -Path "Work" -Filter *.xml -Recurse -ErrorAction SilentlyContinue -Force

$canaryDetected = $false
$DocxSubFiles | foreach {
	IF ($_.FullName -match "footer" -or $_.FullName -match "header") {
		IF (Select-String -Path $_.FullName -Pattern 'canarytoken') {
			$canaryDetected = $true}
	}
}

Remove-Item -Path Work.zip -Force
Remove-Item -Path Work -Recurse -Force

IF ($canaryDetected -eq $true) {
		[System.Windows.MessageBox]::Show("CanaryToken detected in $docxPath",'Results',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)}
ELSE {
		[System.Windows.MessageBox]::Show("No CanaryToken found in $docxPath",'Results',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
	}