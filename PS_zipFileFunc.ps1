function Write-ZipUsing7Zip([string]$FilesToZip, [string]$ZipOutputFilePath, [string]$Password, [ValidateSet('7z','zip','gzip','bzip2','tar','iso','udf')][string]$CompressionType = 'zip', [switch]$HideWindow)
{

<#*******************************************************************************
 Purpose: zip file using 7 zip

 Dependency: 7zip installed

 Reference: https://blog.danskingdom.com/powershell-function-to-create-a-password-protected-zip-file/
  
    
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 20-Mar-2021    William Hu       Re-use
 09-May-2013    Daniel Schroeder Initial
 *******************************************************************************#> 

	# Look for the 7zip executable.
	$pathTo32Bit7Zip = "C:\Program Files (x86)\7-Zip\7z.exe"
	$pathTo64Bit7Zip = "C:\Program Files\7-Zip\7z.exe"
	$THIS_SCRIPTS_DIRECTORY = Split-Path $script:MyInvocation.MyCommand.Path
	$pathToStandAloneExe = Join-Path $THIS_SCRIPTS_DIRECTORY "7za.exe"
	if (Test-Path $pathTo64Bit7Zip) { $pathTo7ZipExe = $pathTo64Bit7Zip } 
	elseif (Test-Path $pathTo32Bit7Zip) { $pathTo7ZipExe = $pathTo32Bit7Zip }
	elseif (Test-Path $pathToStandAloneExe) { $pathTo7ZipExe = $pathToStandAloneExe }
	else { throw "Could not find the 7-zip executable." }
	
	# Delete the destination zip file if it already exists (i.e. overwrite it).
	if (Test-Path $ZipOutputFilePath) { Remove-Item $ZipOutputFilePath -Force }
	
	$windowStyle = "Normal"
	if ($HideWindow) { $windowStyle = "Hidden" }
	
	# Create the arguments to use to zip up the files.
	# Command-line argument syntax can be found at: http://www.dotnetperls.com/7-zip-examples
	$arguments = "a -t$CompressionType ""$ZipOutputFilePath"" ""$FilesToZip"" -mx9"
	if (!([string]::IsNullOrEmpty($Password))) { $arguments += " -p$Password" }
	
	# Zip up the files.
	$p = Start-Process $pathTo7ZipExe -ArgumentList $arguments -Wait -PassThru -WindowStyle $windowStyle

	# If the files were not zipped successfully.
	if (!(($p.HasExited -eq $true) -and ($p.ExitCode -eq 0))) 
	{
		throw "There was a problem creating the zip file '$ZipFilePath'."
	}
}