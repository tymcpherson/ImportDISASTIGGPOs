<#
##################################################################################################################
#
# Microsoft Premier Field Engineering
# ty.mcpherson@microsoft.com
# Import_DISA_GPOs.ps1
# 
# Purpose:
#  This is script is used to automate the process of creating and importing the STIG GPOs
#  that DISA provides.  This script takes ~30 minutes to run.  Hopefully this saves you
#  time in the long run!!!
#
# Usage:
#  1) Download the latest DISA GPO zip file from: https://iase.disa.mil/stigs/gpo/Pages/index.aspx
#  2) Run Powershell script and select downloaded .zip file
#  3) Do something else for about 30 minutes
#  4) Profit!
#
# Script Process:
#  1) Targets the .zip file that was downloaded
#  2) Extracts the contents of the .zip file to %TEMP%
#  3) Creates a migration table
#  4) Creates and imports non-Office GPOs
#  5) Creats the combined Office user and computer GPOs per version
#  6) Creats a temporary random Office Product GPO, then imports that products settings into the combined Office
#     user or computer version
#  7) Removes the temporary random Office Product GPO
#  8) Adds the Office product STIG versioning to the description
#  9) Clean up the migration table file, and extracted .zip contents
#
#  Note:  There are sleep statements within this script that make the processing longer due to race conditions
#         between creating the GPO, adding settings/modifying attributes
#         
#         Also it's been reported that various Security products may have to be disabled during execution
#
#  ChangeLog:
#   *January 20, 2019 - Initial Creation
#   *January 21, 2019 - Added check for file after File/Open Dialog
#   *February 11, 2019 - Using a working directory in case the script is halted prior to completion, then re-ran
#
#
# Microsoft Disclaimer for custom scripts
# ================================================================================================================
# The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts
# are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, 
# without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire
# risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event
# shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be
# liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business
# interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to 
# use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
# ================================================================================================================
#
#>Clear-Host$DateStamp = Get-Date -DisplayHint Date
Import-Module ActiveDirectory
Import-Module GroupPolicy
$PDCE = (Get-ADDOmain).PDCEmulator

Write-Host "##########################" -ForegroundColor Cyan
Write-Host "Beginning the hydration of DISA STIG GPOs  ..." -ForegroundColor Cyan## File Select Dialog box#Add-Type -AssemblyName System.Windows.Forms$ZipFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{    Filter = 'Zip Files (*.zip)|*.zip'    initialDirectory = "$env:USERPROFILE\Downloads"    }[void]$ZipFile.ShowDialog()
#
# Check if user selected a file, or cancelled the File/Open Diaglog
#
If (!$ZipFile.FileName){
        Write-Host `n
        write-host `t"No .zip file selected !!!" -ForegroundColor Red
        Write-Host `t"Terminated" -ForegroundColor Red
        Write-Host `n
        Write-Host "##########################" -ForegroundColor Cyan
        Exit
    }

## Uncompress .ZIP file to Temp variable#
$GPObackupsZipFile = $ZipFile.FileName
$GPOBackupsSafeZipFile = $ZipFile.SafeFileName
$WorkingDir = Get-Random -Minimum 1 -Maximum 9999$GPOExtractDestination = "$env:TEMP\$WorkingDir"Add-Type -assembly "system.io.compression.filesystem"[io.compression.zipfile]::ExtractToDirectory($GPObackupsZipFile, $GPOExtractDestination)
Write-Host `t"Extracted GPOs from $GPOBackupsSafeZipFile ..." -ForegroundColor Yellow

#
# Get Package Version
#
$GetPackageVersion = Get-ChildItem ($GPOExtractDestination) -filter *DISA* -Directory | select name
$PackageVersion = $GetPackageVersion.Name.TrimEnd(" DISA STIG GPO Package")
$GPOBackups = $GPOExtractDestination+"\"+(Get-ChildItem ($GPOExtractDestination) -filter *DISA* -Directory)

#
#Create Migration Table File
#
Write-Host `t"Creating an import migration table ..." -ForegroundColor Yellow

Function Write-This($data, $log)
{
    try
    {
        Add-Content -Path $log -Value $data -ErrorAction Stop
    }
    catch
    {
        write-host $_.Exception.Message
    }
}
    
$Migtable = "$GPOExtractDestination\migtable.migtable"
Write-This "<?xml version=`"1.0`" encoding=`"utf-16`"?>" $Migtable
Write-This "<MigrationTable xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns=`"http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable`">" $Migtable
Write-This "<Mapping>" $Migtable
Write-This "<Type>Unknown</Type>" $Migtable
Write-This "<Source>ADD YOUR DOMAIN ADMINS</Source>" $Migtable
Write-This "<Destination>Domain Admins</Destination>" $Migtable
Write-This "</Mapping>" $Migtable
Write-This "<Mapping>" $Migtable
Write-This "<Type>Unknown</Type>" $Migtable
Write-This "<Source>ADD YOUR ENTERPRISE ADMINS</Source>" $Migtable
Write-This "<Destination>Enterprise Admins</Destination>" $Migtable
Write-This "</Mapping>" $Migtable
Write-This "</MigrationTable>" $Migtable


#
#Create and Import non-Office GPOs
#
Write-Host `t"Creating the non Office GPOs ..." -ForegroundColor Yellow
#Doing non-Office GPOs first
$NonOfficeGPODirectories = Get-ChildItem ($GPObackups) -exclude "ADMX Templates","*Office*" -Directory

ForEach ($GPODirectory in $NonOfficeGPODirectories) {
    $GPOSafeName = $GPODirectory.Name
    # Testing the path due to inconsitencies of the DISA Product
    if (Test-Path $GPObackups\$GPOSafeName\GPO -PathType Container) {$GPODIR = "GPO"}
    if (Test-Path $GPObackups\$GPOSafeName\GPOs -PathType Container) {$GPODIR = "GPOs"}
    $GetGPOGUIDDirs = Get-ChildItem ("$GPObackups"+"\"+$GPOSafeName+"\"+$GPODIR) -Directory                ForEach ($GPOGUIDDIR in $GetGPOGUIDDirs){            [xml]$manifest = Get-Content "$GPOBackups\$GPOSafeName\$GPODIR\$GPOGUIDDIR\bkupInfo.xml"
                ForEach ($gpBackup in $manifest.BackupInst) {                    $GPOName = $gpBackup.GPODisplayName.InnerText                    $GPOBackupID = $gpBackup.ID.InnerText                    Write-Host `t`t"-- $GPOName" -ForegroundColor Green                    Start-Sleep 3                    Import-GPO -BackupId $GPOBackupID -Path $GPObackups\$GPOSafeName\$GPODIR -TargetName $GPOName -Server $PDCE -MigrationTable $Migtable -CreateIfNeeded | Out-Null                 }                             # Add original GPO name to description             $GPODescription = New-Object System.Collections.Generic.List[System.Object]             $GPODescription.Add($GPOName )             $GPODescription.Add("--"+"Added By: $env:USERNAME")             (Get-GPO -Name "$GPOName" -Server $PDCE).Description="$GPODescription"        }   
}

#
#Create, Merge, and Import Office GPOs
#
$OfficeGPODirectories = Get-ChildItem ($GPObackups) -filter "*Office*" -Directory

ForEach ($OfficeGPODirectory in $OfficeGPODirectories) {
    $OfficeGPODirectorySafeName = $OfficeGPODirectory.Name
        if ($OfficeGPODirectory -match 2013) {
           $OfficeVersion = 2013
        }
        else{
           $OfficeVersion = 2016
        }
            
    $NewOfficeUserGPO = "DoD Office $OfficeVersion STIG User $PackageVersion"    $NewOfficeComputerGPO = "DoD Office $OfficeVersion STIG Computer $PackageVersion"

####################################
#
# Functions
#   Credit to:
#   Ashley McGlone
#   Microsoft Premier Field Engineer
#   http://aka.ms/GoateePFE
#   May 2015
#
####################################


Function Copy-GPRegistryValue {


[CmdletBinding()]
Param(
    [Parameter()]
    [ValidateSet('All','User','Computer')]
    [String]
    $Mode = 'All',
    [Parameter()]
    [String[]]
    $SourceGPO,
    [Parameter()]
    [String]
    $DestinationGPO
)
    $ErrorActionPreference = 'Continue'

    Switch ($Mode) {
        'All'      {$rootPaths = "HKCU\Software","HKLM\System","HKLM\Software"; break}
        'User'     {$rootPaths = "HKCU\Software"                              ; break}
        'Computer' {$rootPaths = "HKLM\System","HKLM\Software"                ; break}
    }
    
    If (Get-GPO -Name $DestinationGPO -Server $PDCE -ErrorAction SilentlyContinue) {
        #Write-Verbose "DESTINATION GPO EXISTS [$DestinationGPO]"
    } Else {
        #Write-Verbose "CREATING DESTINATION GPO [$DestinationGPO]"
        New-GPO -Name $DestinationGPO -Server $PDCE|Out-Null
    }

    $ProgressCounter = 0
    $ProgressTotal   = @($SourceGPO).Count   # Syntax for PSv2 compatibility
    ForEach ($SourceGPOSingle in $SourceGPO) {

        #Write-Progress -PercentComplete ($ProgressCounter / $ProgressTotal * 100) -Activity "Copying GPO settings to: $DestinationGPO" -Status "From: $SourceGPOSingle"

        If (Get-GPO -Name $SourceGPOSingle -Server $PDCE -ErrorAction SilentlyContinue) {

            #Write-Verbose "SOURCE GPO EXISTS [$SourceGPOSingle]"

            DownTheRabbitHole -rootPaths $rootPaths -SourceGPO $SourceGPOSingle -DestinationGPO $DestinationGPO


        } Else {
            Write-Warning "SOURCE GPO DOES NOT EXIST [$SourceGPOSingle]"
        }

        $ProgressCounter++
    }

    #Write-Progress -Activity "Copying GPO settings to: $DestinationGPO" -Completed -Status "Complete"

}

Function DownTheRabbitHole {
[CmdletBinding()]
Param(
    [Parameter()]
    [String[]]
    $rootPaths,
    [Parameter()]
    [String]
    $SourceGPO,
    [Parameter()]
    [String]
    $DestinationGPO
)

    $ErrorActionPreference = 'Continue'

    ForEach ($rootPath in $rootPaths) {

        #Write-Verbose "SEARCHING PATH [$SourceGPO] [$rootPath]"
        Try {
            $children = Get-GPRegistryValue -Name $SourceGPO -Key $rootPath -ErrorAction Stop
        }
        Catch {
            #Write-Warning "REGISTRY PATH NOT FOUND [$SourceGPO] [$rootPath]"
            $children = $null
        }

        $Values = $children | Where-Object {-not [string]::IsNullOrEmpty($_.PolicyState)}
        If ($Values) {
            ForEach ($Value in $Values) {
                If ($Value.PolicyState -eq "Delete") {
                    #Write-Verbose "SETTING DELETE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                    If ([string]::IsNullOrEmpty($_.Valuename)) {
                        #Write-Warning "EMPTY VALUENAME, POTENTIAL SETTING FAILURE, CHECK MANUALLY [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"
                        Set-GPRegistryValue -Disable -Name $DestinationGPO -Key $Value.FullKeyPath | Out-Null
                    } Else {

                        # Warn if overwriting an existing value in the DestinationGPO.
                        # This usually does not get triggered for DELETE settings.
                        Try {
                            $OverWrite = $true
                            $AlreadyThere = Get-GPRegistryValue -Name $DestinationGPO -Key $rootPath -ValueName $Value.Valuename -ErrorAction Stop
                        }
                        Catch {
                            $OverWrite = $false
                        }
                        Finally {
                            If ($OverWrite) {
                                #Write-Warning "OVERWRITING PREVIOUS VALUE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)] [$($AlreadyThere.Value -join ';')]"
                            }
                        }

                        Set-GPRegistryValue -Disable -Name $DestinationGPO -Key $Value.FullKeyPath -ValueName $Value.Valuename | Out-Null
                    }
                } Else {
                    # PolicyState = "Set"
                    #Write-Verbose "SETTING SET [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)]"

                    # Warn if overwriting an existing value in the DestinationGPO.
                    # This can occur when consolidating multiple GPOs that may define the same setting, or when re-running a copy.
                    # We do not check to see if the values match.
                    Try {
                        $OverWrite = $true
                        $AlreadyThere = Get-GPRegistryValue -Name $DestinationGPO -Key $rootPath -ValueName $Value.Valuename -ErrorAction Stop
                    }
                    Catch {
                        $OverWrite = $false
                    }
                    Finally {
                        If ($OverWrite) {
                            #Write-Warning "OVERWRITING PREVIOUS VALUE [$SourceGPO] [$($Value.FullKeyPath):$($Value.Valuename)] [$($AlreadyThere.Value -join ';')]"
                        }
                    }

                    $Value | Set-GPRegistryValue -Name $DestinationGPO | Out-Null
                    # Added this sleep action to avoid the collision of AD processing the write and the next write being written
                    Start-Sleep -Seconds 3
                }
            }
        }
                
        $subKeys = $children | Where-Object {[string]::IsNullOrEmpty($_.PolicyState)} | Select-Object -ExpandProperty FullKeyPath
        if ($subKeys) {
            DownTheRabbitHole -rootPaths $subKeys -SourceGPO $SourceGPOSingle -DestinationGPO $DestinationGPO | Out-Null
        }
    }
}

    # Testing the path due to inconsitencies of the DISA Product
    if (Test-Path $GPObackups\$OfficeGPODirectorySafeName\GPO\ -PathType Any) {$GPODIR = "GPO"}
    if (Test-Path $GPObackups\$OfficeGPODirectorySafeName\GPOs\ -PathType Any) {$GPODIR = "GPOs"}    Write-Host `t"Creating the combined Office $OfficeVersion GPOs ..." -ForegroundColor Yellow    $UserGPODescription = New-Object System.Collections.Generic.List[System.Object]    $ComputerGPODescription = New-Object System.Collections.Generic.List[System.Object]    $GetGPOGUIDDirs = Get-ChildItem ("$GPObackups"+"\"+$OfficeGPODirectorySafeName+"\"+$GPODIR) -Directory            ForEach ($GPOGUIDDIR in $GetGPOGUIDDirs){        [xml]$manifest = Get-Content "$GPOBackups\$OfficeGPODirectorySafeName\$GPODIR\$GPOGUIDDIR\bkupInfo.xml"            ForEach ($gpBackup in $manifest.BackupInst){                $NewGPOName = $gpBackup.GPODisplayName.InnerText                $NewGPOBackupID = $gpBackup.ID.InnerText                $RandomNumber = Get-Random -Minimum 1 -Maximum 9999                $RandomNewGPOName = "$RandomNumber$NewGPOName"                New-GPO -Name $RandomNewGPOName -Server $PDCE|Out-Null                    if ("$RandomNewGPOName" -match "User")                    {                    Write-Host `t`t"-- Merging: $NewGPOName" -ForegroundColor Green                    Import-GPO -BackupId $NewGPOBackupID -Path $GPOBackups\$OfficeGPODirectorySafeName\$GPODIR -TargetName $RandomNewGPOName -Server $PDCE | Out-Null                    Copy-GPRegistryValue -Mode User -SourceGPO "$RandomNewGPOName" -DestinationGPO "$NewOfficeUserGPO"                    $UserGPODescription.Add("--"+"$NewGPOName")                    }                else                    {                    Write-Host `t`t"-- Merging: $NewGPOName" -ForegroundColor Green                    Import-GPO -BackupId $NewGPOBackupID -Path $GPOBackups\$OfficeGPODirectorySafeName\$GPODIR -TargetName $RandomNewGPOName -Server $PDCE | Out-Null                    Copy-GPRegistryValue -Mode Computer -SourceGPO "$RandomNewGPOName" -DestinationGPO "$NewOfficeComputerGPO"                    $ComputerGPODescription.Add("--"+"$NewGPOName")                    }                    Remove-GPO -Name "$RandomNewGPOName" -Server $PDCE                $manifest = $null                $gpbackup = $null            }    }   ####################################
#
# Adding DISA Provided GPO names to the descriptions field
# of the final GPOs, for history.
#
# Disabling the uneeded User or Computer policies
# of the final GPOs.
#
####################################    Write-Host `t`t"-- Disabling: Computer settings on $NewOfficeUserGPO" -ForegroundColor Green    (Get-GPO -Name "$NewOfficeUserGPO" -Server $PDCE).gpostatus="ComputerSettingsDisabled"    $UserGPODescription.Add("--"+"Added by: $env:USERNAME")    (Get-GPO -Name "$NewOfficeUserGPO" -Server $PDCE).Description="$UserGPODescription"    Write-Host `t`t"-- Disabling: User settings on $NewOfficeComputerGPO" -ForegroundColor Green    (Get-GPO -Name "$NewOfficeComputerGPO" -Server $PDCE).gpostatus="UserSettingsDisabled"    $ComputerGPODescription.Add("--"+"Added by: $env:USERNAME")    (Get-GPO -Name "$NewOfficeComputerGPO" -Server $PDCE).Description="$ComputerGPODescription"}
#
# Clean Up
#
Remove-Item -Path $GPOExtractDestination -Recurse -Force

Write-Host "Done!" -ForegroundColor Cyan
Write-Host "##########################" -ForegroundColor Cyan