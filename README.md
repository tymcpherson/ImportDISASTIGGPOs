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
#  5) Creates the combined Office user and computer GPOs per version
#  6) Creates a temporary random named Office Product GPO, then merges that products settings into the combined Office
#     user or computer version
#  7) Removes the temporary random Office Product GPO
#  8) Adds the Office product STIG versioning to the description
#  9) Cleans up the %TEMP% working directory
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
