<#
Version: 2.0 
# By Mustafa Nassar, Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.


.SYNOPSIS
This script generates all the needed reports to troubleshoot a move request from or to Exchange Online/OnPrem. It can handle multiple mailboxes at once.

.DESCRIPTION
  This script will create: 
    1. a summary report of a move request from an exchange server that includes the following information:
        a.	The move request's status and percentage.
        b.	The failure types for the move request. 
        c.	The detailed message for each failure. 

    2. Move Request Report 
    3. Move Request Statistics Report 
    4. CSV file including all failure details 
    5. Batch report
    6. Migration users and migration user statistics Reports
    7. Mailbox folders Statistics 
    8. Migration Configuration 

.NOTES
    this script should be excuted in exchange online or Exchange OnPrem Powershell module. 
.EXAMPLE
    .\Get-MigrationReports -Mailboxes Mustafa@contoso.com 
    
    .\Get-MigrationReports -Mailboxes user1@contoso.com, user2@contoso.com, user3@contoso.com 
   
Auther: Mustafa Nassar 

#>

# declare the parameters 
[CmdletBinding()]
param (
    [Parameter( Mandatory = $true, HelpMessage = 'You must specify the name of a mailbox or mailboxes:')] [array] $Mailboxes
    #[Parameter(Mandatory = $false)][XML] $MoveRequestStatistics = (Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose")

)
  $folder = 'Get-MigrationReports'

function Export-XMLReports {
    # Export XML reports:  
    Try {
        if (-not $null -eq $MoveRequest) {
            $MoveRequest | Export-Clixml "$folder\MoveRequest_$Mailbox.xml"
            Add-Content -Path $logFile -Value " [INFO] The Move Request Report has been generated successfully."
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Move Request not exist."
        }
        
        if (-not $null -eq $MoveRequestStatistics) {
            $MoveRequestStatistics | Export-Clixml "$folder\MoveRequestStatistics_$Mailbox.xml"
            Add-Content -Path $logFile -Value " [INFO] The Move Request Statistics Report has been generated successfully."
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Move Request Statistics not exist."
        }
        
        if (-not $null -eq $UserMigration) {
            $UserMigration | Export-Clixml "$folder\MigrationUser_$Mailbox.xml" 
            Add-Content -Path $logFile -Value " [INFO] The User Migration Report has been generated successfully."
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Migration User not exist."
        }

        if (-not $null -eq $UserMigrationStatistics) {
            $UserMigrationStatistics | Export-Clixml "$folder\MigrationUserStatistics_$Mailbox.xml"
            Add-Content -Path $logFile -Value " [INFO] The Migration User Statistics Report has been generated successfully."
    
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Migration User Stistics Report not exist."
        }

        if (-not $null -eq $MigrationBatch) {
            $MigrationBatch | Export-Clixml "$folder\MigrationBatch_$Mailbox.xml"
            Add-Content -Path $logFile -Value " [INFO] The Migration Batch Report has been generated successfully."
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Migration Batch not exist."
        } 
        
        if (-not $null -eq $MigrationEndPoint) {
            $MigrationEndPoint | Export-Clixml "$folder\MigrationEndpoint_$MigrationEndpoint.xml"
            Add-Content -Path $logFile -Value " [INFO] The Migration EndPoint Report has been generated successfully."
        }
        else {
            Add-Content -Path $logFile -Value " [Error] The Migration EndPoint not exist."
        } 
        
        #$MoveRequest | Export-Clixml "MoveRequest_$Mailbox.xml"
        #Add-Content -Path $logFile -Value " [INFO] The Move Request Report has been generated successfully."
        #$MoveRequestStatistics | Export-Clixml "MoveRequestStatistics_$Mailbox.xml"
        #Add-Content -Path $logFile -Value " [INFO] The Move Request Statistics Report has been generated successfully."
    
        #$UserMigration | Export-Clixml "MigrationUser_$Mailbox.xml" 
        #Add-Content -Path $logFile -Value " [INFO] The User Migration Report has been generated successfully."
    
        #$UserMigrationStatistics | Export-Clixml "MigrationUserStatistics_$Mailbox.xml"
        #Add-Content -Path $logFile -Value " [INFO] The Migration User Statistics Report has been generated successfully."
    
        #$MigrationBatch | Export-Clixml "MigrationBatch_$Batch.xml"
        #Add-Content -Path $logFile -Value " [INFO] The Migration Batch Report has been generated successfully."

        #$MigrationEndPoint | Export-Clixml "MigrationEndpoint_$MigrationEndpoint.xml"
        #Add-Content -Path $logFile -Value " [INFO] The Migration EndPoint Report has been generated successfully."

        Get-MigrationConfig | Export-Clixml "$folder\MigrationConfig.xml" 
        Add-Content -Path $logFile -Value " [INFO] The Migration Config Report has been generated successfully."

        $MailboxStatistics | Export-Clixml "$folder\MailboxStatistics_$Mailbox.xml"
        $MoveHistory.MoveHistory[0] | Export-Clixml "$folder\MoveReport-History.xml"
        Add-Content -Path $logFile -Value " [INFO] The Move Request History Report has been generated successfully."

    }
    catch {
        Add-Content -Path $logFile -Value '[ERROR] Unable to export the Reports.'
        Add-Content -Path $logFile -Value $_
        throw
    }

}
function Export-Summary {
    #check the log file 
    if (-not (Test-Path -Path $logfile -ErrorAction Stop )) 
    {
        # Create a new log file if not found.
        New-Item $logfile  -Type File -Force  -ErrorAction SilentlyContinue
    }
    

    #Create a Summary Report: 
    try {
        if (-not (Test-Path -Path $file -ErrorAction Stop )) {
            # Create a new log file if not found.
            New-Item $file   -Type File -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Add-Content -Path $logfile -Value '[ERROR] Unable to Create Summary File.'
        Add-Content -Path $logFile -Value $_
        throw
    }

    [int]       $Percent = $MoveRequestStatistics.PercentComplete
    [string]    $Status  = $MoveRequestStatistics.Status
    [string]    $Message = $MoveRequestStatistics.Message 
    Add-Content $file -Value "This Move Request has the following infomration:"
    Add-Content $file -Value "-----------------------------------------------------------------------"
    Add-Content $file -Value "the status of this Move Request is $Status with $Percent  Percent"
    Add-Content $file -Value ""
    Add-Content $file -Value "$Message "
    Add-Content $file -Value ""
    Add-Content $file -Value "-----------------------------------------------------------------------"
    Add-Content $file -Value ""
    Add-Content $file -Value "The Move Request has the following Failures:"
    Add-Content $file -Value "-----------------------------------------------------------------------"
    Add-Content $file -Value ""
    $Uniquefailure =  $MoveRequestStatistics.Report.Failures | Select-Object FailureType -Unique 
    foreach ($x in $Uniquefailure) {
        $x.FailureType >> ($file)
    }
    Add-Content $file -Value ""
    Add-Content $file -Value ""
    Add-Content $file -Value "Here is more details about each Failure (Note that only the last error is selected in more details):"
    Add-Content $file -Value "-----------------------------------------------------------------------"
    Add-Content $file -Value ""
    $DetailedFailure = foreach ($U in $uniquefailure) {$MoveRequestStatistics.Report.Failures | ? {$_.FailureType -like $U.FailureType} |select Timestamp, FailureType, FailureSide, message -ExpandProperty Message -Last 1 }
    foreach ($f in $DetailedFailure) {
        $f >> ($file); 
        Add-Content $file -Value ""
    }
    
    Add-Content -Path $logFile -Value "[INFO] the summary report has been created successfully."    
    Add-Content -Path $logFile -Value " [INFO] the summary report has been created successfully." 
}

#===================MAIN======================
New-Item $folder -ItemType Directory -Force | Out-Null 
New-Item $logFile -Type File -Force -ErrorAction SilentlyContinue  | Out-Null

foreach ($Mailbox in $Mailboxes) {

# Define the error prefrence for the script 
$ErrorActionPreference = 'SilentlyContinue'
# Declare the general used variables
$MoveRequest = Get-MoveRequest $Mailbox -ErrorAction SilentlyContinue 
$MoveRequestStatistics = Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction SilentlyContinue
$Batch = $MoveRequestStatistics.BatchName  
$MigrationBatch = Get-MigrationBatch $Batch -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction SilentlyContinue
$UserMigration = Get-MigrationUser $Mailbox  -ErrorAction SilentlyContinue
$UserMigrationStatistics = Get-MigrationUserStatistics $Mailbox -IncludeSkippedItems -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -ErrorAction SilentlyContinue 
$Endpoint = $MigrationBatch.SourceEndpoint 
$MigrationEndPoint = Get-MigrationEndpoint -Identity $Endpoint -DiagnosticInfo Verbose -ErrorAction SilentlyContinue
$MailboxStatistics = Get-MailboxStatistics $Mailbox -IncludeMoveReport -IncludeMoveHistory -ErrorAction SilentlyContinue
$MoveHistory = Get-MailboxStatistics $Mailbox -IncludeMoveReport -IncludeMoveHistory -ErrorAction SilentlyContinue
$Uniquefailure =  $MoveRequestStatistics.Report.Failures | select FailureType -Unique 
$DetailedFailure = foreach ($U in $uniquefailure) {$MoveRequestStatistics.Report.Failures | ? {$_.FailureType -like $U.FailureType} |select Timestamp, FailureType, FailureSide, Message -Last 1 |ft -Wrap }
$File = "$folder\Text-Summary_$Mailbox.txt"
$logFile = "$folder\LogFile_$Mailbox.txt" 
New-Item $file   -Type File -Force -ErrorAction SilentlyContinue | Out-Null

    try 
    {
        if (-not $null -eq $MoveRequestStatistics ) 
        {
        Export-XMLReports
        Export-Summary 
        Write-Host -ForegroundColor "Green" "The MoveRequest reports for $Mailbox exported successfully!"
        }   
    }
    catch {
        Add-Content -Path $logFile -Value "[ERROR] The MoveRequest for the $Mailbox cannot be found, please check spelling and try again!"
        Add-Content -Path $logFile -Value $_
        Write-Host -ForegroundColor "Red" "The MoveRequest for the $Mailbox cannot be found, please check spelling and try again!"
            
        throw
    }


}

$compress = @{
    Path = $folder
    CompressionLevel = "Fastest"
    DestinationPath = "Migration-Reports.Zip"
}
Compress-Archive @compress


