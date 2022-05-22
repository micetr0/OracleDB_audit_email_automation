<#*******************************************************************************

 Purpose: generate audit email to send to staff directly from yourAuditEmail@test.ca

 Dependency: app.pkg_for_nightly_run --drives the PLSQL to populate the audit tables listed, 
             TABLE app.audit_information --audit information,  
             TABLE app.audit_count --keep track of audit runs as to limit number of emails sent to user, 
             credential from PSCredential folder

 TO-DO: for go-live change auditEmailtemplate to field to use #$auditSTFEmail instead of tester name  

 SCRIPT IS MEANT TO BE RUN ONCE A DAY - if troubleshooting please make sure to remove recordset from 
                                        the same date period in the app.audit_information and app.audit_count to avoid
                                        picking up duplicate or comment out Run-OracleProcedure
    
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 25-June-2021   William Hu     Added image embedding by using -attachment to emails sent
 18-MAY-2021    William Hu     Modified query to accomodate for emailCounter
                               Tweaked logic/wording for summaryEmail so it sends when audit has 0 records for historic tracking 
 05-MAY-2021    William Hu     Add custom csv header details
 20-Mar-2021    William Hu     Initial version

 *******************************************************************************#>
 
 #load functions
.$psScriptRoot\PS_auditEmailTemplate.ps1
.$psScriptRoot\PS_oracleDBFunctions.ps1
.$psScriptRoot\PS_zipFileFunc.ps1

$logPath = $psScriptRoot+"\RunLogs\" + (get-date -format "dd-MMM-yyyy-HHmmss") + ".log" #((get-item $psScriptRoot).parent.FullName + "\RunLogs\") #one level up
$errorLogPath = $psScriptRoot+"\ErrorLogs\" + (get-date -format "dd-MMM-yyyy-HHmmss") +"error" + ".log"

Start-Transcript -Path $logPath

$beginTime = (Get-Date).DateTime
write-host "MAIN: begin"

#PARISHCAUDIT
$DBpassword =  Get-Content "$psScriptRoot\PSCredential\DBpassword.txt" | ConvertTo-SecureString -Key (Get-Content "$psScriptRoot\PSCredential\encryptAES.key")
$DBcredential = New-Object System.Management.Automation.PSCredential("yourDBCredential",$DBpassword)
$DBHost = "yourDBHost"
$DBService = "yourDBName "

$inboxToUse = "yourAuditEmail@test.ca"
$CC = "yourAuditEmail@test.ca"  #inbox cc set so can use outlook forwarding rule to file these as sent
$CSVoutputDir = "Process-Summary" #where csv or data ouput is stored, these will then be attached to a summary email to be mailed out and then removed from dir
$logDate = (get-date -format "dd-MMM-yyyy-HHmmss")
$summaryBody = " " #reset for each run
$zipswd= "yourZipPassword" #temp for now - working on password manager solution

<#************************add new audit entries here ***************************#>  

#get-content instead to store SQLs but leave in PS for hub example
$auditAUHC03 = New-Object PSObject -Property @{
    Name = "AUHC03"
    SP = "app.pkg_for_nightly_run.spYourStoreProcedure1"
    Query = "SELECT 
                AUDIT_ID,
                AUDIT_TYPE,
                AUDIT_TIME,
                AUDIT_EMAIL_SUBJ,
                audit_stf_id,
                AUDIT_STF,
                AUDIT_STF_EMAIL,
                AUDIT_APP_PER_ID,
                AUDIT_ERROR,
                AUDIT_ERROR_FIX,
                AUDIT_MISC_IDENTIFIER1,
                AUDIT_MISC_IDENTIFIER2

            FROM app.audit_information a
            INNER JOIN app.audit_count c
            ON c.achk_rec_id = a.audit_misc_identifier1
            AND c.achk_rec_amended_stf_id = a.audit_stf_id
            AND c.achk_app_per_id = a.AUDIT_APP_PER_ID
            WHERE a.audit_type = 'AUHC03'
            AND a.audit_time > TRUNC(SYSDATE)
            AND c.achk_aud_counter <= 3"
    }


$auditAUHC04 = New-Object PSObject -Property @{
    Name = "AUHC04"
    SP = "app.pkg_for_nightly_run.spYourStoreProcedure2"
    Query =  "SELECT 
                AUDIT_ID,
                AUDIT_TYPE,
                AUDIT_TIME,
                AUDIT_EMAIL_SUBJ,
                audit_stf_id,
                AUDIT_STF,
                AUDIT_STF_EMAIL,
                AUDIT_APP_PER_ID,
                AUDIT_ERROR,
                AUDIT_ERROR_FIX,
                AUDIT_MISC_IDENTIFIER1,
                AUDIT_MISC_IDENTIFIER2

        FROM app.audit_information a
        INNER JOIN app.audit_count c
            ON c.achk_rec_id = a.audit_error_fix -- ref_id
            AND c.achk_rec_amended_stf_id = a.audit_stf_id
            AND c.achk_app_per_id = a.AUDIT_APP_PER_ID
            WHERE a.audit_type = 'AUHC04'
            AND a.audit_time > TRUNC(SYSDATE)
            --AND c.achk_aud_counter <= 3"
    }  

$auditAUHC06 = New-Object PSObject -Property @{
    Name = "AUHC06"
    SP = "app.pkg_for_nightly_run.spYourStoreProcedure3"
    Query ="SELECT 
            AUDIT_ID,
            AUDIT_TYPE,
            AUDIT_TIME,
            AUDIT_EMAIL_SUBJ,
            audit_stf_id,
            AUDIT_STF,
            AUDIT_STF_EMAIL,
            AUDIT_APP_PER_ID,
            AUDIT_ERROR,
            AUDIT_ERROR_FIX,
            AUDIT_MISC_IDENTIFIER1,
            AUDIT_MISC_IDENTIFIER2,
            AUDIT_MISC_IDENTIFIER3

        FROM app.audit_information a
        INNER JOIN app.audit_count c
            ON c.achk_rec_id = a.audit_misc_identifier3 -- cpt_id
            AND c.achk_rec_amended_stf_id = a.audit_stf_id
            AND c.achk_app_per_id = a.AUDIT_APP_PER_ID
            WHERE a.audit_type = 'AUHC06'
            AND a.AUDIT_ERROR <> ' '
            AND a.audit_time > TRUNC(SYSDATE)
            --AND c.achk_aud_counter <= 3"
    }

$auditAUHC07 = New-Object PSObject -Property @{
    Name = "AUHC07"
    SP = "app.pkg_for_nightly_run.spYourStoreProcedure4"
    Query = "SELECT 
                audit_id,
                audit_type,
                audit_time,
                audit_email_subj,
                audit_stf_id,
                audit_stf,
                audit_stf_email,
                AUDIT_APP_PER_ID,
                audit_error_fix,
                AUDIT_MISC_IDENTIFIER1,
                AUDIT_MISC_IDENTIFIER2,
                AUDIT_MISC_IDENTIFIER3

            FROM app.audit_information a
            INNER JOIN app.audit_count c
            ON c.achk_rec_id = a.audit_misc_identifier3 --cpi_id
            AND c.achk_rec_amended_stf_id = a.audit_stf_id
            AND c.achk_app_per_id = a.AUDIT_APP_PER_ID
            WHERE a.audit_type = 'AUHC07'
            AND a.audit_time > TRUNC(SYSDATE)
            --AND c.achk_aud_counter <= 3"
    }

#Reset variable (when doing multiple runs)
#$Audits = @($auditAUHC07)
$Audits = @($auditAUHC03,$auditAUHC04,$auditAUHC06,$auditAUHC07)

<#*****************************************************************************#> 

foreach ($Audit in $Audits)
{
    $auditSentCounter = 0
    $missingStfEmailCounter = 0

    Run-OracleProcedure $DBcredential.UserName $DBcredential.GetNetworkCredential().Password $DBHost $DBService $audit.SP  -errorAction stop
    $resultSQLDataSet = Run-OracleSQLQuery $DBcredential.UserName $DBcredential.GetNetworkCredential().Password $DBHost $DBService $audit.QUery -errorAction stop

    If ($resultSQLDataSet.Count -eq 0)
    {
        write-host "MAIN: no new entries for $($Audit.Name)"
        $summaryBody += "<p>HH Audit Type: $($Audit.Name) - has no audit entries</p>"
    }
    else
    {
        #save out the query output as CSV in Process-Summary folder
        $resultSQLDataSet | Export-Csv -path "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -NoTypeInformation

        #format csv headers for each CSV audits output - need to add for new audit entry
        if (($Audit.Name -eq 'AUHC03') -and (Test-Path -path "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv"))
        {
            $tempCSV = Import-Csv "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -Header 'AUDIT_ID', 'AUDIT_TYPE', 'AUDIT_TIME','AUDIT_EMAIL_SUBJ','AUDIT_STF_ID', 'AUDIT_STF', 'AUDIT_STF_EMAIL', 'AUDIT_APP_PER_ID', 'AUDIT_ERROR', 'AUDIT_ERROR_FIX', 'FIN_AX_ID' | select -skip 1
            $tempCSV | Export-CSV "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -NoTypeInformation
        }
        if (($Audit.Name -eq 'AUHC04') -and (Test-Path -path "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv"))
        {
            $tempCSV = Import-Csv "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -Header 'AUDIT_ID', 'AUDIT_TYPE', 'AUDIT_TIME','AUDIT_EMAIL_SUBJ','AUDIT_STF_ID', 'AUDIT_STF', 'AUDIT_STF_EMAIL', 'AUDIT_APP_PER_ID', 'INTERVENTION_CODE', 'REF_ID', 'REF_REASON', 'REF_TEAM' | select -skip 1
            $tempCSV | Export-CSV "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -NoTypeInformation
        }
        if (($Audit.Name -eq 'AUHC06') -and (Test-Path -path "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv"))
        {
            $tempCSV = Import-Csv "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -Header 'AUDIT_ID', 'AUDIT_TYPE', 'AUDIT_TIME','AUDIT_EMAIL_SUBJ','AUDIT_STF_ID', 'AUDIT_STF', 'AUDIT_STF_EMAIL', 'AUDIT_APP_PER_ID', 'AUDIT_ERROR', 'AUDIT_ERROR_FIX', 'INTERVENTION_TEAM','SERVICE_START_DATE', 'CPT_ID' | select -skip 1
            $tempCSV | Export-CSV "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -NoTypeInformation
        }
        if (($Audit.Name -eq 'AUHC07') -and (Test-Path -path "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv"))
        {
            $tempCSV = Import-Csv "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -Header 'AUDIT_ID', 'AUDIT_TYPE', 'AUDIT_TIME','AUDIT_EMAIL_SUBJ','AUDIT_STF_ID', 'AUDIT_STF', 'AUDIT_STF_EMAIL', 'AUDIT_APP_PER_ID', 'REF_TEAM', 'INTERVENTION_CODE', 'SERVICE_START_DATE', 'CPI_ID' | select -skip 1
            $tempCSV | Export-CSV "$PSScriptRoot\$CSVoutputDir\$logDate-$($Audit.Name).csv" -NoTypeInformation
        }

        foreach ($entry in $resultSQLDataSet)
        {
            $DBdataInputs = @{
            auditType =     $entry.AUDIT_TYPE
            auditEmailSub = $entry.AUDIT_EMAIL_SUBJ
            auditSTF =      $entry.AUDIT_STF
            auditSTFEmail = $entry.AUDIT_STF_EMAIL
            auditAppPerID =  $entry.AUDIT_APP_PER_ID
            auditError =    $entry.AUDIT_ERROR
            auditErrorFix = $entry.AUDIT_ERROR_FIX
            auditMiscID1 =  $entry.AUDIT_MISC_IDENTIFIER1
            auditMiscID2 =  $entry.AUDIT_MISC_IDENTIFIER2
            auditMiscID3 =  $entry.AUDIT_MISC_IDENTIFIER3
            }

            $to, $subject, $body ,$attachments = Get-auditEmailInfo @DBdataInputs #-Verbose


            if ($Subject -like '*MISSING STAFF EMAIL*')
            {
                $missingStfEmailCounter++
            }

            $mailHash = @{
            to = $to.Split(';') #for when you have to send to multiple person
            cc = $cc
            from = $inboxToUse
            subject = $subject
            body = $body
            smtpserver = "your.smtp.server.ca"
            bodyashtml = $true    
            attachments = $attachments   
            }

            try
            {    
                Send-MailMessage @mailHash -ErrorAction Stop
            }
            catch 
            {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                $Time = get-date 
                "Failed at $Time trying to send audit email for audit: $($audit.Name) for $($auditAppPerID) FailedItem:  $FailedItem. ErrorMsg: $ErrorMessage saved in errorlog" | out-file $errorLogPath -append
                Write-Warning "Failed at $Time trying to send audit email for audit: $($audit.Name) for $($auditAppPerID) FailedItem:  $FailedItem. ErrorMsg: $ErrorMessage"
            Break    
            }                   

            $auditSentCounter++
            $auditType = $entry.AUDIT_TYPE

        } #for-each entry

        Write-host "MAIN: HH Audit Type: $($audit.Name) - $auditSentCounter Emails Sent / $missingStfEmailCounter of the email sent is without staff email"
        $summaryBody += "<p>HH Audit Type: $($audit.Name) - $auditSentCounter Emails Sent / $missingStfEmailCounter of the email sent is without staff email</p>" 
         
    } #else dataset not null
}#end for-each audit


#hash needs to be before csv counter to build dynamic splatting
$summaryMailHash = @{
        to = $inboxToUse
        cc = $cc
        from = $inboxToUse
        subject = "$logDate-HC Audit Summary"
        body = $summaryBody
        smtpserver = "your.smtp.server.ca"
        bodyashtml = $true
        }

$csvOuputCounter = (get-childitem "$PSScriptRoot\$CSVoutputDir\*" | measure).count

If ($csvOuputCounter -eq 0) 
{
    write-host "MAIN: no csv file(s) detected - abort zipping"
}
else
{
    #zip / encrypt with password for all CSVs stored in Process-Summary Folder
    $zipFileName = $logDate + "-HC Audit Run Summary"
    Write-ZipUsing7Zip -FilesToZip "$PSScriptRoot\$CSVoutputDir\*.csv" -ZipOutputFilePath "$PSScriptRoot\$CSVoutputDir\$zipFileName.zip" -Password $zipswd -HideWindow
    $filestoAttach = (Get-ChildItem -Path $PSScriptRoot\$CSVoutputDir -Filter *.zip).FullName
    #add to summaryMailHash through splatting  
    $summaryMailHash.attachments = $filestoAttach 
}

try
{    
    Send-MailMessage @summaryMailHash -ErrorAction Stop
}
catch 
{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    $Time = get-date 
    "Failed at $Time trying to send summary email. FailedItem:  $FailedItem. ErrorMsg: $ErrorMessage saved in errorlog" | out-file $errorLogPath -Append
    Write-Warning "Failed at $Time trying to send audit email for audit: $($audit.Name) for $($auditAppPerID) FailedItem:  $FailedItem. ErrorMsg: $ErrorMessage"
    Break    
}

#clear process summary folder content for next run
Remove-Item -Path  "$PSScriptRoot\$CSVoutputDir\*"

#Run-Stats
$endTime = (get-Date).DateTime
$totalTime = New-TimeSpan -Start $beginTime -End $endTime
write-host "MAIN: end - total run time: $totalTime"
Stop-Transcript