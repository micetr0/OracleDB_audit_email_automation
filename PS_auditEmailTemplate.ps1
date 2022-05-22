function Get-auditEmailInfo
{
 <#*******************************************************************************
 Purpose: using database inputs to format email output 

 Dependency: none

 Output: to, subject, body, attachment (img)

 NOTE: 2 different prefix added to mail subject depending on scenario:

 HC ANALYST NEEDED   - for when theres entry in record where it should only go to HC analyst for further analysis
 MISSING STAFF EMAIL - for when there is no registered APP user email to grab 
    
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 25-June-2021   William Hu     cid embedd image within email
 18-May-2021    William Hu     added misc3 misc4 in parm
 15-APR-2021    William Hu     added ms link
 22-Mar-2021    William Hu     Initial version
 *******************************************************************************#> 

    [CmdletBinding()]
        Param
        ( [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           $auditType,       
          [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           $auditEmailSubj,
          [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           $auditSTF,
          [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
          $auditSTFEmail,
          [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
           $auditAppPerID,
          [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
           $auditError,
          [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
           $auditErrorFix,
          [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
          $auditMiscID1,
          [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
          $auditMiscID2,
          [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
          $auditMiscID3

        )

    #test
    write-verbose "parm value check"
    write-verbose "auditType: $auditType"
    write-verbose "auditEmailSubj: $auditEmailSubj"
    write-verbose "auditSTF: $auditSTF"
    write-verbose "auditSTFEmail: $auditSTFEmail"
    write-verbose "auditAppPerID: $auditAppPerID"
    write-verbose "auditError: $auditError"
    write-verbose "auditErrorFix: $auditErrorFix"
    write-verbose "auditMiscID1: $auditMiscID1" 
    write-verbose "auditMiscID2: $auditMiscID2"
    write-verbose "auditMiscID3: $auditMiscID3"
    write-verbose "auditMiscID4: $auditMiscID4"

    write-verbose "Get-auditEmailInfo: audit_stf_email is:$($entry.AUDIT_STF_EMAIL)"
    #account for missing staff email
    if($auditSTFEmail -eq [DBNULL]::Value)
    {
        write-verbose "Get-auditEmailInfo: empty staff email"
        $stfEmailTO = "yourAuditEmail@test.ca"
        $stfEmailSubject = "MISSING STAFF EMAIL - $($auditEmailSubj)"
    }
    else
    {
        write-verbose "Get-auditEmailInfo: not empty staff email"         
         $stfEmailTO = $auditSTFEmail
         $stfEmailSubject = $entry.AUDIT_EMAIL_SUBJ
    } 

    <#******************************************** AUDIT SPECIFIC RULES *****************************************#>

    if ($auditType -eq "AUHC04") 
    {
        if($auditMiscID1 -eq "HSUP")
        {   write-verbose "Get-auditEmailInfo: audittype AUHC04 and ref reason = HSUP"
            $stfEmailTo = "yourAuditEmail@test.ca"
            $stfEmailSubject = "HC ANALYST NEEDED - $($auditEmailSubj)"
        }
    }


    write-verbose "Get-auditEmailInfo: before switch"

    $rto          = $stfEmailTO
    $rsubject     = $stfEmailSubject

    switch -Wildcard ($auditType)
    {
            'AUHC03' {

            write-verbose "Get-auditEmailInfo: chosen AUHC03"
             

            $rBody        = "<html>

                        <style>

                        #empathsizeText {
                            color: blue;                        
                        }
                       </style>

                        <p>Hi $($auditSTF),</p>

                        <p>A weekly audit in APP has flaggged an Assessment (Assessment Form ID:<b>$($auditMiscID1)</b>) as some condition

                        <P>Some condition Assessment in APP <u>must have</u>:
                        
                        <ol id=""empathsizeText"">
                        <li>Assessment Reason = Some reason</li>
                        <li>A Manager Authorization Date filled in</li>
                        <li>One of the 3 reduced rate fields filled in</li>
                        </ol>

                        <p>Assessment associated with yourAuditEmail@test.ca: <b>$($auditAppPerID)</b> has the following error(s):</p>            
                        <ul> $($auditError) </ul>    

                        <p>Please correct the service order in APP by completing the following change(s):</p>
                        <ul> $($auditErrorFix) </ul>

                        <img src=`"cid:AUHC03_eg.png`">

                        <p>Thank you,</p>

                        <p>APP Admin</p>
                        <img src=`"cid:email_sig.png`">

                       <p>NOTE: If you are unable to see the images in this email, please click <a href=""https://support.microsoft.com/en-us/office/block-or-unblock-automatic-picture-downloads-in-email-messages-15e08854-6808-49b1-9a0a-50b81f2d617a"" target=_blank>here</a> for a one-time setup steps on how to unblock the images.</p>
                        </html>" 
                                   
           $rAttachments = "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC03\AUHC03_eg.png","$($(get-item $psscriptroot).parent.FullName)\Images\email_sig.png"                      
                     Break
                    }

            'AUHC04' {

            write-verbose "Get-auditEmailInfo: chosen AUHC04"
             

            $rBody        = "<html>

                <p>Hi $($auditSTF),</p>

                <p>A query in APP shows the below client has a type of <b>$($auditError)</b> that is attached to a closed ticket (ticket ID:<b> $($auditErrorFix) </b>/ referral reason: <b> $($auditMiscID1)) </b> belonging to <b> $($auditMiscID2) </b> team.</p>              

                <p>APP PER ID: <b>$($auditAppPerID)</b></p>

                <P>If the referral should still be open, please submit a form requesting that the referral be re-opened.</p>
                <p><a href=`"http:\\yourlinkhere.com`" target=_blank>yourlinkhere</a>
                <p>If the intervention should be closed, please enter an end date</p>

                Thank you,
                <p>APP Admin</p>
                <img src=`"cid:email_sig.png`">

                <p>NOTE: If you are unable to see the images in this email, please click <a href=""https://support.microsoft.com/en-us/office/block-or-unblock-automatic-picture-downloads-in-email-messages-15e08854-6808-49b1-9a0a-50b81f2d617a"" target=_blank>here</a> for a one-time setup steps on how to unblock the images.</p>
                    </html>" 
            $rAttachments = "$($(get-item $psscriptroot).parent.FullName)\Images\email_sig.png"    
                     Break
                    }

            'AUHC06' {

            write-verbose "Get-auditEmailInfo: chosen AUHC06"
             
            $rBody        = "<html>

                    <p>Hi $($auditSTF),</p>

                    <p>An audit shows a client with APP PER ID: <b>$($auditAppPerID)</b> has an ticket order with service start date of <b>$($auditMiscID2)</b> on <b>$($auditMiscID1)</b> team with the following error(s):</p>            
                    <ul> $($auditError) </ul>    

                    <p>Please correct the service order in APP by completing the following change(s):</p>
                    <ul> $($auditErrorFix) </ul>

                    <p>Thank you,</p>

                    <p>APP Admin</p>
                    <img src=`"cid:email_sig.png`">

                    <p>NOTE: If you are unable to see the images in this email, please click <a href=""https://support.microsoft.com/en-us/office/block-or-unblock-automatic-picture-downloads-in-email-messages-15e08854-6808-49b1-9a0a-50b81f2d617a"" target=_blank>here</a> for a one-time setup steps on how to unblock the images.</p>
                    </html>"   
            $rAttachments = "$($(get-item $psscriptroot).parent.FullName)\Images\email_sig.png" 
                     Break
                    }


            'AUHC07' {

            write-verbose "Get-auditEmailInfo: chosen AUHC07"
             
            $rBody        = "<html>

                <p>Hi $($auditSTF),</p>

                <p>A query in APP shows the below client has an property type of <b> $($auditMiscID1) </b> referred by the <b> $($auditErrorFix) </b> team with no one attached to it.</p>              

                <p>Please enter the <quote>Specialist</quote> in APP</p>
                <p>APP PER ID: <b> $($auditAppPerID) </b></p>

                <P><img src=`"cid:AUHC07_eg.png`"></p>

                <table>
                <tr>
                    <td><h4><u>Steps to view</u></h4></td>
                </tr>
                <tr>
                    <td>1.Click ""More actions""</td>
                </tr>
                <tr>
                    <td><img src=`"cid:AUHC07_his1.png`"></td>
                </tr>
                <tr>
                    <td>2.Click ""Blahblah2""</td>
                </tr>
                <tr>
                    <td><img src=`"cid:AUHC07_his2.png`"></td>
                </tr>
                <tr>
                    <td>3.Enter a ""End Date""of blah field</td>
                </tr>
                <tr>
                    <td><img src=`"cid:AUHC07_his3.png`"></td>
                </tr>
                <tr>
                    <td>4.Click ""More buttons""</td>
                </tr>
                <tr>
                    <td><img src=`"cid:AUHC07_his4.png`"></td>
                </tr>
                <tr>
                    <td>5.Find the informationr</td>
                </tr>
                <tr>
                    <td><img src=`"cid:AUHC07_his5.png`"></td>
                </tr>
                </table>

                Thank you,
                <p>APP Admin</p>
                <img src=`"cid:email_sig.png`"> 

                <p>NOTE: If you are unable to see the images in this email, please click <a href=""https://support.microsoft.com/en-us/office/block-or-unblock-automatic-picture-downloads-in-email-messages-15e08854-6808-49b1-9a0a-50b81f2d617a"" target=_blank>here</a> for a one-time setup steps on how to unblock the images.</p>
                    </html>" 

            $rAttachments =  "$($(get-item $psscriptroot).parent.FullName)\Images\email_sig.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_eg.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_his1.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_his2.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_his3.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_his4.png",
                             "$($(get-item $psscriptroot).parent.FullName)\Images\AUHC07\AUHC07_his5.png"

                     Break
                    }
                                            
            default { write-error "Get-auditEmailInfo: file pattern does not match any existing HH audits"}
        }

        write-verbose "Get-auditEmailInfo: before return"
        return $rTo, $rSubject, $rBody ,$rAttachments      
}