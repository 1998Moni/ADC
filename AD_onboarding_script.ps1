cls

Remove-Variable * -ErrorAction SilentlyContinue

#region ----------------------------------[Initialisations]----------------------------------

$ErrorActionPreference = "Stop";
$Accounts = @()

$StagingOU = "path of staging OU"
$trgtOU = "path of target OU" 

$logfilepath = "path of Audit_Logs"

#endregion

Start-Transcript -Path "$logfilepath\Log_$(get-date -f MMddyyyy)_AD.log" -Append

Write-Host "----------Starting Script at $(get-date -f "MM/dd/yyyy HH:mm:ss tt") ----------"

#region -----------------------------[Workflows and Functions]-------------------------------

Function Failure
{
    $row1 = New-Object PSObject
    $row1 | Add-Member -MemberType NoteProperty -Name "CWID" -Value $CWID
    $row1 | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName.Name
    $row1 | Add-Member -MemberType NoteProperty -Name "ERROR" -Value $Err

    $FailedUser_AD += $row1

    $body = $NULL
    $body += "<body><table width=""560"" border=""1""><tr style='background-color:yellow'>"
    $FailedUser_AD[0] | ForEach-Object {
    foreach ($property in $_.PSObject.Properties){$body += "<td>$($property.name)</td>"}} 
    $body += "</tr><tr>"
    $FailedUser_AD | ForEach-Object {
    foreach ($property in $_.PSObject.Properties){$body += "<td>$($property.value)</td>"}
    $body += "</tr><tr>"
    }
    $body += "</tr></table></body>"

    #region -----------------------------Failure Email-------------------------------

    if($FailedUSER_AD)
    {  
   
        $messagebody ="
        <br>
        Hi Team,
        <br>
        <br>
        There is an issue while executing User_Onboarding_AD_Script_v20 <br>
        $body
        <br>
        <br>
        Logs are saved on ********* at c:\usrs\Audit_Logs\Log_$(get-date -f MMddyyyy)_AD.log .
        <br>
        <br>
        Regards,
        <br>
        PHC Group IT Service Desk
        <br>
        <a href='http://servicedesk.*******.com'>http://servicedesk.*******.com</a>
        <br>
        <br>
        ** This is a system generated email.
        <br>
        "
        
        $to = ' '
        $bcc = ' '
        
        $name = $DisplayName.Name

        write-host "Sending Error Email for $name" -ForegroundColor Yellow
        write-host ""

        Send-MailMessage -SmtpServer “stmp address” -From "from address" -To $to -Bcc $bcc -Subject “Subject” -BodyAsHtml -Body $messagebody  
            
        $ManagerName=(Get-aduser -filter "displayname -like '$($name)'" -Properties *).Manager
        $Managercwid=(Get-Aduser $ManagerName -Erroraction SilentlyContinue).userprincipalname

       

        $user = "*************"
        $pass = "*************"

        # Build auth header
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $pass)))

        # Set proper headers
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')
        $headers.Add('Content-Type','application/json')

        # Specify endpoint uri
        
        $uri = "https://******.service-now.com/api/hclad/create_ticket/onboarding_incident/AutomationRestUser"
        
        # Specify HTTP method
        $method = "post"

        $body = @{
        short_description="short description";
        domain="*********";
        type="AD"
        assignment_group="";
        state="1";
        description="DisplayName: $($name) `nCWID: $($cwid) `nError: $($Err)";
        user_cwid=$name
        } | ConvertTo-Json -Compress        

        write-host "Creating Failure incident in service now for $name" -ForegroundColor Yellow
        write-host ""

        # Send HTTP request
        $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body

        $result = $response.result

        if($result)
        {
        write-host "Please find below the ticket details:"
        write-host ""
        $result
        }
        else{
        write-host "Error while creating ticket in service now" -ForegroundColor Red
        write-host ""
        }
    }
}

#endregion

#region -----------------------------[Executions]-------------------------------

try
{
    $Accounts = (Get-ADUser -Filter * -SearchBase $StagingOU -SearchScope OneLevel -Verbose -ErrorAction Stop | Where-Object {$_.distinguishedname -notlike "*OU=Partially Provisioned,OU=Staging,OU=****,DC=****,DC=NET"}).samaccountname
 
    foreach($NewUser in $Accounts)
    {    
        $Path = "C:\Users\Documents\Onboarding_Offboarding_Automation\Onboarding_Automation\PS_Scripts\CWID-Inventory-v1.csv"
        $Users = Import-Csv $Path
        $Users_Avail = Import-Csv $Path | Where-Object {$_.Availability -eq "Available"} | Select CWID
        $Available_CWID =  $Users_Avail | Select-Object -First 1
        $DataCount = Import-Csv $Path | Measure-Object  
        $date=Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        $script = "Onboarding"    
               
        write-host "$($NewUser) - Processing" -BackgroundColor Black   
        write-host ""
                 
        $user = $NewUser 

        #region-----------------DISPLAYNAME----------------# 
        try
        {
            $DisplayName = Get-ADUser -Identity $NewUser -ErrorAction Stop | Select Name
        }
        Catch
        {          
            $DisplayName = New-Object PSObject
            $DisplayName | Add-Member -MemberType NoteProperty -Name "Name" -Value $null               
        } 
        $CWID = $Available_CWID.CWID

        #endregion

        #region-----------------SAMACCOUNTNAME-------------#           
        try            
        {                    
            Set-ADUser -identity $NewUser -SamAccountName $CWID -UserPrincipalName $CWID"@******.net" -ErrorAction Stop -Verbose                                               
            Start-Sleep -Seconds 2
            $ErrorCheck = 1                       
        }            
        catch        
        {            
            $Err = "Error occured while adding CWID AND UPN for User OR User not found:- $user"
            $ErrorCheck = 0  
            Failure                
        }
        #endregion

        #region-----------------PROXYADDRESSES-------------# 
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        {
            $ErrorCheck = 0                                              
            try                    
            {
                $proxyaddress1 = "SMTP:"+((get-aduser $CWID).givenname).ToLower()+"."+((get-aduser $CWID).surname).ToLower()+"@*****.com"                    
                $proxyaddress2 = "smtp:"+$CWID+"@******.net"                    
                $proxyaddress3 = "smtp:"+$CWID+"@**************.com"                                        
                set-aduser $CWID -add @{proxyaddresses=$proxyaddress1} -Verbose -ErrorAction Stop                             
                set-aduser $CWID -add @{proxyaddresses=$proxyaddress2} -Verbose -ErrorAction Stop                             
                set-aduser $CWID -add @{proxyaddresses=$proxyaddress3} -Verbose -ErrorAction Stop                            
                write-host ""                            
                write-host "Proxy Addresses added.." -ForegroundColor Green     
                $ErrorCheck = 1                    
            }                    
            catch            
            {                
                $Err = "Error occured while updating Proxy Address"   
                $ErrorCheck = 0 
                Failure            
            } 
        }
        #endregion

        #region-----------------TARGETADDRESSES-------------#    
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        { 
            $ErrorCheck = 0                                              
            try                    
            {   
                $target = "smtp:"+$CWID+"@************.com"                   
                Set-ADUser $CWID -Replace @{targetAddress=($target)} -Verbose -ErrorAction Stop  
                                           
                write-host ""                            
                write-host "Target Address added.." -ForegroundColor Green                
                $ErrorCheck = 1                
            }                    
            catch            
            {                
                $Err = "Error occured while updating Target Address"
                $ErrorCheck = 0 
                Failure                 
            }
        }
        #endregion

        #region-----------------EXTENSION ATTRIBUTE 12-------------# 
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        {
            $ErrorCheck = 0                                              
            try                    
            {     
                $ext12 = ((get-aduser $CWID).givenname).ToLower()+"."+((get-aduser $CWID).surname).ToLower()+"@******.com"                               
                Set-ADUser $CWID -Replace @{extensionAttribute12=($ext12)} -Verbose -ErrorAction Stop                          
                write-host ""                            
                write-host "Extension Attribute12 added.." -ForegroundColor Green                
                $ErrorCheck = 1                
            }                    
            catch            
            {                
                $Err = "Error occured while updating extension attribute12"   
                $ErrorCheck = 0 
                Failure                  
            }
        }
        #endregion

        #region-----------------CITY-------------# 
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        {  
            $ErrorCheck = 0                              
            try                    
            {                            
                Set-ADUser $CWID -Office ((Get-Aduser $CWID -Properties city).city)  -Verbose -ErrorAction Stop                         
                write-host ""                            
                write-host "City Updated.." -ForegroundColor Green  
                $ErrorCheck = 1  
                                 
            }                    
            catch            
            {                
                $Err = "Error occured while updating city"  
                $ErrorCheck = 0 
                Failure                    
            } 
        }

        #endregion

        #region-----------------MAILNICKNAME ATTRIBUTE-------------#  
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        {
            $ErrorCheck = 0                            
            try                    
            {                            
                Set-ADUser $CWID -Replace @{mailNIckname=($CWID)} -Verbose   -ErrorAction Stop                           
                write-host ""                            
                write-host "MailNickName Attribute Updated.." -ForegroundColor Green    
                $ErrorCheck = 1                    
            }                    
            catch            
            {                
                $Err = "Error occured while updating mailNickname attribute"     
                $ErrorCheck = 0 
                Failure                  
            }
        }

        #endregion

        #region-----------------ADDING USER TO GROUP-------------#   
        if($ErrorCheck -eq '1') #To make sure previous step is completed successfully or not   
        {
            $ErrorCheck = 0                                                 
            try
            {    
               'Groups names' | Add-ADGroupMember -Members $CWID -Verbose -ErrorAction Stop
               write-host "`nAdding user to Security Groups.." -ForegroundColor Green
               $ErrorCheck = 1                    
            }
            catch
            {
               $Err = "Error occured while Adding user into Security Groups"  
               $ErrorCheck = 0 
               Failure        
            }
        }

        #endregion

        #region-----------------MOVING TO SUB OU-------------#   
        if($ErrorCheck -eq '1')
        {
            $ErrorCheck = 0
            try
            {
            
               get-aduser $CWID |Move-ADObject -TargetPath $trgtOU -Verbose -ErrorAction Stop
               write-host "`nMoving user to Sub OU.." -ForegroundColor Green
               $ErrorCheck = 1                    
            }
            catch
            {
                $Err = "Error occured while moving user to Sub OU"     
                $ErrorCheck = 0
                Failure                   
            }     
        }
        #endregion

        #region---------------UpdatingArray_CSV-----------------#

        if($ErrorCheck -eq '1')
        {
            Write-Host "`n Done Processing $CWID..." 
            $Success += $CWID             
            foreach($Success_CWID in $Success)
            {      
               for($i=0;$i -lt $DataCount.count; $i++)    
               {            
                   if($Success_CWID -match $Users[$i].CWID)            
                   {                    
                       $Users[$i].Availability = 'Assigned'                     
                   }          
               }     
            }
            $Users | select CWID,Availability | Export-Csv $Path -NoTypeInformation  #### Creating_UpdatedCSV ####
        }

        #endregion

    }#foreach end

}#try end

catch{

    $Err = $Error[0].Exception
    Write-Warning "Something happened! $Err"
    Failure
}
finally{

    $ErrorActionPreference = "Continue"; #Reset the error action pref to default
}

#endregion
    
Stop-Transcript