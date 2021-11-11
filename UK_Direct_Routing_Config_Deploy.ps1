############################################

# Teams UK Standard Direct Routing Config Deployment #
# Description: Deploys a standard configuration for Dial Plans, 
# PSTN Usages, Voice Routes and Voice Routing Policies

# Author: James Storr

############################################


#Import-Module MicrosoftTeams and add any other pre-reqs
Add-Type -AssemblyName System.Windows.Forms
if(!(get-installedmodule -name MicrosoftTeams))
{
    Write-Host "MicrosoftTeams Powershell Module Missing. Please Install by running 'Install-Module -Name MicrosoftTeams' as Administrator"
    Return
}

if(!(get-installedmodule -name ImportExcel))
{
    Write-Host "ImportExcel Powershell Module Missing. Please Install by running 'Install-Module -Name ImportExcel' as Administrator"
    Return
}

Connect-MicrosoftTeams

Write-Host "###### Teams Direct Routing  ######"

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FileBrowser.ShowDialog()
$File = $FileBrowser.FileName
$DialPlan = Import-Excel -Path $File -WorksheetName "DialPlan"
$PSTNUsage = Import-Excel -Path $File -WorksheetName "PSTNUsage"
$VoiceRoute = Import-Excel -Path $File -WorksheetName "VoiceRoute"
$VoicePolicy = Import-Excel -Path $File -WorksheetName "VoicePolicy"
$MaxRetries = 5

$PSTNGateway = Read-Host -Prompt "Enter the PSTN Gateway (SBC)"

#Check if Dial Plan Already Exists. If it doesn't, create one
if(!(Get-CsTenantDialPlan -identity "UK-DialPlan" -ErrorAction 'silentlycontinue'))
    {
        Write-Host ("Dial Plan Doesn't Exist. creating new Dial Plan UK-DialPlan")
        New-CsTenantDialPlan -Identity "UK-DialPlan" -SimpleName "UK-DialPlan"
    }
    else {
        Write-Host "A Dial Plan Already Exists for UK-DialPlan"
    }

$DPCompleted = $false;  
$RetryCount = 0  
#Add the Rules into the Dial Plan. 
while(-not $DPCompleted)
{
    try {

        foreach($rule in $DialPlan)
            {
                #Get list of Normalization rules within each dial plan
                $CheckDP = Get-CsTenantDialPlan -Identity "UK-DialPlan"
                $DPRules = $CheckDP.NormalizationRules

                #Check whether the normalization rule exists within the dial plan
                If(!($DPRules | Where-Object {$_.Name -eq $rule.'Normalization Rule'}))
                    {
                        $RuleName = $rule.'Normalization Rule' 
            
                        Write-Host "Rule Doesn't Exist, creating new rule called $RuleName in UK-DialPlan"

                        #If no rule exists, create the rule within the dial plan
                        $Newrule = New-CsVoiceNormalizationRule -Parent "UK-DialPlan" -Description $RuleName -Pattern $Rule.Pattern -Translation $Rule.Translation -Name $RuleName -IsInternalExtension $False -InMemory
                        Set-CsTenantDialPlan -Identity "UK-DialPlan" -NormalizationRules @{add=$newRule}
                    }
                else{
                        $RuleName = $Rule.Name 
                        Write-host "A Rule already exists for $Rulename in DialPlan UK-DialPlan"
                    }
            }
            $DPCompleted = $true
        
        }
    catch 
        {
            if($RetryCount -ge $MaxRetries)
            {
                Write-Verbose "Max Number of Retries for Adding Dial Plan has been exceeded. Quitting"
                throw
            }
            else {
                $RetryCount++
                Write-Host "Error Occured. Retrying in 5 seconds"
                Start-Sleep -Seconds 5
            }
        }
}


#Get list of existing PSTN Usages and Create new ones if they don't already exist
$PSTNCompleted = $false;  
$RetryCount = 0 
while(-not $PSTNCompleted)
{
    try {
        $CheckPSTN = Get-CsOnlinePstnUsage
        foreach($Usage in $PSTNUsage)
       {
                  
            #Check whether the PSTN Usage exists within the Tenant
            If (!($CheckPSTN | Where-Object {$_.usage -eq $Usage.'PSTN Usage'}))
            {
                $CreateUsage = $Usage.'PSTN Usage'               
                Write-Host "Creating new PSTN Usage named: $CreateUsage"

                #If no PSTN Usage exists, create the PSTN Usage under Global
                Set-CsOnlinePstnUsage -Identity Global -Usage @{add=$Usage.'PSTN Usage'}
                                
            }
            else{
                $CreateUsage = $Usage.'PSTN Usage'
                Write-host "A PSTN Usage already exists named: $CreateUsage"
            }
        }
        $PSTNCompleted = $true
    }
    catch {
        if($RetryCount -ge $MaxRetries)
        {
            Write-Host "Max Number of Retries for Adding PSTN Usage has been exceeded. Quitting."
        }
        else {
            Write-Warning $Error[0]
            Write-Host "Error Occured. Retrying in 5 seconds"
            Start-Sleep -Seconds 5
            $RetryCount++
        }
    }
}


 #Get list of existing Voice routes and create new ones if they don't already exist

$VRCompleted = $false;  
$RetryCount = 0 

while(-not $VRCompleted)
{
    try {
        $CheckVoiceRoute = Get-CsOnlineVoiceRoute
        foreach($VR in $VoiceRoute)
        {
            #Check whether the voice route exists within the tenant
            If(!($CheckVoiceRoute | Where-Object {$_.Identity -eq $VR.'Voice Route'}))
            {
            
                $CreateVR = $VR.'Voice Route'               
                Write-Host "Creating new Voice Route named: $CreateVR"
                Write-Host "Adding Gateway $PSTNGateWay"

                #If no voice route exists, create the voice route (PSTNGateway selected automatically depending on country/region)
                New-CsOnlineVoiceRoute -Identity $VR.'Voice Route' -NumberPattern $VR.'Pattern' -OnlinePstnGatewayList @{add="$PSTNGateway"} -OnlinePstnUsages $VR.'PSTN Usage'

            }               
            else{
                $CreateVR = $VR.'Voice Route'
                Write-host "A Voice Route already exists named: $CreateVR"
            }
        }       
        $VRCompleted =$true
    }
    catch {
        if($RetryCount -ge $MaxRetries)
        {
            Write-Host "Max Number of Retries for Adding Voice Routes has been exceeded. Quitting."
            throw
        }
        else {
            Write-Warning $Error[0]
            Write-Host "Error Occured. Retrying in 5 seconds"
            Start-Sleep -Seconds 5
            $RetryCount++
        }
    }
}

#Get list of existing voice routing policies and create them if they don't already exist
$VPCompleted = $false;  
$RetryCount = 0 
While(-not $VPCompleted)
{
    try {
        $CheckVoicePolicy = Get-CsOnlineVoiceRoutingPolicy
        $Headers = $VoicePolicy | Get-Member -MemberType NoteProperty

        Foreach($VP in $VoicePolicy)
        {
            #Check whether the voice routing policy exists within the tenant
            If(!($CheckVoicePolicy | Where-Object {$_.Identity -eq "Tag:"+$VP.'Voice Policy'}))
            {
                $CreateVP = $VP.'Voice Policy'
                Write-Host "Creating new Voice Routing Policy named: $CreateVP"

                #If no voice routing policy exists, create a new Policy
                New-CsOnlineVoiceRoutingPolicy -Identity $VP.'Voice Policy'
            }
            else{
                $CreateVP = $VP.'Voice Policy'
                Write-Host "A Voice Routing Policy already exists named: $CreateVP"
            }
                    
            #For each voice policy, check and add each PSTN Usage assigned in Pivot Table (VoicePolicy)
            Foreach($header in $Headers)
            {
                $HeaderName = $header.Name.ToString()
                if($VP.$HeaderName -eq 1)
                { 
                $Usage = "UK-"+$HeaderName
                $VPName = $VP.'Voice Policy'
                
                If(!(get-CSonlineVoiceRoutingPolicy -Identity $VPName| Where-Object {$_.OnlinePstnUsages -Contains $Usage}))
                {
                    Write-Host "Adding $Usage to $VPName"
                    Set-CsOnlineVoiceRoutingPolicy -Identity $VPName -OnlinePstnUsages @{add=$Usage}
                }
                else {
                    Write-Host "PSTN Usage $Usage already exists in $VPName"
                } 
                
                }
            
            }
        }
        $VPCompleted = $true;
    }
    catch {
        if($RetryCount -ge $MaxRetries)
        {
            Write-Verbose "Max Number of Retries for Adding Voice Policies has been exceeded. Quitting."
            throw
        }
        else {
            Write-Warning $Error[0]
            Write-Host "Error Occured. Retrying in 5 seconds"
            Start-Sleep -Seconds 5
            $RetryCount++
        }
    }
}
