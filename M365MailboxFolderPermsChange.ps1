#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####

############################
## O365MailboxFolderPermsChange.ps1
##
## See Readme.txt
##
############################
#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
# Get-Command -module ITAutomator  ##Shows a list of available functions
############
$O365_PasswordXML   = $scriptDir+ "\O365_Password.xml"
############
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "IdentityFrom,IdentityTo,Folder,Permission,ApplyEvenIfDelegate" | Add-Content $scriptCSV
    "borg@rhonegroup.com,vuong@rhonegroup.com,Calendar,AvailabilityOnly,TRUE" | Add-Content $scriptCSV
	"borg@rhonegroup.com,vuong@rhonegroup.com,Contacts,None,TRUE" | Add-Content $scriptCSV
	######### Template
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
## ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
#$entries
$entriescount = $entries.count
##

####
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in O365"
Write-Host ""
#Write-Host "admin_username: $($Globals.admin_username)"
Write-Host ""
Write-Host "XML: $(Split-Path $O365_PasswordXML -leaf)"
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()

## ----------Connect/Save Password
$domain=@($entries[0].psobject.Properties.name)[0] # property name of first column in csv
$domain=$entries[0].$domain # contents of property
$domain=$domain.Split("@")[1]   # domain part
$connected_ok = ConnectExchangeOnline -domain $domain
if (-not ($connected_ok))
{ # connect failed
    Write-Host "[Not connected]"
} # connect failed
else
{ # connected OK
    Write-Host "--------------------"
    Write-Host "CONNECTED: $($PSCred.UserName)"
    Write-Host "--------------------"
    $processed=0
    $message="$entriescount Entries. Continue?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
    [int]$defaultChoice = 0
    $choiceRTN = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
    if ($choiceRTN -eq 1)
    { "Aborting" }
    else 
    { ## continue choices

    ### Connect to O365
 
    $choiceLoop=0
    $i=0        
    foreach ($x in $entries)
    {
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
            {
            $message="Process entry "+$i+"?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","Yes to &All","&No","No and E&xit")
            [int]$defaultChoice = 1
            $choiceLoop = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
            }
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
            {
            $processed++

		    ####### Start code for object $x
			$mbfrom = Get-Mailbox $x.IdentityFrom
			$folder = $x.Folder
			$perms = $x.Permission

            #$mbto = Get-Mailbox $x.IdentityTo
            if ($x.IdentityTo -eq "")
            {
                Write-Host "      IdentityTo is <blank>, removing IdentityTo [UNDER_CONSTRUCTION]"
            }
            elseif ($x.IdentityTo -eq "DEFAULT")
            {
                    Write-host "IdentityTo address '$($x.IdentityTo)' this is a keyword meaning 'anyone inside the org'"
                    $to_smtp = $x.IdentityTo
                    $to_name = $x.IdentityTo
            }
            else
            {
             
                ###### check the ForwardTo address
                ### mailbox?
                Try {
                    $OldPref = $global:ErrorActionPreference
                    $global:ErrorActionPreference = 'Stop'
                    $recip = Get-Recipient -identity $x.IdentityTo
                    $to_smtp = $recip.PrimarySmtpAddress
                    $to_name = $recip.Name
                }
                Catch {
                    Write-Warning "IdentityTo address '$($x.IdentityTo)' is not a known email address in this org - trying it anyway"
                    $to_smtp = $x.IdentityTo
                    $to_name = $x.IdentityTo
                    
                }
                Finally {
                    $global:ErrorActionPreference = $OldPref
                }
                ###### check the ForwardTo address 
            }
			####### Display 'before' info
			$y= foreach ($del in $mbfrom.GrantSendOnBehalfTo) {"["+$del+"] "}
			Write-host "From: $mbfrom ($folder) Delegates: $y"
			Write-host "  To: $to_name   ($perms)"
			####
			Write-host "[Permission Before]"
			$getperm= Get-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($to_smtp) -ErrorAction silentlycontinue
			$getperm |Format-Table Foldername,User,AccessRights
			####
			if ($mbfrom.GrantSendOnBehalfTo -contains $to_name)
				{ #is delegate
				if ($x.ApplyEvenIfDelegate -eq $true)
					{ 
					Write-host "WARNING: THIS PERSON IS A DELEGATE and the ApplyEvenIfDelegate settings is TRUE.  MAKING CHANGES." -foregroundcolor Yellow 
					$made_changes = $true
					}
				else
					{
					Write-host "WARNING: THIS PERSON IS A DELEGATE and the ApplyEvenIfDelegate settings is FALSE.  NOT MAKING ANY CHANGES."  -foregroundcolor Yellow 
					$made_changes = $false
					}
				} #is delegate
			else
				{ #not a delegate
				$made_changes = $true
				} #not a delegate
			if ($made_changes)
				{ #make changes
				## see if user is listed
				if ($getperm)
					{#modify existing
					if ($getperm.AccessRights[0] -eq $perms)
						{
						Write-Host "[No change required]" -ForegroundColor Green
						$made_changes=$false
						}
					else
						{
						Set-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($to_smtp) -AccessRights $perms | Out-Null
						}
					}
				else
					{#no existing perms
					Add-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($to_smtp) -AccessRights $perms | Out-Null
					}
				##
				} #make changes  
			if ($made_changes)
				{
				####
                Write-Host "[Changes made]" -ForegroundColor Yellow
				Write-host "[Permission After]"
				$getperm= Get-MailboxFolderPermission ($mbfrom.PrimarySmtpAddress +":\$folder") -User ($to_smtp) -ErrorAction silentlycontinue
				$getperm |Format-Table Foldername,User,AccessRights
				####
				}
			####### End code for object $x
            ####### End code for object $x
            }
        if ($choiceLoop -eq 2)
            {
            write-host ("Entry "+$i+" skipped.")
            }
        if ($choiceLoop -eq 3)
            {
            write-host "Aborting."
            break
            }
        }
    } ## continue choices
    WriteText "Removing any open sessions..."
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    WriteText "------------------------------------------------------------------------------------"
    $message ="Done. " +$processed+" of "+$entriescount+" entries processed. Press [Enter] to exit."
    WriteText $message
    WriteText "------------------------------------------------------------------------------------"
	#################### Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    #################### Transcript Save
} #creds entered
PauseTimed -quiet 3 #$message
Pause