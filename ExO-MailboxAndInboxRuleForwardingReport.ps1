<#PSScriptInfo
.VERSION 1.0
.GUID 4e794dab-0e07-43ed-bbd9-ec685be3421a
.AUTHOR Soren Lindevang
.COMPANYNAME
.COPYRIGHT
.TAGS PowerShell Exchange Online Office 365 Reporting Report Forwarding ForwardingRules InboxRules SMTPForwarding Mailbox MailboxForwarding Azure Automation
.LICENSEURI
.PROJECTURI
.ICONURI
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
#>

<#
.SYNOPSIS 
    Generate Report of Mailbox and Inbox Rules Forwarding Mails to External Recipients.

.DESCRIPTION
    Searches for mailbox forwarding and inbox rules that forward to external recipients.
    
    Option to send a CSV report over email.
    
    Designed for execution in Azure Automation.

    Check out the GitHub Repo for more information: https://github.com/soren-cloud/ExO-MailboxAndInboxRuleForwardingReport

.PARAMETER AutomationPSCredentialName
    Name of the Automation Credential used when connecting to Exchange Online.

    The Account should at least have "Audit Log" rights in the Exchange Online tenant.

    Example: Exchange Online Service Account

.PARAMETER ExcludeExternalDomain
    External recipient domains to be excluded when searching for external forwarding.

    Example 1: ['domain.com']
    Example 2: ['domainA.com','domainB.com']

.PARAMETER ExcludeExternalEmailAddress
    External recipient email adresses to be excluded when searching for external forwarding.

    Example 1: ['abc@domain.com']
    Example 2: ['abc@domain.com','xyz@domain.com']

.PARAMETER SendMailboxForwardingReport 
     If this switch is present, the script sends an email with a CSV file attached, if mailbox forwarding is detected.
     
     If used, please do modify the 'SendMailReport' variables in the 'Declarations' area.


     Example 1: true
     Example 2: false

.PARAMETER SendInboxRuleForwardingReport 
     If this switch is present, the script sends an email with a CSV file attached, if any inbox rules are detected.
     
     If used, please do modify the 'SendMailReport' variables in the 'Declarations' area.


     Example 1: true
     Example 2: false

.INPUTS
    N/A

.OUTPUTS
    N/A

.NOTES
    Version:        1.0
    Author:         Soren Greenfort Lindevang
    Creation Date:  16.04.2018
    Purpose/Change: Initial script development
  
.EXAMPLE
    N/A
#>
[cmdletbinding()]
param (
    [Parameter(
        Mandatory=$true)]
        [string]$AutomationPSCredentialName,
    [Parameter(
        Mandatory=$false)]
        [string[]]$ExcludeExternalDomain,
    [Parameter(
        Mandatory=$false)]
        [string[]]$ExcludeExternalEmailAddress,
    [Parameter(
        Mandatory=$false)]
        [switch]$SendMailboxForwardingReport,
    [Parameter(
        Mandatory=$false)]
        [switch]$SendInboxRuleForwardingReport    
)


#-----------------------------------------------------------[Functions]------------------------------------------------------------

# Test if script is running in Azure Automation
function Test-AzureAutomationEnvironment
    {
    if ($env:AUTOMATION_ASSET_ACCOUNTID)
        {
        Write-Verbose "This script is executed in Azure Automation"
        }
    else
        {
        $ErrorMessage = "This script is NOT executed in Azure Automation."
        throw $ErrorMessage
        }
    }

# Connect to Exchange Online 
function Connect-ExchangeOnline 
    {
    param ($Credential,$Commands)
    try
        {
        Write-Output "Connecting to Exchange Online"
        Get-PSSession | Remove-PSSession       
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential `
            -Authentication Basic -AllowRedirection
        Import-PSSession -Session $Session -DisableNameChecking:$true -AllowClobber:$true -CommandName $Commands | Out-Null
        }
    catch 
        {
        Write-Error -Message $_.Exception
        throw $_.Exception
        }
    Write-Verbose "Successfully connected to Exchange Online"
    }

# Disconnect to Exchange Online 
function Disconnect-ExchangeOnline 
    {
    try
        {
        Write-Output "Disconnecting from Exchange Online"
        Get-PSSession | Remove-PSSession       
        }
    catch 
        {
        Write-Error -Message $_.Exception
        throw $_.Exception
        }
    Write-Verbose "Successfully disconnected from Exchange Online"
    }


Function Get-MailboxForwardingToExternal
    {
    [cmdletbinding()]
    param (
        [Parameter(
            Mandatory=$true)]
            [Alias('Mailboxes')]
            [psobject]$MailboxPSObject,
        [Parameter(
            Mandatory=$true)]
            [Alias('Recipients')]
            [psobject]$RecipientPSObject,
        [Parameter(
            Mandatory=$true)]
            [Alias('InternalDomains')]
            [psobject]$AcceptedDomainPSobject,
        [Parameter(
            Mandatory=$false)]
            [string[]]$ExcludeExternalDomain,
        [Parameter(
            Mandatory=$false)]
            [string[]]$ExcludeExternalEmailAddress
    )
    Begin
        {
        Write-Verbose "Beginning Get-Get-MailboxForwardingToExternal function"
        Write-Verbose "$(($MailboxPSObject | Measure-Object).Count) objects in 'MailboxPSObject'"
        }

    Process
        {
        Write-Verbose "Processing Objects in 'MailboxPSObject'"
        $count = $null
        foreach ($Mailbox in $MailboxPSObject) 
            {
            $count++
            $WriteOutput = $false
            Write-Verbose "Object $count of $(($MailboxPSObject | Measure-Object).Count)"
            Write-Verbose "PrimarySmtpAddress: '$($Mailbox.PrimarySmtpAddress)'"

            # ForwardingAddress
            if ($Mailbox.ForwardingAddress)
                {
                $ForwardingRecipient = $RecipientPSObject | Where-Object {$_.Identity -eq $Mailbox.ForwardingAddress}
                $ForwardingRecipientType = $ForwardingRecipient.RecipientType
                Write-Verbose "ForwardingAddress: $($ForwardingRecipient.Identity)"
                Write-Verbose "ForwardingAddress Type: $ForwardingRecipientType"
                if ($ForwardingRecipientType -eq "MailContact")
                    {
                    $EmailAddress = ($ForwardingRecipient.ExternalEmailAddress -split "SMTP:")[1].Trim("]")
                    $EmailAddressDomain = ($EmailAddress -split "@")[1]
                    if (($AcceptedDomains.DomainName -notcontains $EmailAddressDomain) -and 
                        ($ExcludeExternalEmailAddress -notcontains $EmailAddress) -and
                        ($ExcludeExternalDomain -notcontains $EmailAddressDomain))
                        {
                        $ForwardingHash = $null
                        $ForwardingHash = [ordered]@{
                            PrimarySmtpAddress              = $Mailbox.PrimarySmtpAddress
                            DisplayName                     = $Mailbox.DisplayName
                            ExternalForwardingAddress       = $EmailAddress
                            MailboxForwardingType           = "ForwardingAddress"
                            }
                        $Object = New-Object PSObject -Property $ForwardingHash
                        Write-Verbose "Writing Output"
                        Write-Output $Object
                        }
                    else
                        {
                        Write-Verbose "$EmailAddress - Internal Domain, Excluded Domain or Excluded Email Address"
                        }
                    }
                else
                    {
                    Write-Verbose "Skipping this type of recipient"
                    }
                }
            else
                {
                Write-Verbose "ForwardingAddress: null"
                }

            # ForwardingSmtpAddress
            if ($Mailbox.ForwardingSmtpAddress)
                {
                $MailboxForwardingSmtpAddress = $Mailbox.ForwardingSmtpAddress
                Write-Verbose "ForwardingSmtpAddress: $MailboxForwardingSmtpAddress"
                $EmailAddress = ($MailboxForwardingSmtpAddress -split "SMTP:")[1].Trim("]")
                $EmailAddressDomain = ($EmailAddress -split "@")[1]
                if (($AcceptedDomains.DomainName -notcontains $EmailAddressDomain) -and 
                    ($ExcludeExternalEmailAddress -notcontains $EmailAddress) -and
                    ($ExcludeExternalDomain -notcontains $EmailAddressDomain))
                    {
                    $ForwardingHash = $null
                    $ForwardingHash = [ordered]@{
                        PrimarySmtpAddress              = $Mailbox.PrimarySmtpAddress
                        DisplayName                     = $Mailbox.DisplayName
                        ExternalForwardingAddress       = $EmailAddress
                        MailboxForwardingType           = "ForwardingSmtpAddress"
                        }
                    $Object = New-Object PSObject -Property $ForwardingHash
                    Write-Verbose "Writing Output"
                    Write-Output $Object
                    }
                else
                    {
                    Write-Verbose "$EmailAddress - Internal Domain, Excluded Domain or Excluded Email Address"
                    }

                }
            else
                {
                Write-Verbose "ForwardingSMTPAddress: null"
                }
            }
        }

    End
        {
        Write-Verbose "End of Get-InboxRuleForwardingToExternal function"
        }
    }

Function Get-InboxRuleForwardingToExternal
    {
    [cmdletbinding()]
    param (
        [Parameter(
            Mandatory=$true)]
            [Alias('Mailboxes')]
            [psobject]$MailboxPSObject,
        [Parameter(
            Mandatory=$true)]
            [Alias('InternalDomain')]
            [psobject]$AcceptedDomainPSobject,
        [Parameter(
            Mandatory=$false)]
            [string[]]$ExcludeExternalDomain,
        [Parameter(
            Mandatory=$false)]
            [string[]]$ExcludeExternalEmailAddress
    )
    Begin
        {
        Write-Verbose "Beginning Get-InboxRuleForwardingToExternal function"
        Write-Verbose "$(($MailboxPSObject | Measure-Object).Count) objects in 'MailboxPSObject'"
        }

    Process
        {
        Write-Verbose "Processing Objects in 'MailboxPSObject'"
        $count = $null
        foreach ($Mailbox in $MailboxPSObject) 
            {
            $count++
            Write-Verbose "Object $count of $(($MailboxPSObject | Measure-Object).Count)"
            Write-Verbose "PrimarySmtpAddress: '$($Mailbox.PrimarySmtpAddress)'"
            $ForwardingRules = $null
            try
                {
                $Rules = Get-InboxRule -Mailbox $Mailbox.PrimarySmtpAddress
                }
            catch
                {
                Write-Error -Message $_.Exception
                }
            $ForwardingRules = $Rules | Where-Object {$_.ForwardTo -or $_.ForwardAsAttachmentTo}           
            foreach ($Rule in $ForwardingRules)
                {
                $Recipients = @()
                $Recipients = $Rule.ForwardTo | Where-Object {$_ -match "SMTP"}
                $Recipients += $Rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}
     
                $ExternalRecipients = @()
 
                foreach ($Recipient in $Recipients) 
                    {
                    $EmailAddress = ($Recipient -split "SMTP:")[1].Trim("]")
                    $EmailAddressDomain = ($EmailAddress -split "@")[1]
                    if (($ExcludeExternalEmailAddress -notcontains $EmailAddress) -and
                        ($AcceptedDomains.DomainName -notcontains $EmailAddressDomain) -and
                        $ExcludeExternalDomain -notcontains $EmailAddressDomain)
                        {
                        $ExternalRecipients += $EmailAddress
                        }
                    else
                        {
                        Write-Verbose "$EmailAddress - Internal Domain, Excluded Domain or Excluded Email Address"
                        }
                    }
                if ($ExternalRecipients) 
                    {
                    $ExternalRecipientsString = $ExternalRecipients -join ", "
                    Write-Verbose "Rule '$($Rule.Name)' forwards to '$ExternalRecipientsString'"
 
                    $RuleHash = $null
                    $RuleHash = [ordered]@{
                        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                        DisplayName        = $Mailbox.DisplayName
                        RuleId             = $Rule.Identity
                        RuleName           = $Rule.Name
                        RuleDescription    = $Rule.Description
                        ExternalRecipients = $ExternalRecipientsString 
                        }
                    $Object = New-Object PSObject -Property $RuleHash
                    Write-Verbose "Writing Output"
                    Write-Output $Object
                    }
                else
                    {
                    Write-Verbose "Rule '$($Rule.Name)' does not contain external forwarding"
                    }
                }
            }
        }

    End
        {
        Write-Verbose "End of Get-InboxRuleForwardingToExternal function"
        }
    }


#----------------------------------------------------------[Declarations]----------------------------------------------------------

# General Send Report Variables
$ReportSmtpServer = "smtp.office365.com"
$ReportSmtpPort = 587
$ReportSmtpPSCredentialName = $AutomationPSCredentialName
$ReportSmtpFrom = "serviceaccount@domain.com"
$ReportSmtpTo = "name@domain.com"

# SendInboxRuleForwardingReport Variables
$SendInboxRuleForwardingReportSubject = "Report: Inbox Rules with Forwarding to External Recipients"
$SendInboxRuleForwardingReportBody = "CSV file attached, containing Inbox Rule Information" 

# SendMailboxForwardingReport Variables
$SendMailboxForwardingReportSubject = "Report: Mailbox Forwarding with External Recipients"
$SendMailboxForwardingReportBody = "CSV file attached, containing Mailbox Forwarding Information" 


#-----------------------------------------------------------[Execution]-----------------------------------------------------------

# Check if script is executed in Azure Automation
Test-AzureAutomationEnvironment

Write-Output "::: Parameters :::"
Write-Output "AutomationPSCredentialName:    $AutomationPSCredentialName"
Write-Output "ExcludeExternalDomain:         $ExcludeExternalDomain"
Write-Output "ExcludeExternalEmailAddress:   $ExcludeExternalEmailAddress"
Write-Output "SendMailboxForwardingReport:   $SendMailboxForwardingReport"
Write-Output "SendInboxRuleForwardingReport: $SendInboxRuleForwardingReport"
Write-Output ""

# Get AutomationPSCredential
Write-Output "::: Connection :::"
try
    {
    Write-Output "Importing Automation Credential"
    $Credential = Get-AutomationPSCredential -Name $AutomationPSCredentialName -ErrorAction Stop
    }
catch 
    {
    Write-Error -Message $_.Exception
    throw $_.Exception
    }
Write-Verbose "Successfully imported credentials"

# Connect to Exchange Online
Connect-ExchangeOnline -Credential $Credential -Commands "Get-AcceptedDomain","Get-Mailbox","Get-InboxRule","Get-Recipient"
Write-Output ""

# Import Accepted Domains
try
    {
    Write-Verbose "Importing List of Accepted Domains"
    $AcceptedDomains = Get-AcceptedDomain -ErrorAction Stop
    }
catch 
    {
    Write-Error -Message $_.Exception
    throw $_.Exception
    }
Write-Verbose "Successfully Imported List of Accepted Domains"

# Import All Mailboxes
try
    {
    Write-Verbose "Importing List of Mailboxes"
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop `
        | Select-Object Identity,DisplayName,PrimarySmtpAddress,ForwardingAddress,ForwardingSmtpAddress
    }
catch 
    {
    Write-Error -Message $_.Exception
    throw $_.Exception
    }
Write-Verbose "Successfully Imported List of Mailboxes"

if (!$Mailboxes)
    {
    $ErrorMessage = "No Mailboxes Found!"
    throw $ErrorMessage
    }

# Import All Recipients
try
    {
    Write-Verbose "Importing List of Recipients"
    $Recipients = Get-Recipient -ResultSize Unlimited -ErrorAction Stop `
        | Select-Object Identity,RecipientType,ExternalEmailAddress
    }
catch 
    {
    Write-Error -Message $_.Exception
    throw $_.Exception
    }
Write-Verbose "Successfully Imported List of Recipients"


# Process Mailboxes
Write-Output "::: Analyzing Mailbox Forwarding:::"
Write-Output "Mailboxes in Scope for Search: $(($Mailboxes | Measure-Object).Count)"

$MailboxForwarding = Get-MailboxForwardingToExternal -MailboxPSObject $Mailboxes -RecipientPSObject $Recipients -AcceptedDomainPSobject $AcceptedDomains `
    -ExcludeExternalEmailAddress $ExcludeExternalEmailAddress -ExcludeExternalDomain $ExcludeExternalDomain

Write-Output "Mailbox Forwarding to External Recipients: $(($MailboxForwarding | Measure-Object).Count)"
Write-Output $($MailboxForwarding | fl)
Write-Output ""

# Inbox Rules
Write-Output "::: Analyzing Inbox Rules:::"
Write-Output "Mailboxes in Scope for Search: $(($Mailboxes | Measure-Object).Count)"

$RulesWithForwarding = Get-InboxRuleForwardingToExternal -MailboxPSObject $Mailboxes -AcceptedDomainPSobject $AcceptedDomains `
    -ExcludeExternalEmailAddress $ExcludeExternalEmailAddress -ExcludeExternalDomain $ExcludeExternalDomain

Write-Output "Inbox Rules Forwarding to External Recipients: $(($RulesWithForwarding | Measure-Object).Count)"
Write-Output $RulesWithForwarding
Write-Output ""

# Send Mail Report
if (($SendMailboxForwardingReport -and $MailboxForwarding) -or ($SendInboxRuleForwardingReport -and $RulesWithForwarding))
    {
    Write-Output "::: Send Mail Report :::"
    Write-Output "Importing Automation Credential"
    try
        {
        $ReportSmtpPSCredential = Get-AutomationPSCredential -Name $ReportSmtpPSCredentialName -ErrorAction Stop
        }
    catch 
        {
        Write-Error -Message $_.Exception
        throw $_.Exception
        }
    $ReportTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"
    if ($SendMailboxForwardingReport -and $MailboxForwarding)
        {
        Write-Output "Generate Mailbox Forwarding CSV file"
        try
            {
            $CSVFileName = "MailboxForwarding_" + $ReportTime + ".csv"
            $CSVFilePath = $env:TEMP + "\" + $CSVFileName
            $MailboxForwarding | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
            }
        catch 
            {
            Write-Error -Message $_.Exception
            }
        Write-Output "Send e-mail to '$ReportSmtpTo'"
        try
            {
            Send-MailMessage -To $ReportSmtpTo -From $ReportSmtpFrom -Subject $SendMailboxForwardingReportSubject `
                -Body $SendMailboxForwardingReportBody -BodyAsHtml -Attachments $CSVFilePath -SmtpServer $ReportSmtpServer `
                -Port $ReportSmtpPort -UseSsl -Credential $ReportSmtpPSCredential -ErrorAction Stop
            }
        catch 
            {
            Write-Error -Message $_.Exception
            }
        }
    if ($SendInboxRuleForwardingReport -and $RulesWithForwarding)
        {
        Write-Output "Generate Inbox Rule Forwarding CSV file"
        try
            {
            $CSVFileName = "InboxForwardingRules_" + $ReportTime + ".csv"
            $CSVFilePath = $env:TEMP + "\" + $CSVFileName
            $RulesWithForwarding | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
            }
        catch 
            {
            Write-Error -Message $_.Exception
            }
        Write-Output "Send e-mail to '$ReportSmtpTo'"
        try
            {
            Send-MailMessage -To $ReportSmtpTo -From $ReportSmtpFrom -Subject $SendInboxRuleForwardingReportSubject `
                -Body $SendInboxRuleForwardingReportBody -BodyAsHtml -Attachments $CSVFilePath -SmtpServer $ReportSmtpServer `
                -Port $ReportSmtpPort -UseSsl -Credential $ReportSmtpPSCredential -ErrorAction Stop
            }
        catch 
            {
            Write-Error -Message $_.Exception
            }
        }
    }
Write-Output ""

# Close Session
Disconnect-ExchangeOnline
Write-Output ""

# Script Completed
Write-Output "Script Completed"
