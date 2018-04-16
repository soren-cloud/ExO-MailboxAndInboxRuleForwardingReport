ExO-MailboxAndInboxRuleForwardingReport
===========
Generate Report of Mailbox and Inbox Rules Forwarding Mails to External Recipients in an Exchange Online tenant.

This script was developed as part of a blog [article] on [soren.cloud].

*Note: This script is designed for execution in an Azure Automation runbook!*


## Requirements
* Office 365 tenant with Exchange Online mailboxes 
 
* Azure Subscription


## Prerequisites
See [prerequisites] section in this [article].

## Usage
Copy the content of the script into a Azure Automation PowerShell Runbook. Then test and deploy (schedule) :-)

**Disclaimer: No warranties. Use at your own risk.**

## Parameters
* **-AutomationPSCredentialName**, Name of the Automation Credential used when connecting to Exchange Online.
* **-ExcludeExternalDomain**, External recipient domains to be excluded when searching for external forwarding.
* **-ExcludeExternalEmailAddress**, External recipient email adresses to be excluded when searching for external forwarding.
* **-SendMailboxForwardingReport**, If this switch is present, the script sends an email with a CSV file attached, if mailbox forwarding is detected.
* **-SendInboxRuleForwardingReport**, If this switch is present, the script sends an email with a CSV file attached, if any inbox rules are detected.

## Examples
*Remember: This script is designed for execution in a Azure Automation runbook!*

`AutomationPSCredentialName: Exchange Online Service Account`, `ExcludeExternalDomain: ['domainA.com']`, `ExcludeExternalEmailAddress: ['abc@domain.com']`, `SendMailboxForwardingReport: true`, `SendInboxRuleForwardingReport: true`

Connect with service account 'Exchange Online Service Account', Exclude domain 'domainA.com' and address 'abc@domain.com' from the search. Send a report if Mailbox Forwarding and/or Inbox Rule Forwarding is present.

`AutomationPSCredentialName: Exchange Online Service Account`, `ExcludeExternalDomain: ['domainA.com','domainB.com']`, `ExcludeExternalEmailAddress: ['abc@domain.com','xyz@domain.com']`, `SendMailboxForwardingReport: false`, `SendInboxRuleForwardingReport: true`

Connect with service account 'Exchange Online Service Account'. Exclude domains 'domainA.com' and 'domainB.com' from the search. Also exclude addresses 'abc@domain.com' and 'xyz@domain.com'. Only send a report if Inbox Rule Forwarding is present.

## More Information
Article: <http://soren.cloud/>


## Credits
Written by: Søren Lindevang

Find me on:

* My Blog: <http://soren.cloud/>
* Twitter: <https://twitter.com/SorenLindevang>
* LinkedIn: <https://www.linkedin.com/in/lindevang/>
* GitHub: <https://github.com/soren-cloud>

[article]: http://soren.cloud/o365-secure-score-azure-automation-part-2-external-forwarding-report/
[my blog]: http://soren.cloud/
[soren.cloud]: http://soren.cloud/
[prerequisites]: http://soren.cloud/o365-secure-score-azure-automation-part-2-external-forwarding-report/#Prerequisites