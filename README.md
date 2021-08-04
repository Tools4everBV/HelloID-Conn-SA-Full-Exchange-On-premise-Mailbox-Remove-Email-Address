# HelloID-Conn-SA-Full-Exchange-On-premise-Mailbox-Remove-Email-Address

<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides the functionality to remove an non primary emailaddress from a mailbox. The following options are available:
 1. Give a name to lookup a mailbox
 2. The result will show you a list of mailboxes. You will need select to correct one
 3. Select the email address you want to remove
 4. The selected email address will be removed

<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.


## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>ExchangeConnectionUri</td><td>********</td><td>Exchange server URI</td></tr>
  <tr><td>ExchangeAdminUsername</td><td>domain/user</td><td>Exchange server admin account</td></tr>
  <tr><td>ExchangeAdminPassword</td><td>********</td><td>Exchange server admin password</td></tr>
  <tr><td>ExchangeSearchOU</td><td>Example.com/Users</td><td>Exchange server OrganizationalUnit to search</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source '[powershell-datasource]_Exchange-mailbox-remove-email-address-get-mailbox'
This Powershell data source runs a query to search for the mailbox.

### Powershell data source '[powershell-datasource]_Exchange-mailbox-remove-email-address'
This Powershell data source runs a query to search for the existing emailaddresses of the selected mailbox.

### Delegated form task '[Exchange-mailbox-remove-email-address'
This delegated form task removes the Email Address from the mailbox

# HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/