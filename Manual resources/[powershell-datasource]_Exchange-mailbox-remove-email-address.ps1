<#----- Exchange On-Premises: [powershell-datasource]_Exchange-mailbox-add-email-address-get-mailbox -----#>
# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos -ErrorAction Stop
    #-AllowRedirection
    # $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]"
} catch {
    Write-Information "Error connecting to Exchange using the URI [$exchangeConnectionUri]"
    Write-Information "Failed to connect to Exchange using the URI [$exchangeConnectionUri]"
    Write-Error "$($_.Exception.Message)"
    throw $_
}

try {
    $ParamsGetMailbxox = @{
        Identity = $datasource.mailBox.UserPrincipalName
    }

    Write-Information "SearchQuery Identity -eq $($ParamsGetMailbxox.Identity)"
    $mailBoxes = Invoke-Command -Session $exchangeSession -ScriptBlock {
        Param ($ParamsGetMailbxox)
        Get-Mailbox @ParamsGetMailbxox
    } -ArgumentList $ParamsGetMailbxox

    $mailboxes = $mailBoxes | Select-Object -ExpandProperty emailAddresses | Where-Object { $_ -clike "smtp*" }
    $resultCount = @($mailboxes).Count
    Write-Information "Result count: $resultCount"
    if ($resultCount -gt 0) {
        foreach ($mailbox in $mailboxes) {
            $returnObject = @{
                emailAddress = $mailbox.replace("smtp:", "")
            }
            Write-Output $returnObject
        }
    }
} catch {
    Write-Error "Error searching AD user [$searchValue]. Error: $($_.Exception.Message)"
}

# Disconnect from Exchange
try {
    Remove-PSSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange"
} catch {
    Write-Error "Error disconnecting from Exchange"
    Write-Error "$($_.Exception.Message)"
    throw $_
}
<#----- Exchange On-Premises: End -----#>

