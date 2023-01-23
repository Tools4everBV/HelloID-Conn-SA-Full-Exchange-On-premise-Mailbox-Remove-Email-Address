$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$alias = $form.gridmailbox.Alias 
$displayname = $form.gridmailbox.displayName
$UserPrincipalName = $form.gridmailbox.Userprincipalname
$DistinguishedName = $form.gridmailbox.DistinguishedName
$mailBoxToRemove = $form.emailAddress.leftToRight

# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    #-AllowRedirection
    $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
    $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
} catch {
    Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
    $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

try {
    $emailAddressList = $MailBoxToRemove.emailAddress
    foreach ($emailAddress in $emailAddressList) {        
        try{
            $ParamsSetMailbox = @{
                Identity       = $UserPrincipalName
                EmailAddresses = @{
                    remove = $emailAddress
                }
            }
            $null = Invoke-Command -Session $exchangeSession -ErrorAction Stop -ScriptBlock {
                Param ($ParamsSetMailbox)
                Set-Mailbox @ParamsSetMailbox
            } -ArgumentList $ParamsSetMailbox
            Write-Information "Successfully removed emailaddress [$emailAddress] for [$UserPrincipalName]"
            $Log = @{
                Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Successfully removed emailaddress [$emailAddress] for [$UserPrincipalName]." # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $emailAddress # optional (free format text) 
                TargetIdentifier  = $UserPrincipalName # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
        catch {
            Write-Error "Error removing emailaddresss [$emailAddress] for [$UserPrincipalName]. Error: $($_.Exception.Message)"
            $Log = @{
                Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Error removing emailaddresss [$emailAddress] for [$UserPrincipalName]." # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $emailAddress # optional (free format text) 
                TargetIdentifier  = $UserPrincipalName # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
    }
} catch {
    Write-Error "Error removing emailaddresses [$($emailAddressList -join ",")] for [$UserPrincipalName]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Error removing emailaddresses [$($emailAddressList -join ",")] for [$UserPrincipalName]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($emailAddressList -join ",") # optional (free format text) 
        TargetIdentifier  = $UserPrincipalName # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log    
}

# Disconnect from Exchange
try{
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
} catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
    $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
<#----- Exchange On-Premises: End -----#>

