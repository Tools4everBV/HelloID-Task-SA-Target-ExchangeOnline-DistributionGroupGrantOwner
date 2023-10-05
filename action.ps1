# HelloID-Task-SA-Target-ExchangeOnline-DistributionGroupGrantOwner
#################################################################
# Form mapping
$formObject = @{
    Identity = $form.GroupIdentity
    ManagedBy  = @{
       Add = $form.Owners.Name
    }
    BypassSecurityGroupManagerCheck = $true
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnline action: [DistributionGroupGrantOwner] for: [$($formObject.GroupIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Set-DistributionGroup', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    $null = Set-DistributionGroup @formObject -Confirm:$false  -ErrorAction 'Stop'

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.GroupIdentity
        TargetDisplayName = $formObject.GroupIdentity
        Message           = "ExchangeOnline action: [DistributionGroupGrantOwner][$($formObject.ManagedBy.Add -join ', ')] to group [$($formObject.GroupIdentity)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnline action: [DistributionGroupGrantOwner][$($formObject.ManagedBy.Add -join ', ')]  to group [$($formObject.GroupIdentity)] executed successfully"

} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.GroupIdentity
        TargetDisplayName = $formObject.GroupIdentity
        Message           = "Could not execute ExchangeOnline action: [DistributionGroupGrantOwner] for: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [DistributionGroupGrantOwner] for: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
#################################################################
