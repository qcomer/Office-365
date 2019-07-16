$credential = Get-Credential
Connect-MsolService -Credential $credential
 
$userRequiringAccess = Read-Host -Prompt "Input the user's email that is requesting the access"
$accessRight = Read-Host -Promt "Level of rights being requested: Reviewer, Editor, etc."
$companyName = Read-Host -Prompt "Tenant Name"
  
$customers = Get-msolpartnercontract -All | Where-Object {$_.name -match $companyName}
foreach ($customer in $customers) {
    $InitialDomain = Get-MsolDomain -TenantId $customer.TenantId | Where-Object {$_.IsInitial -eq $true}
    Write-Host "Setting Calendar Permissions for $($customer.Name)" -ForegroundColor Green
    $DelegatedOrgURL = "https://outlook.office365.com/powershell-liveid?DelegatedOrg=" + $InitialDomain.Name
    $s = New-PSSession -ConnectionUri $DelegatedOrgURL -Credential $credential -Authentication Basic -ConfigurationName Microsoft.Exchange -AllowRedirection
    Import-PSSession $s -CommandName Get-Mailbox, Get-MailboxFolderPermission, Set-MailboxFolderPermission, Add-MailboxFolderPermission -AllowClobber
 
    $mailboxes = Get-mailbox
    $userRequiringAccess = Get-mailbox $userRequiringAccess
    foreach ($mailbox in $mailboxes) {
        $accessRights = $null
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -erroraction SilentlyContinue
         
        if ($accessRights.accessRights -notmatch $accessRight -and $mailbox.primarysmtpaddress -notcontains $userRequiringAccess.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
            Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
            try {
                Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
            }
            catch {
                Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
            } 
            $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress
            if ($accessRights.accessRights -match $accessRight) {
                Write-Host "Successfully added $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Green
            }
            else {
                Write-Host "Could not add $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Red
            }
        }else{
            Write-Host "Permission level already exists for $($userrequiringaccess.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
        }
    }
    Remove-PSSession $s
}