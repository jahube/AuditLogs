# CONNECT Exchange Online
# [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Register-PSRepository -Default
# install-module ExchangeOnlineManagement -SkipPublisherCheck
# install-module -name PowershellGet -Force -SkipPublisherCheck
# Uninstall-Module PowershellGet -MaximumVersion "1.0.0.1" -Force -Confirm:$false -EA stop

 IF (!(get-accepteddomain -EA silentlycontinue)) { Connect-ExchangeOnline }

# $OfflineMode $false - [ONLINE]  collect Logs + parse
# $OfflineMode $true - [OFFLINE] XML-File-Dialog + Parse

$OfflineMode = $false

# desktop/MS-Logs+Timestamp

$ts = Get-Date -Format yyyyMMdd_hhmmss
$DesktopPath = ([Environment]::GetFolderPath('Desktop'))
$logsPATH = mkdir "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$ts"

# check PS Session + check Exo Module V2 (+ install if not found) + connect + $credentials
IF($OfflineMode -eq $false) {
IF(!@(Get-PSSession | where { $_.State -ne "broken" } )) {
IF(!@(Get-InstalledModule ExchangeOnlineManagement -ErrorAction SilentlyContinue)) { install-module exchangeonlinemanagement -Scope CurrentUser }

IF(!@($Credentials)) {$Credentials = Get-credential } ; IF(!@($ADMIN)) {$ADMIN = $Credentials.UserName }
Try { Connect-ExchangeOnline -Credential $Credentials -EA stop } catch { Connect-ExchangeOnline -UserPrincipalName $ADMIN } }

Start-Transcript "$logsPATH\Transcript_$ts.txt"
$FormatEnumerationLimit = -1

IF (!($Credentials.UserName -in (get-RoleGroupMember "Organization Management").primarySMTPaddress)) { Add-RoleGroupMember "Organization Management" -Member $ADMIN
Try { Connect-ExchangeOnline -Credential $Credentials -EA stop } catch { Connect-ExchangeOnline -UserPrincipalName $ADMIN } }

IF (!($USER)) { Try { $MBXs = Get-ExoMailbox -ResultSize unlimited -EA stop } catch { $MBXs = get-mailbox -ResultSize unlimited } # read mailboxes - try { ExoMBX } catch { classic }
IF ($MBXs.Count -gt "400") { $USER = Read-Host -Prompt "Affected User [Userprincipalname]" }                                                   # Above threshold - ask for manual user input
                      ELSE { $USER = @($MBXs | select Pr*ess,Dis*me,Use*me | Out-GridView -Passthru -Title "Select User").userprincipalname }} # below threshold - Out-gridview  Mailbox Select
IF (!($start)) { [int]$start = "-90" }
IF (!($data)) { $data = Search-MailboxAuditLog -Identity $user -ShowDetails -StartDate (get-date).AddDays($start) -EndDate (get-date) }

$data | Export-Clixml "$logsPATH\mailboxlogs.xml"

# check results BEFORE
$MBX = Get-Mailbox $user ; $PROW1 = $MBX.AuditOwner ; $PRDL1 = $MBX.AuditDelegate
 Write-host "BEFORE: AuditOwner = $($PROW1.count)" -foregroundcolor yellow;  Write-host "AuditOwner: $($PROW1)" -foregroundcolor Cyan
 Write-host "BEFORE: AuditOwner = $($PRDL1.count)" -foregroundcolor yellow;  Write-host "AuditOwner: $($PRDL1)" -foregroundcolor Cyan
 
# Apply ALL DETAILS
$Parameter = @{ identity = $user ; AuditEnabled = $true ;
AuditOwner = 'AddFolderPermissions', 'ApplyRecord', 'Create', 'Send', 'HardDelete', 'MailboxLogin', 'ModifyFolderPermissions', 'Move', 'MoveToDeletedItems', 'RecordDelete', 'RemoveFolderPermissions', 'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateCalendarDelegation', 'UpdateInboxRules' ;
AuditDelegate = 'AddFolderPermissions', 'ApplyRecord', 'Create', 'FolderBind', 'HardDelete', 'ModifyFolderPermissions', 'Move', 'MoveToDeletedItems', 'RecordDelete', 'RemoveFolderPermissions', 'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateInboxRules' ;
AuditAdmin = 'Copy', 'Create', 'HardDelete', 'MoveToDeletedItems', 'RecordDelete', 'RemoveFolderPermissions', 'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateCalendarDelegation', 'UpdateInboxRules' }
Set-Mailbox @Parameter 

# On /Off to refresh update
set-MailboxAuditBypassAssociation -Identity $user -AuditBypassEnabled $true  #OFF
set-MailboxAuditBypassAssociation -Identity $user -AuditBypassEnabled $false  #ON

# recheck results AFTER
$MBX = Get-Mailbox $user ; $PROW1 = $MBX.AuditOwner ; $PRDL1 = $MBX.AuditDelegate
 Write-host "AFTER: AuditOwner = $($PROW.count)" -foregroundcolor yellow;  Write-host "AuditOwner: $($PROW)" -foregroundcolor Cyan
 Write-host "AFTER: AuditOwner = $($PRDL.count)" -foregroundcolor yellow;  Write-host "AuditOwner: $($PRDL)" -foregroundcolor Cyan

# enable Unified Audit logs
IF(!((Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled)) {
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true ;
Write-host "Unified Audit log was disabled - ENABLING NOW" -F yellow }

get-mailbox $user | select AuditEnabled
get-mailbox $user | select -expandproperty auditadmin
get-mailbox $user | select -expandproperty auditdelegate
get-mailbox $user | select -expandproperty auditowner }

# Open File Dialog - Offline mode
Else { $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
   InitialDirectory = [Environment]::GetFolderPath('Desktop')
             Filter = 'XML Files (*.xml)|*.xml' }
              $null = $FileBrowser.ShowDialog()
              $Data = Import-Clixml $FileBrowser.FileName }

$data | select-object operation,clientprocessname,clientinfostring,ClientVersion,clientip -Unique | ft > "$logsPATH\Types-Unique.txt"

# deletion actions only
$softdelete = $data | where {$_.operation -eq "softdelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$harddelete = $data | where {$_.operation -eq "harddelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$MoveToDltd = $data | where {$_.operation -eq "MoveToDeletedItems" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList

# all actions
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip | select count,Name | Sort count -Descending | ft > "$logsPATH\Types-Summary.txt"

$harddelete | FT > "$logsPATH\harddelete-Details.txt"
$harddelete | Export-CSV "$logsPATH\harddelete-Details.csv" -NoTypeInformation
$harddelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\harddelete.txt"
$harddelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\harddelete-Summary.txt"

$softdelete | FT > "$logsPATH\softdelete-Details.txt"
$softdelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\softdelete.txt"
$softdelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\softdelete-Summary.txt"
$softdelete | Export-CSV "$logsPATH\softdelete-Details.csv" -NoTypeInformation

$MoveToDltd | FT > "$logsPATH\MoveToDltd-Details.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\MoveToDltd.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\MoveToDltd-Summary.txt"
$MoveToDltd | Export-CSV "$logsPATH\MoveToDltd-Details.csv" -NoTypeInformation

$data | FT > "$logsPATH\ALL-Details.txt"
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL.txt"
$data | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion,clientip | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL-Summary.txt"
$data | Export-CSV "$logsPATH\ALL-Details.csv" -NoTypeInformation

Stop-Transcript

Compress-Archive -Path $logsPATH -DestinationPath "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$($USER.replace('@',"-"))_$ts.Zip" -Force # Zip Logs
Invoke-Item $DesktopPath\MS-Logs # open file manager
$title = "Mailbox Audit Log Overview for [$user] from $((get-date).AddDays($start))"
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip | select count,Name | Sort count,Operation -Descending | Out-GridView -T $title