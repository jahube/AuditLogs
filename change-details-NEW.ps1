      start-transcript -verbose

$user = ""  # modify AFFECTED@USER.com

# desktop/MS-Logs+Timestamp

$ts = Get-Date -Format yyyyMMdd_hhmmss
$DesktopPath = ([Environment]::GetFolderPath('Desktop'))
$logsPATH = mkdir "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$ts"

Start-Transcript "$logsPATH\Transcript_$ts.txt"
$FormatEnumerationLimit = -1

# check PS Session + check Exo Module V2 (+ install if not found) + connect + $credentials

IF(!@(Get-PSSession | where { $_.State -ne "broken" } )) {
IF(!@(Get-InstalledModule ExchangeOnlineManagement -ErrorAction SilentlyContinue)) { install-module exchangeonlinemanagement -Scope CurrentUser }

IF(!@($Credentials)) {$Credentials = Get-credential } ; IF(!@($ADMIN)) {$ADMIN = $Credentials.UserName }
Try { Connect-ExchangeOnline -Credential $Credentials -EA stop } catch { Connect-ExchangeOnline -UserPrincipalName $ADMIN } }

IF (!($Credentials.UserName -in (get-RoleGroupMember "Organization Management").primarySMTPaddress)) { Add-RoleGroupMember "Organization Management" -Member $ADMIN
Try { Connect-ExchangeOnline -Credential $Credentials -EA stop } catch { Connect-ExchangeOnline -UserPrincipalName $ADMIN } }

Try {$All = Get-ExoMailbox -ResultSize unlimited -EA stop } catch { $All = get-mailbox -ResultSize unlimited } 
 IF ($All.Count -gt "400") { $Data = Read-Host -Prompt "Affected User [Userprincipalname]" }                      # Above threshold - ask for manual user input
                      ELSE { $Data = @($All | select Pr*ess,Dis*me,Use*me | Out-GridView -Passthru -Title "Select User").userprincipalname }} # below threshold

# enable Unified Audit logs
IF(!((Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled)) {
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true ;
Write-host "Unified Audit log was disabled - ENABLING NOW" -F yellow }

Foreach ($U in $Data) { Try { $M = Get-ExoMailbox $U -EA stop } catch { $M = get-mailbox $U }

 Write-host "BEFORE: [AuditOwner] $(@($M.AuditOwner).count) [AuditOwner]  $($M.AuditOwner)" -F yellow
 Write-host "BEFORE: [ Delegate ] $(@($M.AuditDelegate).count) [ Delegate ]  $($M.AuditDelegate)" -F yellow
 Write-host "BEFORE: [AuditAdmin] $(@($M.AuditAdmin).count) [ Delegate ]  $($M.AuditAdmin)" -F yellow

        $All = 'Create', 'HardDelete', 'MoveToDeletedItems', 'RecordDelete', 'RemoveFolderPermissions','SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateInboxRules'
$OWNRandDLGT =  'AddFolderPermissions', 'ApplyRecord', 'ModifyFolderPermissions','Move','UpdateCalendarDelegation' ;  $Owner = 'Send', 'MailboxLogin'
   $Delegate = 'FolderBind', 'SendAs', 'SendOnBehalf' ;  $Admin = 'Copy', 'SendAs', 'SendOnBehalf', 'UpdateCalendarDelegation'

                       # Apply ALL DETAILS
         $Parameter = @{ identity = $user ;  
                     AuditEnabled = $true ;
                       AuditOwner = $All + $OWNRandDLGT + $Owner
                    AuditDelegate = $All + $OWNRandDLGT + $Delegate
                       AuditAdmin = $All + $Admin }
                      Set-Mailbox @Parameter 

                      Set-Mailbox -Identity $M.UserPrincipalname -AuditEnabled $false        #OFF
                      Set-Mailbox -Identity $M.UserPrincipalname -AuditEnabled $true         #ON = update refresh Unified Audit
set-MailboxAuditBypassAssociation -Identity $M.UserPrincipalname -AuditBypassEnabled $true   #OFF
set-MailboxAuditBypassAssociation -Identity $M.UserPrincipalname -AuditBypassEnabled $false  #ON = update refresh Mbx Audit

 Try { $M = Get-ExoMailbox $U -EA stop } catch { $M = get-mailbox $U }

 Write-host "AFTER: [AuditOwner] $(@($M.AuditOwner).count) [AuditOwner]  $($M.AuditOwner)" -F green
 Write-host "AFTER: [ Delegate ] $(@($M.AuditDelegate).count) [ Delegate ]  $($M.AuditDelegate)" -F green
 Write-host "AFTER: [AuditAdmin] $(@($M.AuditAdmin).count) [ Delegate ]  $($M.AuditAdmin)" -F green

 } # End Foreach per User

stop-transcript

#End Script