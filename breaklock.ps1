param(
    [Parameter(Mandatory=$true)][String]$ComputerName,
    [Switch]$Force
)
Import-Module CAT\LogonSession -ErrorAction Stop
$sessions = Get-LogonSession -ComputerName $ComputerName

if (-not $sessions) {
    Write-Host "$ComputerName is not loggedon."
} else {
    foreach ($session in $sessions) {
        $status = if ($session.Locked) { 'Locked' } else { $session.ConnectionState }
        $logondiff = ((Get-Date) - $($session.LoginTime)).TotalMinutes
        $lockdiff = if ($session.LockTime) { "{0:N0}" -f ((Get-Date) - $($session.LockTime)).TotalMinutes } else { "0" }
        Write-Host "Computer: $ComputerName"
        Write-Host "Loggedon user: $($session.UserAccount)"
        Write-Host "Status: $status"
        Write-Host "Logon date/time: $($session.LoginTime)"
        if ($session.Locked) { Write-Host "Lock duration: $($lockdiff)m" } else { Write-Host "Session is NOT locked." }
        # Ask for confirmation, if needed.
        if (-not $Force) {
            Write-Host "Confirm breaklock (type n to cancel)? " -NoNewLine
            $resp = (Read-Host).toLower()
            if ($resp.StartsWith('n')) {
                Write-Host "`nAborted.`n"
                continue
            }
        }
        try {
            $session.Logoff()
            Write-Host "`nLock broken for $($session.UserAccount) on $ComputerName.`n"
        } catch {
            Write-Error "`nBreaklock failed for $($session.UserAccount) on $ComputerName!`n"
        }
    }
}
