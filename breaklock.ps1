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
            send_spam($($session.UserName), $ComputerName)
        } catch {
            Write-Error "`nBreaklock failed for $($session.UserAccount) on $ComputerName!`n"
        }
    }
}

##thunderburd##
function send_spam ($user, $ComputerName) 
{
    ##Date and time for the email message
    $date = (Get-Date).ToShortDateString()
    $time = (Get-Date).ToShortTimeString()
    
    ##Email settings
    $Smtp = "mailhost.cecs.pdx.edu"
    $From = "support@cat.pdx.edu"
    $To = "$user@cecs.pdx.edu"
    $Subject = "Your Windows Session has been Terminated"
    $CC = "support@cat.pdx.edu"
    
    ##Body Paragraph of the Email, broken up for easier reading##
    $Body = "At approximately $time, on $date, I found you to have locked the screen on $ComputerName."
    $Body += " These machines are to be left open to other users as much as possible. If you need to step away"
    $Body += " from the computer for more than 15 minutes, you are required to log out."
    $Body += "`n`nIf you have a class project that requires you to be logged in for an extended period of time please run hiberfoo by completing the following steps:"
    $Body += "`n`n1. Open My Computer`n2. In the address bar type: \\frost\programs\programs\hiberfoo`n3. Double Click on the Hiberfoo icon`n4. Click submit."
    $Body += "`n`nIn the meantime, we have terminated your session on $ComputerName."

    ##Try to send the email. If it fails, display an error message
    try{
        Send-MailMessage -SmtpServer $Smtp -From $From -To $To -Subject $Subject -Body $Body -Cc $CC
        Write-Host "`nBreaklock spam sent to $To."
    } catch {
        Write-Error "`nSpam failed to send, please send manually."
    }
}
