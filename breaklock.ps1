param(
    [Parameter(Mandatory=$true)][String]$ComputerName,
    [Switch]$Force
)
Import-Module CAT\LogonSession -ErrorAction Stop
$sessions = Get-LogonSession -ComputerName $ComputerName

function send_spam ($user, $ComputerName) {
    
    $mail = New-Object System.Net.Mail.MailMessage

    $date = (Get-Date).ToShortDateString()
    $time = (Get-Date).ToShortTimeString()

    $mail.From = "support@cat.pdx.edu"
    $mail.To.Add("$user@cecs.pdx.edu")
    $mail.Subject = "Your Windows Session has been Terminated"
    $mail.CC.Add("support@cat.pdx.edu")
    $mail.Headers.Add("X-TTS", "COMP")

     ##Body Paragraph of the Email, broken up for easier reading##
    $mail.Body = "At approximately $time, on $date, I found you to have locked the screen on $ComputerName." +
            " These machines are to be left open to other users as much as possible. If you need to step away" +
            " from the computer for more than 15 minutes, you are required to log out." +
            "`n`nIf you have a class project that requires you to be logged in for an extended period of time please run hiberfoo by completing the following steps:" +
            "`n`n1. Open My Computer`n2. In the address bar type: \\frost\programs\programs\hiberfoo`n3. Double Click on the Hiberfoo icon`n4. Click submit." +
            "`n`nIn the meantime, we have terminated your session on $ComputerName."
    
    $smtp = New-Object System.Net.Mail.SmtpClient("mailhost.cecs.pdx.edu")
    
   

    try{
        $smtp.Send($mail)
        #Send-MailMessage -SmtpServer $Smtp -From $From -To $To -Subject $Subject -Body $Body -Cc $CC
        Write-Host "`nBreaklock spam sent to $($mail.To)."
    } catch {
        Write-Error "`nSpam failed to send, please send manually."
    }
}

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
            Get-TSSession -ComputerName $ComputerName -UserName $session.UserName | Stop-TSSession -Force
            Write-Host "`nLock broken for $($session.UserAccount) on $ComputerName.`n"
            send_spam -user $($session.UserName) -ComputerName $ComputerName
        } catch {
            Write-Error "`nBreaklock failed for $($session.UserAccount) on $ComputerName!`n"
        }

    }
}

