$env = Get-Content .\env.json | ConvertFrom-Json
$identities = $env.identities # Shared mailboxes to scan
$recipient = $env.recipient # Address to send reminder to

while (1) {
    [datetime]$run_time = $env.run_time # Time to run
    [datetime]$current_time = Get-Date
    $wait_time = ($run_time - $current_time).TotalSeconds + 5 # Time until run time
    if ($wait_time -lt 0) {
        $wait_time += 86400 # If run time is in the past, add 1 day
    }
    Write-Output "Waiting for $($wait_time/60) minutes"
    Start-Sleep $wait_time # Wait for amount of time until run

    # Code:
    Write-Output "`n`nRunning on $(Get-Date)"
    $filterDate = (Get-Date).AddDays(-2) # 48 hours ago

    Add-Type -assembly "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    Write-Output "Connected to Outlook"

    foreach ($i in $identities) {
        $inbox = $namespace.GetSharedDefaultFolder($namespace.CreateRecipient($i), [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
        [array]$emails = @()
        $emails = $inbox.Items | Where-Object {($_.ReceivedTime -le $filterDate) -and ($_.SenderEmailType -ne "EX")}
        Write-Output "Found emails for $i"
        
        $msg = $outlook.CreateItem(0)
        $msg.To = $recipient
        if ($emails.Count -eq 0) {
            $msg.Subject = "$i is all caught up!"
        } else {
            $msg.Subject = "$i - $($emails.Count) unresponded email(s)" 
        }
        $msg.Body = ""
        foreach ($e in $emails) {
            $msg.Body += "From $($e.SenderName) ($($e.SenderEmailAddress)) - Subject: $($e.Subject)`rSent at $($e.ReceivedTime.ToString())`rAssigned Category: $($e.Categories)`n`n"
        }
        $msg.Body += "This message was sent by a robot. beep"

        $msg.Send()
        Write-Output "Sent reminder to $recipient for $i"
    }
}