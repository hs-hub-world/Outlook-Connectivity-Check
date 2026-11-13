
# Outlook Online / Offline Health Check
# For Classic Outlook (Windows)

$ErrorActionPreference = "Stop"

$result = [ordered]@{
    OutlookRunning   = $false
    InternetAccess   = $false
    ExchangeReachable = $false
    MailboxConnected = $false
    Status           = "UNKNOWN"
}

# 1. Check Outlook process
$result.OutlookRunning = [bool](
    Get-Process OUTLOOK -ErrorAction SilentlyContinue
)

# 2. Check internet connectivity
try {
    $null = Invoke-WebRequest `
        -Uri "https://www.microsoft.com" `
        -Method Head `
        -TimeoutSec 10

    $result.InternetAccess = $true
}
catch {
    $result.InternetAccess = $false
}

# 3. Check Microsoft Exchange Online connectivity
try {
    $null = Invoke-WebRequest `
        -Uri "https://outlook.office365.com" `
        -Method Head `
        -TimeoutSec 10

    $result.ExchangeReachable = $true
}
catch {
    # HTTP 401 / 403 can mean the server is reachable
    # even though anonymous access is denied.
    if ($_.Exception.Response) {
        $result.ExchangeReachable = $true
    }
}

# 4. Check actual Outlook MAPI connection
if ($result.OutlookRunning) {
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject(
            "Outlook.Application"
        )

        $namespace = $outlook.GetNamespace("MAPI")
        $null = $namespace.Folders.Count

        # ExchangeConnectionMode is an Outlook OlExchangeConnectionMode enum
        $mode = [int]$namespace.ExchangeConnectionMode

        # 0 = No Exchange connection
        # 1 = Offline
        # 2 = Online
        # 3 = Disconnected
        # 4 = Cached Connected (Headers)
        # 5 = Cached Connected (Full)
        # 6 = Cached Connected (Drizzle)

        $result.MailboxConnected = $mode -in @(2,4,5,6)
        $result.Status = if ($result.MailboxConnected) {
            "ONLINE"
        } else {
            "OFFLINE / DISCONNECTED (Mode $mode)"
        }
    }
    catch {
        $result.Status = "Unable to read Outlook connection state"
    }
}
else {
    $result.Status = "OUTLOOK NOT RUNNING"
}

[PSCustomObject]$result | Format-List
