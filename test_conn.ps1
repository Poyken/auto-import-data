$servers = @(
    "dbserver.hycap.co.kr,5398",
    "dbserver.hycap.co.kr",
    "dbserver.hycap.co.kr,1433"
)

$pwd = "vina1234%6&8"
$user = "vinaadmin"

Write-Host "=== TEST DATABASE CONNECTION ===" -ForegroundColor Cyan
try {
    Write-Host "Public IP: $((Invoke-RestMethod -Uri 'https://ifconfig.me/ip' -UseBasicParsing).Trim())" -ForegroundColor Yellow
} catch {
    Write-Host "Public IP: (cannot check)"
}
Write-Host "--------------------------------"

foreach ($srv in $servers) {
    Write-Host "Testing connection to: $srv" -ForegroundColor White
    $connString = "Server=$srv;Database=SmartFactoryV2;User ID=$user;Password=$pwd;TrustServerCertificate=True;Connection Timeout=5;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $conn.Open()
        $stopwatch.Stop()
        Write-Host "[SUCCESS] Connected to $srv in $($stopwatch.ElapsedMilliseconds) ms!" -ForegroundColor Green
        $conn.Close()
        break
    } catch {
        $stopwatch.Stop()
        Write-Host "[FAILED] Cannot connect to $srv." -ForegroundColor Red
        $errMsg = $_.Exception.Message
        if ($errMsg -match "wait operation timed out" -or $errMsg -match "Error: 258") {
            Write-Host "   -> Reason: Timeout (Blocked by Firewall or Server is offline / wrong port)`n" -ForegroundColor Gray
        } else {
            Write-Host "   -> Reason: $errMsg`n" -ForegroundColor Gray
        }
    }
}

Write-Host "--------------------------------"
Write-Host "[DIAGNOSTICS & FIXES]" -ForegroundColor Cyan
Write-Host "If all methods returned Timeout (Error 258), your network (IP 192.168.25.64) cannot reach the DB Server (211.34.149.156)."
Write-Host "1. Check if you need to connect to a VPN (Pulse Secure / FortiClient / OpenVPN)."
Write-Host "2. Your Public IP is blocked by the DB Firewall. Ask IT (Korea) to whitelist your Public IP on port 5398."
Write-Host "3. If the DB is local to your factory (Vinatech), 'dbserver.hycap.co.kr' might need to be resolved via hosts file to a local IP (e.g. 192.168.25.156) instead of the internet."
