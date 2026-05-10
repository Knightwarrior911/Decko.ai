param([string]$Url)
$ProgressPreference = 'SilentlyContinue'
try {
    $r = Invoke-WebRequest -Uri $Url -UseBasicParsing `
        -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' `
        -TimeoutSec 30
    Write-Host "Status: $($r.StatusCode)"
    Write-Host "Length: $($r.Content.Length)"
    if ($r.Content.Length -gt 0) {
        $n = [Math]::Min(400, $r.Content.Length)
        Write-Host "First $n chars:"
        Write-Host $r.Content.Substring(0, $n)
    }
} catch {
    Write-Host "ERROR: $_"
}
