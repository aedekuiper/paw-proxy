[CmdletBinding()]
param (
    [ValidateSet("devamsterdamumc", "accamsterdamumc", "amsterdamumc")]
    [Parameter(Mandatory = $true)]
    [string] $TenantName
)

# Create new GUID to be used during requests
$clientId = [guid]::NewGuid().ToString()

# Set default proxy
$defaultProxy = 'PROXY 127.0.0.2:8080'

# Fetch Microsoft 365 endpoints of type 2 (Optimize + Allow traffic direct) 
$opts = @{
    Type = 2
    Instance = 'Worldwide'
    ClientRequestId = $clientId
    DefaultProxySettings = $defaultProxy
    TenantName = $TenantName
    LowerCase = $true
}

# Load the Get-PacFile script (install if needed)
if (-not (Get-InstalledScript Get-PacFile -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Get-PacFile script from PSGallery..."
    Install-Script -Name Get-PacFile -Scope CurrentUser -Force
}
# Get script location so we can call it. 
$getPacCommand = (Join-Path (Get-InstalledScript -Name get-pacfile).InstalledLocation 'Get-Pacfile.ps1')

# Get web service version info and prepend to the PAC file as a comment
$ws = "https://endpoints.office.com"
$version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientId)

$filename = ("$($opts.TenantName)_$($opts.Instance).pac").ToLower()


# Check the current version in the PAC file and compare it to the web service version
if (Test-Path -PathType Leaf -Path $filename) {
    Write-Host "Found existing PAC file at $($filename). Checking version..."
    $pacHeader = Get-Content -Path $filename | Select-Object -First 1
} else {
    Write-Host "No existing PAC file found at $($filename). A new one will be created."
    $pacHeader = ""
}

$pacVersion = $pacHeader -replace '.*version\s*([0-9]+).*','$1'

$wsVersion = $version.latest

if ($pacVersion -ne $wsVersion) {
    Write-Warning "PAC file version ($pacVersion) does not match latest web service version ($wsVersion)."
} else {
    Write-Host "PAC file version matches web service version: $wsVersion. Expect no changes besides any custom domain overrides."
}

# Call Get-PacFile to retrieve the PAC contentß
$pacContent = & $getPacCommand @opts

# Append your custom override domains (ProxyOverride list)
$customDomains = @(
    "*.githubusercontent.com", "*.azure.com", "*.azure.net", "*.microsoft.com", "*.windowsupdate.com",
    "*.microsoftonline.com", "*.microsoftonline.cn", "*.windows.net",
    "*.windowsazure.com", "*.windowsazure.cn", "*.azure.cn",
    "*.loganalytics.io", "*.applicationinsights.io", "*.vsassets.io",
    "*.azure-automation.net", "*.visualstudio.com", "portal.office.com",
    "*.aspnetcdn.com", "*.sharepointonline.com", "*.msecnd.net",
    "*.msocdn.com", "*.webtrends.com", "amsterdamumc.secretservercloud.eu", "*.surfconext.nl"
)

# Build shExpMatch clauses
$shClauses = ($customDomains | ForEach-Object { "shExpMatch(host, `"$($_)`")" }) -join " ||`n        "

# Define constants for loopback, localhost, intranet
$hostBypass = @"
    // Local/intranet bypass
    if (host == "localhost" || shExpMatch(host, "127.*") || isPlainHostName(host)) {
        return "DIRECT";
    }
"@

# Inject into PAC—wrap default direct-from-Get-PacFile logic
$injection = @"
    host = host.toLowerCase();
$hostBypass

    // Custom override domains
    if (
        $shClauses
    ) {
        return "DIRECT";
    }

"@

# Replace the host normalization in pacContent to include the injection
$pacText = $pacContent -replace 'host = host.toLowerCase\(\);', $injection

# Output to file
Set-Content -Path $filename -Value $pacText -Encoding UTF8

# Append version info as comment at the top of the file
Set-Content -Path $filename -Value ("// This pac file contains data from the $($version.instance) instance for serviceArea $($version.serviceArea) version $($version.latest)" + "`n" + (Get-Content $filename | Out-String))

Write-Host "PAC file generated at: $filename"



