# Pins this PC's current LAN/Wi-Fi address, so phones keep using the same https://<IP>:3000.
# -Undo hands the address back to the router (DHCP). Called by "Dat IP co dinh.bat" / "Tra lai IP tu dong.bat".
# ASCII only: Windows PowerShell 5.1 reads a .ps1 without BOM in the ANSI code page.
param([switch]$Undo)
$ErrorActionPreference = 'Stop'

# The adapter that is up and has a gateway: the Wi-Fi or cable to the router, not VPNs or virtual adapters.
$config = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1
if (-not $config) {
    Write-Host 'Khong tim thay mang dang ket noi toi router. Hay ket noi wifi roi chay lai.'
    exit 1
}
$index = $config.InterfaceIndex
$ip = @($config.IPv4Address)[0].IPAddress
$prefix = @($config.IPv4Address)[0].PrefixLength
$gateway = @($config.IPv4DefaultGateway)[0].NextHop
$dhcp = (Get-NetIPInterface -InterfaceIndex $index -AddressFamily IPv4).Dhcp

if ($Undo) {
    if ($dhcp -eq 'Enabled') {
        Write-Host "May dang nhan IP tu dong ($ip), khong can lam gi."
        exit 0
    }
    # The static default route outlives DHCP being switched back on, so drop it first.
    Remove-NetRoute -InterfaceIndex $index -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -Confirm:$false -ErrorAction SilentlyContinue
    Set-NetIPInterface -InterfaceIndex $index -AddressFamily IPv4 -Dhcp Enabled
    Set-DnsClientServerAddress -InterfaceIndex $index -ResetServerAddresses
    Write-Host "Da tra lai IP tu dong cho '$($config.InterfaceAlias)'."
    exit 0
}

if ($dhcp -eq 'Disabled') {
    Write-Host "IP da co dinh tu truoc: $ip"
} else {
    Set-NetIPInterface -InterfaceIndex $index -AddressFamily IPv4 -Dhcp Disabled
    Remove-NetIPAddress -InterfaceIndex $index -AddressFamily IPv4 -Confirm:$false -ErrorAction SilentlyContinue
    Remove-NetRoute -InterfaceIndex $index -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -Confirm:$false -ErrorAction SilentlyContinue
    New-NetIPAddress -InterfaceIndex $index -AddressFamily IPv4 -IPAddress $ip -PrefixLength $prefix -DefaultGateway $gateway | Out-Null
    Set-DnsClientServerAddress -InterfaceIndex $index -ServerAddresses $gateway
    Write-Host "Da co dinh IP $ip tren '$($config.InterfaceAlias)'."
}
Write-Host ''
Write-Host "Dien thoai vao bang: https://${ip}:3000"
