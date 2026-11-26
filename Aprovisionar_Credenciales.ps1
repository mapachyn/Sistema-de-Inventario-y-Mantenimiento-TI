#requires -version 5.1
$ErrorActionPreference='Stop'
$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Ejecute PowerShell como administrador.' }
Write-Host 'Aprovisionamiento unico por un administrador autorizado de TI.'
Write-Host 'No ejecutar en equipos administrados por LAPS sin aprobacion de TI.' -ForegroundColor Yellow
$a=Read-Host 'Contrasena del administrador local (oculta)' -AsSecureString
$b=Read-Host 'Contrasena de Soporte (oculta)' -AsSecureString
function Convertir([Security.SecureString]$s) {
    $ptr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}
$dir=Join-Path $env:ProgramData 'InventarioTI'
New-Item -Path $dir -ItemType Directory -Force | Out-Null
# ACL restringida: SYSTEM y administradores locales. DPAPI ligado a este equipo.
& icacls.exe $dir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'No se pudo proteger la carpeta de credenciales.' }
try {
    $json=@{admin=(Convertir $a);soporte=(Convertir $b)} | ConvertTo-Json -Compress
    Add-Type -AssemblyName System.Security
    $raw=[Text.Encoding]::UTF8.GetBytes($json)
    $encrypted=[Security.Cryptography.ProtectedData]::Protect($raw,$null,[Security.Cryptography.DataProtectionScope]::LocalMachine)
    [IO.File]::WriteAllBytes((Join-Path $dir 'credenciales.dpapi'),$encrypted)
    [Array]::Clear($raw,0,$raw.Length)
    Write-Host 'Credenciales cifradas localmente. No se guardaron en BAT ni PS1.' -ForegroundColor Green
} finally { $json=$null; $a=$null; $b=$null }
