#requires -version 5.1
# NOTA: guardar este archivo en ASCII (sin tildes) para que PowerShell 5.1 no lo malinterprete.
param(
    [Parameter(Mandatory=$true)][ValidateSet('Inventario','Eventos','Cuentas','Antivirus','Firewall','Bateria','Temporales','Papelera','Disco','Intune','TrabajoEscuela','Permisos','RevocarAdmin','Apps','DistribucionAlmacenamiento')][string]$Tarea,
    [Parameter(Mandatory=$true)][string]$Estado
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Publicar([string]$situacion,[int]$porcentaje,[string]$detalle,[object]$datos=$null) {
    $porcentaje=[Math]::Max(0,[Math]::Min(100,$porcentaje))
    $obj=[ordered]@{estado=$situacion;porcentaje=$porcentaje;detalle=$detalle;actualizado=(Get-Date).ToString('o');datos=$datos}
    $json=$obj|ConvertTo-Json -Depth 9 -Compress
    $temp=$Estado+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
    [IO.File]::WriteAllText($temp,$json,([Text.UTF8Encoding]::new($false)))
    $publicado=$false
    try {
        for($intento=0;$intento -lt 12 -and -not $publicado;$intento++){
            try{
                if([IO.File]::Exists($Estado)){[IO.File]::Replace($temp,$Estado,[NullString]::Value)}
                else{[IO.File]::Move($temp,$Estado)}
                $publicado=$true
            }catch [IO.IOException]{
                # Si el supervisor esta leyendo el JSON, reintentar antes de declarar un fallo.
                Start-Sleep -Milliseconds (40+($intento*25))
            }
        }
        if(-not $publicado){
            # Ultimo recurso, apertura con lectura compartida. Los lectores reintentan JSON parcial.
            $stream=[IO.FileStream]::new($Estado,[IO.FileMode]::Create,[IO.FileAccess]::Write,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
            try{
                $bytes=[Text.Encoding]::UTF8.GetBytes($json)
                $stream.Write($bytes,0,$bytes.Length)
                $stream.Flush($true)
                $publicado=$true
            }finally{$stream.Dispose()}
        }
    }finally{
        if([IO.File]::Exists($temp)){try{[IO.File]::Delete($temp)}catch{}}
    }
}
function Paso([string]$mensaje,[int]$porcentaje=0) { Publicar 'En ejecucion' $porcentaje $mensaje }
function Terminar([string]$mensaje,[object]$datos=$null,[string]$estadoFinal='Completada') { Publicar $estadoFinal 100 $mensaje $datos }

# ---------- Ayudas compartidas (Papelera y Apps) ----------

# Tamano en MB de una carpeta (0 si no existe o no se puede leer)
function MedirMB([string]$ruta) {
    $s=(Get-ChildItem -LiteralPath $ruta -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    return [Math]::Round([int64]$s/1MB,1)
}

# Cierra un proceso y todos sus hijos (se usa cuando una tarea excede su tiempo)
function MatarArbol([int]$procId) {
    try{
        foreach($h in @(Get-CimInstance Win32_Process -Filter ('ParentProcessId='+$procId) -ErrorAction SilentlyContinue)){
            MatarArbol ([int]$h.ProcessId)
        }
    }catch{}
    try{ Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }catch{}
}

# Ejecuta un programa oculto, espera hasta $maxSeg segundos y devuelve si termino y su codigo de salida
function EjecutarEspera([string]$exe,[string]$argumentos,[int]$maxSeg) {
    $p=Start-Process -FilePath $exe -ArgumentList $argumentos -WindowStyle Hidden -PassThru
    $null=$p.Handle   # sin esto, ExitCode puede salir vacio
    if($p.WaitForExit($maxSeg*1000)){
        return [pscustomobject]@{terminado=$true;codigo=[int]$p.ExitCode}
    }
    MatarArbol $p.Id
    return [pscustomobject]@{terminado=$false;codigo=$null}
}

# Localiza winget.exe (puede no estar en el PATH si el script corre como otra cuenta)
function BuscarWinget {
    $c=Get-Command winget.exe -ErrorAction SilentlyContinue
    if($c){return $c.Source}
    try{
        $r=Get-ChildItem (Join-Path $env:ProgramFiles 'WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe') -ErrorAction SilentlyContinue |
           Sort-Object FullName -Descending | Select-Object -First 1
        if($r){return $r.FullName}
    }catch{}
    return $null
}

# Apps instaladas: programas (registro), paquetes Appx y rutas de aplicacion
$regUninst='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
           'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
           'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
function Get-Instaladas {
    $n=@()
    try{ $n+=@(Get-ItemProperty $regUninst -ErrorAction SilentlyContinue | Where-Object DisplayName | Select-Object -ExpandProperty DisplayName) }catch{}
    try{ $n+=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) }catch{}
    try{ $n+=@(Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName) }catch{}
    return @($n | Sort-Object -Unique)
}

try {
    Paso 'Iniciando tarea...' 0
    switch ($Tarea) {
        'Inventario' {
            $cs=Get-CimInstance Win32_ComputerSystem; Paso 'Equipo y fabricante consultados' 20
            $os=Get-CimInstance Win32_OperatingSystem; Paso 'Sistema operativo consultado' 40
            $bios=Get-CimInstance Win32_BIOS; $cpu=Get-CimInstance Win32_Processor | Select-Object -First 1; Paso 'BIOS y procesador consultados' 60
            $mem=@(Get-CimInstance Win32_PhysicalMemory); Paso 'Memoria fisica consultada' 80
            $unidades=@(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object { [pscustomobject]@{unidad=$_.DeviceID; libre_gb=[Math]::Round($_.FreeSpace/1GB,1); total_gb=[Math]::Round($_.Size/1GB,1)} })
            Terminar 'Inventario de hardware recopilado.' ([ordered]@{equipo=$env:COMPUTERNAME; usuario_sesion=$cs.UserName; fabricante=$cs.Manufacturer; modelo=$cs.Model; serie=$bios.SerialNumber; sistema_operativo=$os.Caption; version=$os.Version; cpu=$cpu.Name; ram_gb=[Math]::Round($cs.TotalPhysicalMemory/1GB,2); modulos_ram=$mem.Count; unidades=$unidades})
        }
        'Eventos' {
            $desde=(Get-Date).AddDays(-30)
            $datos=@();$avisos=@()
            $logs=@('System','Application')
            foreach($i in 0..($logs.Count-1)) {
                $log=$logs[$i]
                Paso ('Consultando eventos del registro '+$log) ([int](100*$i/$logs.Count))
                try {
                    # Limite para evitar capturar registros historicos excesivamente grandes.
                    $ev=@(Get-WinEvent -FilterHashtable @{LogName=$log;StartTime=$desde;Level=@(1,2,3)} -MaxEvents 15000 -ErrorAction Stop)
                    $top=@($ev|Group-Object ProviderName,Id|Sort-Object Count -Descending|Select-Object -First 6|ForEach-Object {
                        [pscustomobject]@{evento=$_.Name;cantidad=$_.Count}
                    })
                    $datos+= [pscustomobject]@{registro=$log; consultados=$ev.Count; criticos=@($ev|Where-Object Level -eq 1).Count; errores=@($ev|Where-Object Level -eq 2).Count; advertencias=@($ev|Where-Object Level -eq 3).Count; principales=$top; muestra_limitada=($ev.Count -eq 15000)}
                } catch {
                    $m=$_.Exception.Message
                    if($m -match '(?i)no events were found|no se encontraron eventos') {
                        $datos+= [pscustomobject]@{registro=$log;consultados=0;criticos=0;errores=0;advertencias=0}
                    } else { $avisos+=($log+': '+$m) }
                }
                Paso ('Registro '+$log+' terminado') ([int](100*($i+1)/$logs.Count))
            }
            if($datos.Count -eq 0) { throw ('No fue posible consultar los registros. '+($avisos -join '; ')) }
            $s=if($avisos.Count){'Con incidencias'}else{'Completada'}
            Terminar ('Analizados eventos de los ultimos 30 dias. '+($avisos -join '; ')) @{registros=$datos} $s
        }
        'Cuentas' {
            $archivo=Join-Path $env:ProgramData 'InventarioTI\credenciales.dpapi'
            if(-not (Test-Path -LiteralPath $archivo)) {
                Terminar 'Credenciales corporativas no aprovisionadas. No se cambiaron cuentas.' $null 'Omitida'
                break
            }
            $observaciones=@();$administradorActualizado=$false;$soporteListo=$false
            $secrets=$null;$raw=$null;$sAdmin=$null;$sSoporte=$null
            try {
                Add-Type -AssemblyName System.Security
                $raw=[Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($archivo),$null,[Security.Cryptography.DataProtectionScope]::LocalMachine)
                $secrets=([Text.Encoding]::UTF8.GetString($raw)|ConvertFrom-Json)
                [Array]::Clear($raw,0,$raw.Length);$raw=$null
                Paso 'Credenciales locales descifradas temporalmente' 15
                $laps=(Test-Path 'HKLM:\SOFTWARE\Microsoft\Policies\LAPS') -or (Test-Path 'HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd')
                $admin=Get-LocalUser|Where-Object {$_.SID.Value -match '-500$'}|Select-Object -First 1
                if(-not $admin){$observaciones+='No se identifico el administrador integrado (RID 500).'}
                elseif($laps){$observaciones+='Detectada configuracion LAPS: NO se modifico la contrasena del administrador.'}
                else {
                    try {$sAdmin=ConvertTo-SecureString ([string]$secrets.admin) -AsPlainText -Force; Set-LocalUser -Name $admin.Name -Password $sAdmin; $administradorActualizado=$true}
                    catch {$observaciones+=('Administrador: '+$_.Exception.Message)}
                }
                Paso 'Revisada cuenta administrativa' 55
                try {
                    $sSoporte=ConvertTo-SecureString ([string]$secrets.soporte) -AsPlainText -Force
                    $soporte=Get-LocalUser -Name 'Soporte' -ErrorAction SilentlyContinue
                    if($soporte){Set-LocalUser -Name 'Soporte' -Password $sSoporte}
                    else {New-LocalUser -Name 'Soporte' -Password $sSoporte -Description 'Usuario estandar de soporte TI' | Out-Null}
                    $grupo=Get-LocalGroup -SID 'S-1-5-32-544'
                    $soporte=Get-LocalUser -Name 'Soporte'
                    $miembros=@(Get-LocalGroupMember -Group $grupo.Name -ErrorAction Stop|Where-Object {$_.SID -and $_.SID.Value -eq $soporte.SID.Value})
                    if($miembros.Count){Remove-LocalGroupMember -Group $grupo.Name -Member 'Soporte' -ErrorAction Stop}
                    $soporteListo=$true
                } catch {$observaciones+=('Soporte: '+$_.Exception.Message)}
                $final=if($observaciones.Count){'Con incidencias'}else{'Completada'}
                Terminar ('Administrador actualizado: '+$administradorActualizado+'; Soporte estandar: '+$soporteListo+'. '+($observaciones -join ' ')) @{administrador_actualizado=$administradorActualizado;soporte_estandar=$soporteListo; advertencias=$observaciones} $final
            } finally {
                if($raw){[Array]::Clear($raw,0,$raw.Length)}
                $secrets=$null;$sAdmin=$null;$sSoporte=$null
            }
        }
        'Antivirus' {
            if(-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) {
                Terminar 'Microsoft Defender no disponible; puede existir antivirus corporativo.' $null 'Omitida'; break
            }
            $antes=Get-MpComputerStatus
            if(-not $antes.AMServiceEnabled -or -not $antes.AntivirusEnabled) {
                Terminar 'Microsoft Defender Antivirus no esta activo; no se ejecuto QuickScan.' @{real_time=$antes.RealTimeProtectionEnabled} 'Omitida'; break
            }
            $avisos=@()
            Paso 'Verificando antivirus; iniciando actualizacion de firmas' 0
            try { Update-MpSignature -ErrorAction Stop | Out-Null } catch { $avisos+=('Firmas: '+$_.Exception.Message) }
            Paso 'Ejecutando analisis rapido de Defender; el motor no informa porcentaje' 0
            try { Start-MpScan -ScanType QuickScan -ErrorAction Stop | Out-Null } catch { $avisos+=('Analisis: '+$_.Exception.Message) }
            $despues=Get-MpComputerStatus
            $estadoFinal=if($avisos.Count){'Con incidencias'}else{'Completada'}
            Terminar ('Defender analizado. '+($avisos -join '; ')) @{producto='Microsoft Defender';firma_fecha=$despues.AntivirusSignatureLastUpdated;ultimo_analisis_rapido=$despues.QuickScanEndTime;tiempo_real=$despues.RealTimeProtectionEnabled;avisos=$avisos} $estadoFinal
        }
        'Firewall' {
            $p=@(Get-NetFirewallProfile|Select-Object Name,Enabled,DefaultInboundAction,DefaultOutboundAction)
            $desactivados=@($p|Where-Object {$_.Enabled -eq $false})
            $resultado=if($desactivados.Count){'Con incidencias'}else{'Completada'}
            Terminar ('Perfiles consultados: '+$p.Count+'; desactivados: '+$desactivados.Count) @{perfiles=$p} $resultado
        }
        'Bateria' {
            $b=@(Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
            if(-not $b.Count) {Terminar 'No se detecto bateria (equipo de escritorio o datos no expuestos por WMI).' @{disponible=$false} 'Omitida';break}
            Paso 'Consultando capacidad de bateria' 50
            $datos=[ordered]@{disponible=$true;porcentaje=$b[0].EstimatedChargeRemaining;estado_codigo=$b[0].BatteryStatus;salud_porcentaje=$null}
            $design=Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue|Select-Object -First 1
            $full=Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue|Select-Object -First 1
            if($design -and $full -and $design.DesignedCapacity -gt 0){$datos.salud_porcentaje=[Math]::Round(100*$full.FullChargedCapacity/$design.DesignedCapacity,1)}
            Terminar ('Carga: '+$datos.porcentaje+'%; salud estimada: '+$(if($null -eq $datos.salud_porcentaje){'No disponible'}else{[string]$datos.salud_porcentaje+'%'})) $datos
        }
        'Temporales' {
            # Limpieza segura y aislada: NO borrar archivos usados por el supervisor.
            # Los archivos recientes (menos de 24 horas) se conservan para no
            # interferir con instalaciones, aplicaciones o tareas en ejecucion.
            $corte=(Get-Date).AddHours(-24)
            $raices=@($env:TEMP,(Join-Path $env:WINDIR 'Temp')) | Where-Object {$_ -and (Test-Path -LiteralPath $_ -PathType Container)} | Select-Object -Unique
            $archivos=New-Object 'System.Collections.Generic.List[object]'
            $directorios=New-Object 'System.Collections.Generic.List[object]'
            $pendientes=New-Object 'System.Collections.Generic.Stack[object]'
            $recientes=0; $enlaces=0; $erroresEnumeracion=0
            Paso 'Identificando temporales antiguos (se conservan los de menos de 24 horas)...' 0
            foreach($raiz in $raices){
                try{
                    foreach($hijo in @(Get-ChildItem -LiteralPath $raiz -Force -ErrorAction Stop)){
                        $pendientes.Push($hijo)
                    }
                }catch{$erroresEnumeracion++}
            }
            while($pendientes.Count -gt 0){
                $elemento=$pendientes.Pop()
                try {
                    if(($elemento.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){$enlaces++;continue}
                    if($elemento.PSIsContainer){
                        # Excluir permanentemente los expedientes de estado de cualquier
                        # ejecucion InventarioTI, incluyendo esta ejecucion paralela.
                        if($elemento.Name -like 'InventarioTI_*'){continue}
                        $directorios.Add($elemento)
                        try{
                            foreach($hijo in @(Get-ChildItem -LiteralPath $elemento.FullName -Force -ErrorAction Stop)){
                                $pendientes.Push($hijo)
                            }
                        }catch{$erroresEnumeracion++}
                    } else {
                        if($elemento.LastWriteTime -ge $corte){$recientes++}
                        else{$archivos.Add($elemento)}
                    }
                }catch{$erroresEnumeracion++}
            }
            $total=$archivos.Count
            $eliminados=0;$bloqueados=0;$procesados=0;$carpetasEliminadas=0
            if($total -gt 0){
                foreach($archivo in $archivos){
                    try{
                        # Proteccion extra: revalidar la antiguedad justo antes de borrar.
                        $actual=Get-Item -LiteralPath $archivo.FullName -Force -ErrorAction Stop
                        if($actual.LastWriteTime -ge $corte){$recientes++}
                        elseif(($actual.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){$enlaces++}
                        else{Remove-Item -LiteralPath $actual.FullName -Force -ErrorAction Stop;$eliminados++}
                    }catch{
                        # Otros archivos bloqueados no detienen esta tarea ni las demas.
                        $bloqueados++
                    }
                    $procesados++
                    if(($procesados % [Math]::Max(1,[int][Math]::Floor($total/100))) -eq 0 -or $procesados -eq $total){
                        Paso ('Temporales evaluados '+$procesados+'/'+$total) ([int][Math]::Floor(100*$procesados/$total))
                    }
                }
            }
            # Solo elimina carpetas que quedaron vacias; nunca usa -Recurse.
            foreach($directorio in @($directorios | Sort-Object { $_.FullName.Length } -Descending)){
                try{
                    $actual=Get-Item -LiteralPath $directorio.FullName -Force -ErrorAction Stop
                    if($actual.LastWriteTime -lt $corte){
                        Remove-Item -LiteralPath $actual.FullName -Force -ErrorAction Stop
                        $carpetasEliminadas++
                    }
                }catch{
                    # Sigue ocupada o contiene archivos: se conserva.
                }
            }
            $situacion=if($bloqueados -or $erroresEnumeracion){'Con incidencias'}else{'Completada'}
            $mensaje='Archivos antiguos eliminados: '+$eliminados+'/'+$total+'; bloqueados/no accesibles: '+$bloqueados+'; recientes conservados: '+$recientes+'; carpetas vacias: '+$carpetasEliminadas
            Terminar $mensaje @{archivos_evaluados=$total;archivos_eliminados=$eliminados;bloqueados=$bloqueados;recientes_conservados=$recientes;carpetas_vacias_eliminadas=$carpetasEliminadas;enlaces_omitidos=$enlaces;errores_enumeracion=$erroresEnumeracion;edad_minima_horas=24} $situacion
        }
        'Papelera' {
            # Se elimina la carpeta $Recycle.Bin de cada disco fijo (cubre a todos los usuarios);
            # Windows la vuelve a crear sola. Se usa "rd" en un proceso aparte para que la salida
            # de error de cmd no detenga la tarea con $ErrorActionPreference = 'Stop'.
            $unidades=@(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object -ExpandProperty DeviceID)
            if($unidades.Count -eq 0){ Terminar 'No se detectaron discos fijos.' $null 'Omitida'; break }
            $filas=@(); $totalLiberado=0; $parciales=0
            for($i=0;$i -lt $unidades.Count;$i++){
                $d=$unidades[$i]
                $rb=$d+'\$Recycle.Bin'
                Paso ('Vaciando papelera de '+$d+' ('+($i+1)+'/'+$unidades.Count+')') ([int][Math]::Floor(100*$i/$unidades.Count))
                if(-not (Test-Path -LiteralPath $rb)){
                    $filas+=[pscustomobject]@{unidad=$d;antes_mb=0;despues_mb=0;liberado_mb=0;resultado='Sin papelera'}
                    continue
                }
                $antes=MedirMB $rb
                $r=EjecutarEspera 'cmd.exe' ('/c rd /s /q '+$rb) 300
                $despues=MedirMB $rb
                $liberado=[Math]::Round($antes-$despues,1)
                $totalLiberado=[Math]::Round($totalLiberado+$liberado,1)
                if(-not $r.terminado){ $res='Tiempo agotado'; $parciales++ }
                elseif($despues -gt 0){ $res='Parcial (archivos en uso)'; $parciales++ }
                else{ $res='Vaciada' }
                $filas+=[pscustomobject]@{unidad=$d;antes_mb=$antes;despues_mb=$despues;liberado_mb=$liberado;resultado=$res}
            }
            $lineas=@()
            foreach($f in $filas){
                $lineas+=('  '+$f.unidad+' '+$f.resultado+': '+$f.antes_mb+' MB -> '+$f.despues_mb+' MB (liberados '+$f.liberado_mb+' MB)')
            }
            $resumen='Papelera: liberados '+$totalLiberado+' MB en '+$unidades.Count+' unidad(es)'
            if($parciales){ $resumen+='; incompleta en '+$parciales }
            $detalle=$resumen+"`n"+'-- POR UNIDAD --'+"`n"+($lineas -join "`n")
            $final=if($parciales){'Con incidencias'}else{'Completada'}
            Terminar $detalle @{resumen=[ordered]@{unidades=$unidades.Count;liberado_mb=$totalLiberado;con_incidencias=$parciales};unidades=$filas} $final
        }
        'Apps' {
            # Apps requeridas: nombre = patron que se busca en programas / paquetes Appx / rutas de aplicacion
            $requeridas=[ordered]@{
                'Outlook'          = 'Outlook'
                'Teams'            = 'Teams'
                'Word'             = 'WINWORD|Microsoft 365|Office'
                'Excel'            = 'EXCEL|Microsoft 365|Office'
                'OneDrive'         = 'OneDrive'
                'Adobe Acrobat'    = 'Acrobat|AcroRd'
                'ServiceDesk Plus' = 'ServiceDesk'
                'AnyDesk'          = 'AnyDesk'
                'Windows App'      = 'Windows365|Windows App'
                'Azure VPN Client' = 'AzureVpn|Azure VPN'
            }
            # Apps que se intentan actualizar con winget (solo las faciles)
            $paraActualizar=@(
                [pscustomobject]@{app='Teams';id='Microsoft.Teams'},
                [pscustomobject]@{app='OneDrive';id='Microsoft.OneDrive'},
                [pscustomobject]@{app='AnyDesk';id='AnyDesk.AnyDesk'},
                [pscustomobject]@{app='Adobe Acrobat Reader';id='Adobe.Acrobat.Reader.64-bit'}
            )

            # --- Segmento 1: apps requeridas ---
            Paso 'Verificando aplicaciones requeridas...' 5
            $inst=Get-Instaladas
            $filasReq=@()
            foreach($k in $requeridas.Keys){
                $patron=$requeridas[$k]
                $ok=[bool]($inst | Where-Object { $_ -match $patron } | Select-Object -First 1)
                $filasReq+=[pscustomobject]@{app=$k;estado=$(if($ok){'Instalada'}else{'Falta'})}
            }
            $faltantes=@($filasReq | Where-Object { $_.estado -eq 'Falta' })
            $nombresFaltantes=@($faltantes | ForEach-Object { $_.app })

            # --- Segmento 2: actualizaciones ---
            Paso 'Preparando actualizaciones...' 30
            $acts=@()
            $wg=BuscarWinget
            if(-not $wg){
                $acts+=[pscustomobject]@{app='winget';resultado='No disponible; se omiten las actualizaciones por winget';nivel='aviso'}
            } else {
                for($i=0;$i -lt $paraActualizar.Count;$i++){
                    $it=$paraActualizar[$i]
                    Paso ('Actualizando '+$it.app+' ('+($i+1)+'/'+$paraActualizar.Count+')') ([int](30+[Math]::Floor(50*$i/$paraActualizar.Count)))
                    $r=EjecutarEspera $wg ('upgrade --id '+$it.id+' -e --silent --accept-source-agreements --accept-package-agreements') 480
                    if(-not $r.terminado){ $txt='Tiempo agotado (8 min)'; $niv='error' }
                    elseif($r.codigo -eq 0){ $txt='Actualizada'; $niv='ok' }
                    elseif($r.codigo -eq -1978335189){ $txt='Ya esta al dia'; $niv='ok' }
                    elseif($r.codigo -eq -1978335212){ $txt='No instalada (winget no la encuentra)'; $niv='aviso' }
                    else{ $txt='Error (codigo '+$r.codigo+')'; $niv='error' }
                    $acts+=[pscustomobject]@{app=$it.app;resultado=$txt;nivel=$niv}
                }
            }
            Paso 'Actualizando Office...' 85
            $c2r=Join-Path $env:CommonProgramFiles 'Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
            if(Test-Path -LiteralPath $c2r){
                $r=EjecutarEspera $c2r '/update user displaylevel=false forceappshutdown=false' 120
                if(-not $r.terminado){ $acts+=[pscustomobject]@{app='Office';resultado='Tiempo agotado (2 min)';nivel='error'} }
                elseif($r.codigo -eq 0){ $acts+=[pscustomobject]@{app='Office';resultado='Actualizacion solicitada (se aplica en segundo plano)';nivel='ok'} }
                else{ $acts+=[pscustomobject]@{app='Office';resultado=('Error (codigo '+$r.codigo+')');nivel='error'} }
            } else {
                $acts+=[pscustomobject]@{app='Office';resultado='No aplica (no es Click-to-Run)';nivel='aviso'}
            }

            # --- Armado del resultado segmentado ---
            $nOk=$filasReq.Count-$faltantes.Count
            $nAplicadas=@($acts | Where-Object { $_.resultado -eq 'Actualizada' -or $_.resultado -like 'Actualizacion solicitada*' }).Count
            $nAlDia=@($acts | Where-Object { $_.resultado -eq 'Ya esta al dia' }).Count
            $nError=@($acts | Where-Object { $_.nivel -eq 'error' }).Count
            $lineasReq=@()
            foreach($f in $filasReq){
                $marca=if($f.estado -eq 'Instalada'){'[OK]    '}else{'[FALTA] '}
                $lineasReq+=('  '+$marca+$f.app)
            }
            $lineasAct=@()
            foreach($a in $acts){
                $marca=switch($a.nivel){ 'ok'{'[OK] '} 'aviso'{'[..] '} default{'[!!] '} }
                $lineasAct+=('  '+$marca+$a.app+': '+$a.resultado)
            }
            $resumen='Apps: '+$nOk+'/'+$filasReq.Count+' requeridas instaladas; actualizaciones: '+$nAplicadas+' aplicadas, '+$nAlDia+' al dia, '+$nError+' con error'
            $detalle=$resumen+"`n"+'-- APPS REQUERIDAS --'+"`n"+($lineasReq -join "`n")+"`n"+'-- ACTUALIZACIONES --'+"`n"+($lineasAct -join "`n")
            $final=if($faltantes.Count -or $nError){'Con incidencias'}else{'Completada'}
            $datos=[ordered]@{
                resumen=[ordered]@{requeridas_total=$filasReq.Count;instaladas=$nOk;faltantes=$nombresFaltantes;actualizaciones_aplicadas=$nAplicadas;al_dia=$nAlDia;con_error=$nError}
                requeridas=$filasReq
                actualizaciones=$acts
            }
            Terminar $detalle $datos $final
        }
        'Disco' {
            $d=@(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3'|ForEach-Object {[pscustomobject]@{unidad=$_.DeviceID;libre_gb=[Math]::Round($_.FreeSpace/1GB,2);total_gb=[Math]::Round($_.Size/1GB,2);ocupacion_porcentaje=if($_.Size -gt 0){[Math]::Round(100*($_.Size-$_.FreeSpace)/$_.Size,1)}else{$null}}})
            Terminar ('Espacio libre consultado en '+$d.Count+' unidades.') @{unidades=$d}
        }
        'DistribucionAlmacenamiento' {
            # Inventario de tamanos LOGICOS del perfil interactivo. No elimina archivos.
            # Se evitan enlaces/junctions (reparse points) y se limita el tiempo.
            Paso 'Localizando el perfil del usuario interactivo' 5
            $usuario=(Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
            $perfil=$null
            $sidUsuario=$null
            if(-not [string]::IsNullOrWhiteSpace($usuario)) {
                try {
                    $sidUsuario=([Security.Principal.NTAccount]::new($usuario)).Translate([Security.Principal.SecurityIdentifier]).Value
                    $perfil=@(Get-CimInstance Win32_UserProfile -Filter ("SID='"+$sidUsuario+"'") -ErrorAction Stop | Select-Object -First 1)[0].LocalPath
                }catch{}
            }
            if(-not $perfil -and $usuario) {
                # En Entra ID algunos usuarios no se resuelven mediante NTAccount.Translate.
                # Se consulta el SID propietario de explorer.exe en la sesion interactiva.
                try {
                    $explorers=@(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop)
                    foreach($explorer in $explorers) {
                        $propietario=Invoke-CimMethod -InputObject $explorer -MethodName GetOwner -ErrorAction SilentlyContinue
                        if($propietario -and (($propietario.Domain+'\'+$propietario.User) -ieq $usuario)) {
                            $propietarioSid=Invoke-CimMethod -InputObject $explorer -MethodName GetOwnerSid -ErrorAction SilentlyContinue
                            if($propietarioSid -and $propietarioSid.Sid) {
                                $sidUsuario=[string]$propietarioSid.Sid
                                $perfil=@(Get-CimInstance Win32_UserProfile -Filter ("SID='"+$sidUsuario+"'") -ErrorAction Stop | Select-Object -First 1)[0].LocalPath
                                if($perfil){break}
                            }
                        }
                    }
                }catch{}
            }
            if(-not $perfil -and $usuario -and $usuario -ieq [Security.Principal.WindowsIdentity]::GetCurrent().Name) {
                $perfil=$env:USERPROFILE
            }
            if(-not $perfil -or -not (Test-Path -LiteralPath $perfil -PathType Container)) {
                Terminar 'No se pudo identificar con seguridad el perfil interactivo; no se midio otra cuenta.' @{usuario=$usuario;motivo='Perfil interactivo no identificado'} 'Omitida'
                break
            }
            $perfil=[IO.Path]::GetFullPath($perfil).TrimEnd('\')
            Paso ('Analizando las carpetas de '+$perfil) 12
            $inicioAnalisis=Get-Date
            $limiteSegundos=420
            $limiteArchivos=1200000
            $totalArchivos=0L
            $omitidos=0L
            $caducado=$false
            $nombres=@{ 'Downloads'='Descargas';'Documents'='Documentos';'Desktop'='Escritorio';'Pictures'='Imagenes';'Videos'='Videos';'Music'='Musica';'AppData'='Datos de aplicaciones';'OneDrive'='OneDrive';'Favorites'='Favoritos';'Contacts'='Contactos';'Saved Games'='Juegos guardados';'Searches'='Busquedas';'Links'='Accesos directos' }
            $carpetas=@(Get-ChildItem -LiteralPath $perfil -Force -Directory -ErrorAction SilentlyContinue | Where-Object { -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) } | Sort-Object Name)
            $filas=New-Object 'System.Collections.Generic.List[object]'
            $bytesPerfil=0L
            $i=0
            foreach($carpeta in $carpetas) {
                if($caducado){break}
                $i++
                $nombre=if($nombres.ContainsKey($carpeta.Name)){$nombres[$carpeta.Name]}else{$carpeta.Name}
                Paso ('Midiendo '+$nombre+' ('+$i+'/'+$carpetas.Count+')') ([Math]::Min(90,12+[int](78*$i/[Math]::Max(1,$carpetas.Count))))
                $bytes=0L
                $archivosCarpeta=0L
                $pendientes=New-Object 'System.Collections.Generic.Stack[string]'
                $pendientes.Push($carpeta.FullName)
                while($pendientes.Count -gt 0) {
                    if(((Get-Date)-$inicioAnalisis).TotalSeconds -ge $limiteSegundos -or $totalArchivos -ge $limiteArchivos){$caducado=$true;break}
                    $actual=$pendientes.Pop()
                    try {
                        foreach($f in [IO.DirectoryInfo]::new($actual).EnumerateFiles()) {
                            if($f.Attributes -band [IO.FileAttributes]::ReparsePoint){continue}
                            try{$bytes += [long]$f.Length;$totalArchivos++;$archivosCarpeta++}catch{$omitidos++}
                            if(($totalArchivos % 1000) -eq 0 -and (((Get-Date)-$inicioAnalisis).TotalSeconds -ge $limiteSegundos -or $totalArchivos -ge $limiteArchivos)) {$caducado=$true;break}
                        }
                    }catch{$omitidos++}
                    if($caducado){break}
                    try {
                        foreach($sub in [IO.DirectoryInfo]::new($actual).EnumerateDirectories()) {
                            if(-not ($sub.Attributes -band [IO.FileAttributes]::ReparsePoint)){$pendientes.Push($sub.FullName)}
                        }
                    }catch{$omitidos++}
                }
                $bytesPerfil+=$bytes
                [void]$filas.Add([pscustomobject]@{carpeta=$nombre;directorio=$carpeta.Name;tamano_gb=[Math]::Round($bytes/1GB,2);tamano_bytes=$bytes;archivos=$archivosCarpeta})
            }
            # Archivos que se encuentran directamente en la raiz del perfil.
            $raizBytes=0L
            if(-not $caducado) {
                try {
                    foreach($f in [IO.DirectoryInfo]::new($perfil).EnumerateFiles()) {
                        if(-not ($f.Attributes -band [IO.FileAttributes]::ReparsePoint)){$raizBytes += [long]$f.Length}
                    }
                }catch{$omitidos++}
            }
            if($raizBytes -gt 0){[void]$filas.Add([pscustomobject]@{carpeta='Archivos raiz del perfil';directorio='.';tamano_gb=[Math]::Round($raizBytes/1GB,2);tamano_bytes=$raizBytes;archivos=$null})}
            $bytesPerfil+=$raizBytes
            Paso 'Consultando espacio libre en la unidad' 94
            $unidad=[IO.Path]::GetPathRoot($perfil)
            $disco=Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='"+$unidad.TrimEnd('\\')+"'") -ErrorAction SilentlyContinue
            $libreGb=if($disco){[Math]::Round($disco.FreeSpace/1GB,2)}else{$null}
            $totalGb=if($disco){[Math]::Round($disco.Size/1GB,2)}else{$null}
            $ordenadas=@($filas | Sort-Object tamano_bytes -Descending)
            $top=@($ordenadas | Select-Object -First 8 | ForEach-Object {$_.carpeta+': '+$_.tamano_gb+' GB'})
            $detalle='Perfil: '+[Math]::Round($bytesPerfil/1GB,2)+' GB. '+($top -join ' | ')
            if($detalle.Length -gt 1600){$detalle=$detalle.Substring(0,1597)+'...'}
            $estadoFinal=if($caducado -or $omitidos -gt 0){'Con incidencias'}else{'Completada'}
            if($caducado){$detalle+=' | Medicion parcial: tiempo o limite de archivos alcanzado.'}
            if($omitidos -gt 0){$detalle+=' | Carpetas/archivos no accesibles: '+$omitidos+'.'}
            $datos=[ordered]@{
                usuario=$usuario;perfil=$perfil;unidad=$unidad;
                perfil_total_gb=[Math]::Round($bytesPerfil/1GB,2);
                unidad_total_gb=$totalGb;unidad_libre_gb=$libreGb;
                carpetas=$ordenadas;archivos_analizados=$totalArchivos;
                incidencias_acceso=$omitidos;medicion_parcial=$caducado;
                alcance='Tamano logico de archivos del perfil interactivo. No es una distribucion completa de la unidad; no incluye Windows, programas ni otros perfiles.'
            }
            Terminar $detalle $datos $estadoFinal
        }
        'Intune' {
            $svc=Get-Service -Name IntuneManagementExtension -ErrorAction SilentlyContinue
            $status=if($svc){$svc.Status.ToString()}else{'No detectado'}
            Terminar ('Intune Management Extension: '+$status) @{servicio=$status}
        }
        'TrabajoEscuela' {
            if(-not (Get-Command dsregcmd.exe -ErrorAction SilentlyContinue)){throw 'dsregcmd.exe no esta disponible.'}
            $salida=(& dsregcmd.exe /status 2>&1 | Out-String)
            $azure=if($salida -match 'AzureAdJoined\s*:\s*(YES|NO)'){$Matches[1]}else{'No disponible'}
            $domain=if($salida -match 'DomainJoined\s*:\s*(YES|NO)'){$Matches[1]}else{'No disponible'}
            Terminar ('Microsoft Entra unido: '+$azure+'; dominio unido: '+$domain) @{azure_ad_joined=$azure;domain_joined=$domain}
        }
        'RevocarAdmin' {
            # Solo revoca la membresia DIRECTA del usuario corporativo interactivo.
            # No modifica cuentas locales, cuentas integradas, grupos ni roles Entra/AD.
            Paso 'Identificando al usuario que inicio sesion' 10
            $usuario=(Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
            if([string]::IsNullOrWhiteSpace($usuario)) {
                Terminar 'No se identifico un usuario interactivo; no se cambiaron permisos.' @{revocado=$false;motivo='Sin usuario interactivo'} 'Omitida'
                break
            }
            $partes=$usuario -split '\\',2
            if($partes.Count -ne 2 -or $partes[0] -ieq $env:COMPUTERNAME -or $partes[0] -eq '.') {
                Terminar ('Cuenta local o no identificable ('+$usuario+'); no se revocan permisos.') @{usuario=$usuario;revocado=$false;motivo='Cuenta local o no identificable'} 'Omitida'
                break
            }
            Paso ('Consultando la asignacion administrativa de '+$usuario) 30
            $grupo=Get-LocalGroup -SID 'S-1-5-32-544' -ErrorAction Stop
            $miembros=@(Get-LocalGroupMember -Group $grupo -ErrorAction Stop)
            $miembro=@($miembros|Where-Object {$_.Name -ieq $usuario}|Select-Object -First 1)
            if($miembro.Count -eq 0) {
                Terminar ('El usuario '+$usuario+' no tiene asignacion directa como administrador local.') @{usuario=$usuario;revocado=$false;motivo='Sin asignacion directa'} 'Omitida'
                break
            }
            $objetivo=$miembro[0]
            $fuente=[string]$objetivo.PrincipalSource
            if($fuente -notin @('AzureAD','ActiveDirectory')) {
                Terminar ('Cuenta no corporativa comprobada ('+$fuente+'); sin cambios.') @{usuario=$usuario;revocado=$false;origen=$fuente;motivo='Origen no permitido'} 'Omitida'
                break
            }
            if(-not $objetivo.SID) { throw 'No se pudo verificar el SID de la cuenta objetivo; no se cambiaron permisos.' }
            # Proteccion adicional de las cuentas integradas de Windows.
            if($objetivo.SID.Value -match '-500$' -or $partes[1] -ieq 'Administrador' -or $partes[1] -ieq 'Administrator') {
                Terminar 'Cuenta administrativa protegida; no se cambiaron permisos.' @{usuario=$usuario;revocado=$false;motivo='Cuenta protegida'} 'Omitida'
                break
            }
            # Exigir al menos otra cuenta LOCAL habilitada y miembro directo del grupo.
            # Esta prueba no garantiza que se conozca la contrasena: TI debe validarla previamente.
            Paso 'Verificando una cuenta administrativa local de respaldo' 55
            $respaldo=@()
            foreach($cuentaLocal in @(Get-LocalUser -ErrorAction Stop | Where-Object {$_.Enabled})) {
                if(@($miembros|Where-Object {$_.SID -and $_.SID.Value -eq $cuentaLocal.SID.Value}).Count -gt 0) {
                    $respaldo+= $cuentaLocal.Name
                }
            }
            if($respaldo.Count -eq 0) {
                Terminar 'Sin otro administrador local habilitado: revocacion omitida para evitar bloqueo administrativo.' @{usuario=$usuario;revocado=$false;motivo='Sin respaldo local habilitado'} 'Omitida'
                break
            }
            # Evitar que una ejecucion sin privilegios intente un cambio parcial.
            $principal=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
            if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                throw 'Esta tarea requiere ejecutarse elevada como administrador.'
            }
            Paso ('Revocando asignacion directa de '+$usuario) 80
            Remove-LocalGroupMember -Group $grupo -Member $objetivo -ErrorAction Stop
            $restantes=@(Get-LocalGroupMember -Group $grupo -ErrorAction Stop)
            $aunExiste=@($restantes|Where-Object {$_.SID -and $_.SID.Value -eq $objetivo.SID.Value}).Count -gt 0
            if($aunExiste){throw 'Se solicito la revocacion, pero la cuenta continua en Administradores locales.'}
            Terminar ('Asignacion directa revocada para '+$usuario+'. Cerrar sesion para renovar privilegios.') ([ordered]@{usuario=$usuario;origen=$fuente;revocado=$true;sid=$objetivo.SID.Value;respaldo_local_detectado=$respaldo;advertencia='No verifica la contrasena del respaldo ni revoca membresias indirectas por grupos de Entra/AD.'})
        }
        'Permisos' {
            $principal=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
            $isAdmin=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            Terminar ('Permisos administrativos: '+$isAdmin) @{elevado=$isAdmin} $(if($isAdmin){'Completada'}else{'Con incidencias'})
        }
    }
} catch {
    $mensajeError=$_.Exception.Message
    try { Publicar 'Error' 0 $mensajeError }
    catch {
        # Nunca ocultar el error original si tampoco se puede publicar el estado.
        try {
            $j=[ordered]@{estado='Error';porcentaje=0;detalle=$mensajeError;actualizado=(Get-Date).ToString('o');datos=$null} | ConvertTo-Json -Compress
            [IO.File]::WriteAllText($Estado,$j,([Text.UTF8Encoding]::new($false)))
        } catch {
            [Console]::Error.WriteLine('Fallo al escribir estado: '+$_.Exception.Message)
        }
    }
    [Console]::Error.WriteLine('Tarea '+$Tarea+' fallida: '+$mensajeError)
    exit 1
}