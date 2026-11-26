#requires -version 5.1
<#
 Inventario y mantenimiento TI - controlador de procesos independientes.
 Modo automatico: limite global de 35 minutos para todas las tareas.
 Modo manual: una sola tarea individual por ejecucion.
 Los archivos de estado temporales son privados del proceso; el unico reporte
 persistente de cada ejecucion es el documento Word generado junto al BAT.
#>
[CmdletBinding()]
param([ValidateSet('Menu','Automatico','Manual')][string]$Modo='Menu',
      [string]$Tarea='')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$tareasScript=Join-Path $root 'Tareas.ps1'
$plantilla=Join-Path $root 'Constancia_Mantenimiento_Preventivo.docx'
$maxSegundos=35*60
$catalogo=[ordered]@{
    Inventario='Inventario de hardware'
    Eventos='Eventos de Windows (30 dias)'
    Cuentas='Cuentas locales (Administrador y Soporte)'
    Antivirus='Antivirus: Defender y analisis rapido'
    Firewall='Perfiles del firewall'
    Apps='Aplicaciones: verificar y actualizar'
    Bateria='Bateria: carga y salud'
    Temporales='Eliminacion de archivos temporales'
    Papelera='Vaciado de papelera (despues de temporales)'
    Disco='Espacio y uso del disco'
    DistribucionAlmacenamiento='Distribucion del almacenamiento por carpetas'
    Intune='Servicio de Intune'
    TrabajoEscuela='Vinculacion corporativa'
    Permisos='Permisos de ejecucion'
    RevocarAdmin='Verificacion y Correccion de permisos en el usuario'
}
if(-not (Test-Path -LiteralPath $tareasScript)){throw "Falta Tareas.ps1 junto al BAT: $tareasScript"}
function Nuevo-Registro([string]$id,[string]$temp){
    [pscustomobject]@{
        Id=$id;Nombre=[string]$catalogo[$id];Estado='Pendiente';Porcentaje=0;
        Detalle='';Datos=$null;Inicio=$null;Fin=$null;Proceso=$null;
        Archivo=(Join-Path $temp ($id+'.json'));StdOut=(Join-Path $temp ($id+'.out'));
        StdErr=(Join-Path $temp ($id+'.err'))
    }
}
function Iniciar-Tarea($reg){
    try {
        $reg.Inicio=Get-Date
        $reg.Estado='En ejecucion'
        $reg.Porcentaje=0
        $argumentos='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$tareasScript+'" -Tarea "'+$reg.Id+'" -Estado "'+$reg.Archivo+'"'
        $reg.Proceso=Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $argumentos -PassThru -WindowStyle Hidden -RedirectStandardOutput $reg.StdOut -RedirectStandardError $reg.StdErr -ErrorAction Stop
    }catch{
        $reg.Estado='Error';$reg.Fin=Get-Date;$reg.Detalle='No se pudo iniciar: '+$_.Exception.Message
    }
}
function Leer-Estado($reg){
    if(-not (Test-Path -LiteralPath $reg.Archivo)){return $false}
    try {
        $lectura=[IO.FileStream]::new($reg.Archivo,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        try{$reader=[IO.StreamReader]::new($lectura,[Text.Encoding]::UTF8);$contenido=$reader.ReadToEnd()}
        finally{if($reader){$reader.Dispose()}else{$lectura.Dispose()}}
        $json=$contenido|ConvertFrom-Json -ErrorAction Stop
        if($json.PSObject.Properties['porcentaje']){$reg.Porcentaje=[math]::Max(0,[math]::Min(100,[int]$json.porcentaje))}
        if($json.PSObject.Properties['detalle']){$reg.Detalle=[string]$json.detalle}
        if($json.PSObject.Properties['datos']){$reg.Datos=$json.datos}
        if($json.PSObject.Properties['estado'] -and ([string]$json.estado) -in @('Completada','Con incidencias','Omitida','Error')){
            $reg.Estado=[string]$json.estado
            $reg.Fin=Get-Date
            return $true
        }
    }catch{
        # Un archivo puede estar en transicion; nunca detener el resto por esto.
    }
    return $false
}
function Actualizar-Tarea($reg){
    if($reg.Estado -ne 'En ejecucion'){return}
    $termino=Leer-Estado $reg
    if($termino){return}
    if($null -eq $reg.Proceso){$reg.Estado='Error';$reg.Detalle='No se creo proceso.';$reg.Fin=Get-Date;return}
    try{$reg.Proceso.Refresh()}catch{}
    if($reg.Proceso.HasExited){
        # Pequeño margen para lectura del ultimo estado escrito por el hijo.
        Start-Sleep -Milliseconds 125
        if(Leer-Estado $reg){return}
        $reg.Estado='Error';$reg.Fin=Get-Date
        $errorTexto=''
        try{if(Test-Path -LiteralPath $reg.StdErr){$errorTexto=[IO.File]::ReadAllText($reg.StdErr)}}catch{}
        if($errorTexto.Length -gt 700){$errorTexto=$errorTexto.Substring(0,700)}
        $reg.Detalle=if($errorTexto.Trim()){$errorTexto.Trim()}else{'El proceso termino sin escribir un estado valido. Codigo: '+$reg.Proceso.ExitCode}
    }
}
function Detener-Tarea($reg){
    if($null -eq $reg.Proceso){return}
    try {
        $reg.Proceso.Refresh()
        if(-not $reg.Proceso.HasExited){
            # /T intenta finalizar los subprocesos que el trabajador pudo iniciar.
            & "$env:SystemRoot\System32\taskkill.exe" /PID $reg.Proceso.Id /T /F 2>$null | Out-Null
            try{$reg.Proceso.Refresh()}catch{}
            if(-not $reg.Proceso.HasExited){Stop-Process -Id $reg.Proceso.Id -Force -ErrorAction SilentlyContinue}
        }
    }catch{}
}
function Pintar-Pantalla($lista,[datetime]$inicio,[string]$tipo){
    Clear-Host
    Write-Host 'INVENTARIO Y MANTENIMIENTO TI - PANEL DE ACTIVIDADES' -ForegroundColor Cyan
    $elapsed=[int]((Get-Date)-$inicio).TotalSeconds
    $tiempo=[timespan]::FromSeconds([math]::Max(0,$elapsed))
    if($tipo -eq 'Automatico'){
        $restante=[timespan]::FromSeconds([math]::Max(0,$maxSegundos-$elapsed))
        Write-Host ('Modo: Automatico | Transcurrido: {0:hh\:mm\:ss} | Restante: {1:hh\:mm\:ss}' -f $tiempo,$restante)
    }else{Write-Host ('Modo: Manual | Transcurrido: {0:hh\:mm\:ss}' -f $tiempo)}
    Write-Host ('-'*98)
    foreach($r in $lista){
        $avance=if($r.Estado -eq 'Pendiente'){'Pendiente'}elseif($r.Estado -eq 'En ejecucion' -and $r.Porcentaje -eq 0){'En ejecucion'}else{('{0}%' -f $r.Porcentaje)}
        $color=switch($r.Estado){'Completada'{'Green'} 'Error'{'Red'} 'Tiempo agotado'{'Yellow'} 'Con incidencias'{'Yellow'} 'Omitida'{'DarkYellow'} default{'Gray'}}
        Write-Host ('{0,-45} {1,-15} {2}' -f $r.Nombre,$avance,$r.Estado) -ForegroundColor $color
        if($r.Estado -in @('En ejecucion','Error','Tiempo agotado','Con incidencias','Omitida') -and $r.Detalle){
            $d=$r.Detalle.Replace("`r",' ').Replace("`n",' ')
            if($d.Length -gt 105){$d=$d.Substring(0,102)+'...'}
            Write-Host ('   '+$d) -ForegroundColor DarkGray
        }
    }
    Write-Host ('-'*98)
    Write-Host 'El avance es real cuando la tarea permite medirlo; de lo contrario indica En ejecucion.'
}
function Solicitar-Texto([string]$titulo,[string]$predeterminado='', [bool]$obligatorio=$false){
    do {
        $etiqueta=if($predeterminado){$titulo+' ['+$predeterminado+']'}else{$titulo}
        $res=(Read-Host $etiqueta)
        if($null -eq $res){$res=''}
        $res=$res.Trim()
        if(-not $res){$res=$predeterminado}
        if($obligatorio -and -not $res){Write-Host 'Dato obligatorio, intente de nuevo.' -ForegroundColor Yellow}
    }while($obligatorio -and -not $res)
    return [string]$res
}
function Solicitar-Opcion([string]$titulo,[string[]]$opciones){
    Write-Host "`n$titulo" -ForegroundColor Cyan
    for($i=0;$i -lt $opciones.Count;$i++){Write-Host ('  {0}. {1}' -f ($i+1),$opciones[$i])}
    while($true){
        $texto=Read-Host 'Elija un numero'
        $numero=0
        if([int]::TryParse($texto,[ref]$numero) -and $numero -ge 1 -and $numero -le $opciones.Count){return $opciones[$numero-1]}
        Write-Host 'Opcion no valida.' -ForegroundColor Yellow
    }
}
function Obtener-CorreosUsuario([string]$usuarioSesion){
    # Busca los correos Microsoft/Office/OneDrive del usuario con sesion y deja solo los corporativos.
    $raiz='Registry::HKEY_CURRENT_USER';$esActual=$true
    try{
        if($usuarioSesion){
            $sid=(New-Object Security.Principal.NTAccount($usuarioSesion)).Translate([Security.Principal.SecurityIdentifier]).Value
            if($sid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -and (Test-Path -LiteralPath ('Registry::HKEY_USERS\'+$sid))){
                $raiz='Registry::HKEY_USERS\'+$sid;$esActual=$false
            }
        }
    }catch{}
    $crudos=@()
    try{$crudos+=@(Get-ItemProperty -Path ($raiz+'\SOFTWARE\Microsoft\OneDrive\Accounts\*') -ErrorAction Stop|ForEach-Object{[string]$_.UserEmail})}catch{}
    try{$crudos+=@(Get-ChildItem -Path ($raiz+'\SOFTWARE\Microsoft\IdentityCRL\StoredIdentities') -ErrorAction Stop|ForEach-Object{[string]$_.PSChildName})}catch{}
    try{$crudos+=@(Get-ItemProperty -Path ($raiz+'\SOFTWARE\Microsoft\Office\16.0\Common\Identity\Identities\*') -ErrorAction Stop|ForEach-Object{[string]$_.EmailAddress})}catch{}
    if($esActual){try{$crudos+=@(& "$env:SystemRoot\System32\whoami.exe" /upn 2>$null)}catch{}}
    return @($crudos|ForEach-Object{([string]$_).Trim().ToLowerInvariant()}|Where-Object{$_ -match '@(avon|garnier)\.cr$'}|Select-Object -Unique)
}
function Corporacion-De-Correo([string]$correo){
    if($correo -match '@empresa-a\.invalid$'){return 'EMPRESA A'}
    if($correo -match '@empresa-b\.invalid$'){return 'EMPRESA B'}
    return $null
}
function Solicitar-Administrativos($registros){
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host '  DATOS PARA CONSTANCIA WORD Y POWER AUTOMATE' -ForegroundColor Cyan
    Write-Host '========================================================' -ForegroundColor Cyan
    Write-Host 'Las tareas terminaron; complete ahora los datos administrativos.'
    $inventario=@($registros|Where-Object Id -eq 'Inventario'|Select-Object -First 1)
    $datos=if($inventario.Count){$inventario[0].Datos}else{$null}
    $usuarioDefecto=if($datos -and $datos.usuario_sesion){[string]$datos.usuario_sesion}else{''}
    $candidatos=@(Obtener-CorreosUsuario $usuarioDefecto)
    $correo=''
    if($candidatos.Count -eq 1){
        $correo=$candidatos[0]
        Write-Host ('Correo detectado: '+$correo) -ForegroundColor Green
    }elseif($candidatos.Count -gt 1){
        $correo=Solicitar-Opcion 'Se detectaron varios correos corporativos. Elija el del usuario' $candidatos
    }
    $corpsDetectadas=@($candidatos|ForEach-Object{Corporacion-De-Correo $_}|Select-Object -Unique)
    if($corpsDetectadas.Count -eq 1){
        $corp=[string]$corpsDetectadas[0]
        Write-Host ('Corporacion detectada: '+$corp) -ForegroundColor Green
    }else{
        $corp=Solicitar-Opcion 'Corporacion' @('EMPRESA A','EMPRESA B')
    }
    if(-not $candidatos.Count){$correo=Solicitar-Texto 'Correo del usuario (opcional)'}
    $prefijo=if($corp -eq 'EMPRESA B'){'GYG'}else{'AV'}
    while($true){
        $placa=Solicitar-Texto ('Placa del activo (ejemplo '+$prefijo+'-1234)') '' $true
        if($placa -match ('^'+$prefijo+'-[0-9]{4}$')){break}
        Write-Host 'Formato incorrecto. Escriba el prefijo, un guion y 4 numeros.' -ForegroundColor Yellow
    }
    $estado=Solicitar-Opcion 'Estado del equipo' @('En uso','En reparación','Fuera de Uso')
    $departamentos=if($corp -eq 'EMPRESA B'){
        @('Departamento de ingeniería','Departamento de Recursos Humanos','Departamento de finanzas','Departamento de contabilidad','Departamento de desarrollo/TI','Departamento Legal','Departamento Administrativo','Departamento de Mercadeo')
    }else{
        @('Departamento de Ventas','Departamento de Recursos Humanos','Departamento Gerencia General','Departamento de Contabilidad','Departamento de Diseño','Departamento de Digital','Departamento de Compras','Departamento de Categoria','Departamento de Servicio al Cliente','Departamento de Logistica','Departamento de Mercadeo')
    }
    $depto=Solicitar-Opcion 'Departamento (debe coincidir con SharePoint)' $departamentos
    $categoria='Equipo portátil'
    $usuario=Solicitar-Texto 'Usuario asignado' $usuarioDefecto $true
    $jefatura=Solicitar-Texto 'Jefatura inmediata (opcional)'
    # Garantia obligatoria: Enter asigna fecha centinela de garantia vencida.
    # La captura solo acepta dd/MM/yyyy; se convierte a ISO para el flujo y el reporte.
    $fechaIso=$null
    while($true){
        $texto=Solicitar-Texto 'Finalizacion de garantia (dd/MM/yyyy; Enter = vencida: 11/11/1111)'
        if(-not $texto){
            $texto='11/11/1111'
            Write-Host 'Se registra garantia vencida: 11/11/1111.' -ForegroundColor Yellow
        }
        $f=[datetime]::MinValue
        $cultura=[Globalization.CultureInfo]::InvariantCulture
        if($texto -cmatch '^[0-9]{2}/[0-9]{2}/[0-9]{4}$' -and
           [datetime]::TryParseExact($texto,'dd/MM/yyyy',$cultura,[Globalization.DateTimeStyles]::None,[ref]$f)){
            $fechaIso=$f.ToString('yyyy-MM-dd');break
        }
        Write-Host 'Fecha no valida. Use exactamente dd/MM/aaaa (ejemplo: 12/12/2026), o Enter para 11/11/1111.' -ForegroundColor Yellow
    }
    $accesorios=Solicitar-Texto 'Accesorios / placas / series (Enter = Sin accesorios)' 'Sin accesorios'
    $observaciones=Solicitar-Texto 'Observaciones (opcional)'
    return [pscustomobject]@{
        placa=$placa.ToUpperInvariant();estado=$estado;corporacion=$corp;categoria=$categoria;
        usuario=$usuario;correo=$correo;departamento=$depto;jefatura=$jefatura;
        garantia=$fechaIso;accesorios=$accesorios;observaciones=$observaciones;
        tecnico=([Security.Principal.WindowsIdentity]::GetCurrent().Name)
    }
}
function Crear-Payload($administrativos,$registros){
    if(-not $administrativos){return $null}
    $inv=@($registros|Where-Object Id -eq 'Inventario'|Select-Object -First 1)
    $ev=@($registros|Where-Object Id -eq 'Eventos'|Select-Object -First 1)
    $info=if($inv.Count){$inv[0].Datos}else{$null}
    $serie=if($info){[string]$info.serie}else{''}
    $fabricante=if($info){[string]$info.fabricante}else{''}
    $modelo=if($info){[string]$info.modelo}else{''}
    $os=if($info){[string]$info.sistema_operativo}else{''}
    $ram=if($info){[string]$info.ram_gb}else{''}
    $cpu=if($info){[string]$info.cpu}else{''}
    $d=if($info){($info.unidades|ConvertTo-Json -Compress -Depth 4)}else{''}
    $criticos=0;$errores=0
    if($ev.Count -and $ev[0].Datos -and $ev[0].Datos.registros){
        foreach($r in @($ev[0].Datos.registros)){$criticos+=[int]$r.criticos;$errores+=[int]$r.errores}
    }
    $fallas=@($registros|Where-Object {$_.Estado -in @('Error','Con incidencias','Tiempo agotado')})
    $res=if($fallas.Count){'Con incidencias'}else{'Completado'}
    $listaHallazgos=(@($fallas|ForEach-Object {$_.Nombre+': '+$_.Detalle}) -join ' | ')
    return [ordered]@{
        placa=[string]$administrativos.placa;estado=[string]$administrativos.estado;
        corporacion=[string]$administrativos.corporacion;categoria=[string]$administrativos.categoria;
        fabricante=$fabricante;modelo=$modelo;especificaciones=('CPU: '+$cpu+' | RAM: '+$ram+' GB | Unidades: '+$d);
        serie=$serie;garantia=$administrativos.garantia;usuario=[string]$administrativos.usuario;
        correo=[string]$administrativos.correo;departamento=[string]$administrativos.departamento;
        jefatura=[string]$administrativos.jefatura;accesorios=[string]$administrativos.accesorios;
        observaciones=[string]$administrativos.observaciones;sistema_operativo=$os;
        memoria_ram=$ram;almacenamiento=$d;fecha_mantenimiento=(Get-Date -Format 'yyyy-MM-dd');
        resultado_mantenimiento=$res;eventos_criticos=$criticos;errores_windows=$errores;
        diagnostico_eventos=('Criticos: '+$criticos+'; errores: '+$errores);
        hallazgos=$listaHallazgos;proximo_mantenimiento=$null;
        resultados_tareas=(@($registros|ForEach-Object {[ordered]@{tarea=$_.Id;estado=$_.Estado;detalle=$_.Detalle}})|ConvertTo-Json -Compress -Depth 4)
    }
}
function Enviar-A-PowerAutomate($datos,$registros){
    $respuesta=[ordered]@{estado='No configurado';detalle='No se envio informacion.';fecha=(Get-Date).ToString('o')}
    $ruta=Join-Path $root 'Configuracion.json'
    if(-not (Test-Path -LiteralPath $ruta)){$respuesta.detalle='No existe Configuracion.json.';return $respuesta}
    try{$config=Get-Content -LiteralPath $ruta -Raw -Encoding UTF8|ConvertFrom-Json -ErrorAction Stop}
    catch{$respuesta.estado='Error';$respuesta.detalle='Configuracion.json no valido: '+$_.Exception.Message;return $respuesta}
    if(-not $config.PowerAutomate -or -not $config.PowerAutomate.Habilitado){$respuesta.detalle='Envio deshabilitado en Configuracion.json.';return $respuesta}
    $url=[string]$config.PowerAutomate.UrlFlujo
    if(-not $url -or $url -notmatch '^https://'){$respuesta.estado='Error';$respuesta.detalle='Falta URL HTTPS del flujo nuevo en Configuracion.json.';return $respuesta}
    if(-not $datos){$respuesta.estado='Error';$respuesta.detalle='No existen datos administrativos.';return $respuesta}
    if(-not $datos.serie){$respuesta.estado='Error';$respuesta.detalle='No hay numero de serie; no se envia.';return $respuesta}
    if(-not $datos.garantia){
        $respuesta.estado='Omitida';$respuesta.detalle='Garantia vacia: SharePoint la exige en Create item. No se envio para evitar error.';return $respuesta
    }
    try{
        $segundos=45
        if($config.PowerAutomate.TiempoEsperaSegundos){$segundos=[math]::Max(10,[math]::Min(120,[int]$config.PowerAutomate.TiempoEsperaSegundos))}
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
        $json=$datos|ConvertTo-Json -Depth 9 -Compress
        $bytes=[Text.Encoding]::UTF8.GetBytes($json)
        Write-Host "`nEnviando inventario a Power Automate..." -ForegroundColor Cyan
        $r=Invoke-RestMethod -Uri $url -Method Post -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec $segundos -ErrorAction Stop
        if($null -ne $r -and $r.existe -eq $true){
            $respuesta.estado='Existente';$respuesta.detalle='El activo ya existe. No se sobrescribio ni creo duplicado.'
        }elseif($null -ne $r -and $r.creado -eq $true){
            $respuesta.estado='Creado';$respuesta.detalle='Power Automate confirmo creacion del activo.'
        }else{
            $respuesta.estado='Sin confirmar';$respuesta.detalle='Se obtuvo respuesta del flujo sin confirmacion explicita de creacion; verificar ejecucion.'
        }
    }catch{
        $respuesta.estado='Error';$respuesta.detalle=$_.Exception.Message
        # No revelar URL firmada en la consola ni en el Word.
        if($respuesta.detalle -match 'https?://'){$respuesta.detalle='Error HTTP al contactar el flujo. Revise historial de ejecuciones y la configuracion.'}
    }
    return $respuesta
}
function Generar-ReporteSupervisado($registros,$administrativos,$envio,[string]$tipo,[datetime]$inicio,[string]$temp){
    $trabajador=Join-Path $root 'Generar_Reporte.ps1'
    if(-not (Test-Path -LiteralPath $trabajador)){throw 'Falta Generar_Reporte.ps1 junto al BAT.'}
    $identificador=if($administrativos -and $administrativos.placa){$administrativos.placa}else{$env:COMPUTERNAME}
    $identificador=([string]$identificador -replace '[^a-zA-Z0-9_-]','_')
    $salida=Join-Path $root ('Reporte_Mantenimiento_'+$identificador+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.docx')
    $carga=[ordered]@{
        modo=$tipo;inicio=$inicio.ToString('o');fin=(Get-Date).ToString('o');equipo=$env:COMPUTERNAME;
        administrativo=$administrativos;power_automate=$envio;
        registros=@($registros|ForEach-Object {
            [ordered]@{id=$_.Id;nombre=$_.Nombre;estado=$_.Estado;porcentaje=$_.Porcentaje;
            detalle=$_.Detalle;inicio=if($_.Inicio){$_.Inicio.ToString('o')}else{$null};
            fin=if($_.Fin){$_.Fin.ToString('o')}else{$null};datos=$_.Datos}
        })
    }
    $archivoCarga=Join-Path $temp 'reporte_carga.json'
    [IO.File]::WriteAllText($archivoCarga,($carga|ConvertTo-Json -Depth 12),([Text.UTF8Encoding]::new($false)))
    $logOut=Join-Path $temp 'word.out';$logErr=Join-Path $temp 'word.err'
    $args='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$trabajador+'" -Entrada "'+$archivoCarga+'" -Destino "'+$salida+'" -Plantilla "'+$plantilla+'"'
    Write-Host "`nGenerando Word (espera maxima: 120 segundos)..." -ForegroundColor Cyan
    $p=Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $args -PassThru -WindowStyle Hidden -RedirectStandardOutput $logOut -RedirectStandardError $logErr
    try{
        if(-not $p.WaitForExit(120000)){
            try{& "$env:SystemRoot\System32\taskkill.exe" /PID $p.Id /T /F 2>$null|Out-Null}catch{}
            throw 'Word no respondio dentro de 120 segundos; revise si hay un dialogo oculto o plantilla bloqueada.'
        }
        if($p.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $salida)){
            $m=if(Test-Path -LiteralPath $logErr){[IO.File]::ReadAllText($logErr)}else{'No hay salida de error.'}
            if($m.Length -gt 600){$m=$m.Substring(0,600)}
            throw ('No se genero el documento. '+$m)
        }
        Write-Host ('REPORTE WORD GENERADO: '+$salida) -ForegroundColor Green
    }finally{try{$p.Dispose()}catch{}}
}
function Lineas-Tarea($r){
    $d=$r.Datos;$l=@()
    try{
        if($d){
            switch($r.Id){
                'Inventario'{
                    $l+=('Equipo: '+$d.equipo+' | '+$d.fabricante+' '+$d.modelo)
                    $l+=('Serie: '+$d.serie+' | SO: '+$d.sistema_operativo)
                    $l+=('CPU: '+$d.cpu+' | RAM: '+$d.ram_gb+' GB')
                }
                'Eventos'{
                    foreach($x in @($d.registros)){
                        $m=if(([int]$x.criticos+[int]$x.errores) -gt 0){'[!!] '}else{''}
                        $l+=($m+$x.registro+': '+$x.criticos+' criticos, '+$x.errores+' errores, '+$x.advertencias+' advertencias')
                    }
                }
                'Antivirus'{
                    if($d.producto){$l+=([string]$d.producto+' | tiempo real: '+$d.tiempo_real)}
                    if($d.firma_fecha){$l+=('Firmas actualizadas: '+$d.firma_fecha)}
                }
                'Firewall'{
                    foreach($p in @($d.perfiles)){
                        $act=([string]$p.Enabled -in @('1','True'))
                        $l+=($(if($act){'[OK] '}else{'[!!] '})+$p.Name+': '+$(if($act){'Activo'}else{'DESACTIVADO'}))
                    }
                }
                'Bateria'{
                    if($d.disponible){
                        $l+=('Carga actual: '+$d.porcentaje+'%')
                        if($null -ne $d.salud_porcentaje){
                            $m=if([double]$d.salud_porcentaje -lt 60){'[!!] '}else{''}
                            $l+=($m+'Salud estimada: '+$d.salud_porcentaje+'%')
                        }
                    }
                }
                'Temporales'{
                    $l+=('Eliminados: '+$d.archivos_eliminados+' de '+$d.archivos_evaluados+' archivos antiguos')
                    $l+=('Conservados (recientes): '+$d.recientes_conservados+' | carpetas vacias eliminadas: '+$d.carpetas_vacias_eliminadas)
                    if([int]$d.bloqueados -gt 0){$l+=('[!!] Bloqueados o no accesibles: '+$d.bloqueados)}
                }
                'Papelera'{
                    if($d.resumen){$l+=('Espacio liberado: '+$d.resumen.liberado_mb+' MB')}
                    foreach($u in @($d.unidades)){
                        $m=if($u.resultado -eq 'Vaciada' -or $u.resultado -eq 'Sin papelera'){''}else{'[!!] '}
                        $l+=($m+$u.unidad+' '+$u.resultado+' ('+$u.liberado_mb+' MB liberados)')
                    }
                }
                'Disco'{
                    foreach($u in @($d.unidades)){
                        $m=if($null -ne $u.ocupacion_porcentaje -and [double]$u.ocupacion_porcentaje -ge 90){'[!!] '}else{''}
                        $l+=($m+$u.unidad+' '+$u.libre_gb+' GB libres de '+$u.total_gb+' GB ('+$u.ocupacion_porcentaje+'% usado)')
                    }
                }
                'DistribucionAlmacenamiento'{
                    $l+=('Perfil: '+$d.perfil_total_gb+' GB | Unidad '+$d.unidad+': '+$d.unidad_libre_gb+' GB libres / '+$d.unidad_total_gb+' GB')
                    foreach($c in @($d.carpetas | Select-Object -First 12)){
                        $l+=(' - '+$c.carpeta+': '+$c.tamano_gb+' GB')
                    }
                    if($d.medicion_parcial -or $d.incidencias_acceso -gt 0){$l+='[!!] Analisis parcial o archivos no accesibles'}
                    $l+='Nota: uso del perfil, no desglose completo del disco.'
                }
                'Apps'{
                    $req=@($d.requeridas)
                    $ok=@($req|Where-Object{$_.estado -eq 'Instalada'}|ForEach-Object{$_.app})
                    $l+=('[OK] Instaladas ('+$ok.Count+'/'+$req.Count+'): '+($ok -join ', '))
                    foreach($f in @($req|Where-Object{$_.estado -ne 'Instalada'})){$l+=('[FALTA] '+$f.app+' no esta instalada')}
                    foreach($a in @($d.actualizaciones)){
                        $m=switch($a.nivel){'ok'{'[OK] '} 'aviso'{'[..] '} default{'[!!] '}}
                        $l+=($m+'Actualizacion '+$a.app+': '+$a.resultado)
                    }
                }
            }
        }
    }catch{$l=@()}
    if(-not $l.Count -and $r.Detalle){
        $primera=($r.Detalle -split "`n")[0].Trim()
        if($primera.Length -gt 90){$primera=$primera.Substring(0,87)+'...'}
        $l=@($primera)
    }
    return $l
}
function Mostrar-Reporte($lista,[datetime]$inicio){
    $dur=[timespan]::FromSeconds([math]::Max(0,[int]((Get-Date)-$inicio).TotalSeconds))
    $n=@{}
    foreach($r in $lista){$n[$r.Estado]=1+[int]$n[$r.Estado]}
    $problemas=[int]$n['Con incidencias']+[int]$n['Error']+[int]$n['Tiempo agotado']+[int]$n['Cancelada']
    $linea=('='*70)
    Write-Host ''
    Write-Host $linea -ForegroundColor Cyan
    Write-Host ('  REPORTE DE MANTENIMIENTO  |  '+$env:COMPUTERNAME+'  |  '+(Get-Date -Format 'dd/MM/yyyy HH:mm')) -ForegroundColor Cyan
    Write-Host ('  Duracion: {0:hh\:mm\:ss}' -f $dur) -ForegroundColor Cyan
    Write-Host $linea -ForegroundColor Cyan
    Write-Host ('  Tareas: {0}   ' -f $lista.Count) -NoNewline
    Write-Host ('Completadas: {0}   ' -f [int]$n['Completada']) -ForegroundColor Green -NoNewline
    Write-Host ('Con problemas: {0}   ' -f $problemas) -ForegroundColor $(if($problemas){'Yellow'}else{'Gray'}) -NoNewline
    Write-Host ('Omitidas: {0}' -f [int]$n['Omitida']) -ForegroundColor DarkYellow
    $secciones=[ordered]@{
        'EQUIPO'=@('Inventario','Permisos','TrabajoEscuela','Intune')
        'SEGURIDAD'=@('Antivirus','Firewall','Cuentas')
        'SALUD DEL EQUIPO'=@('Bateria','Disco','DistribucionAlmacenamiento','Eventos')
        'LIMPIEZA'=@('Temporales','Papelera')
        'APLICACIONES'=@('Apps')
    }
    $usados=@($secciones.Values|ForEach-Object{$_})
    $otras=@($lista|Where-Object{$_.Id -notin $usados}|ForEach-Object{$_.Id})
    if($otras.Count){$secciones['OTRAS']=$otras}
    foreach($titulo in $secciones.Keys){
        $items=@($lista|Where-Object{$_.Id -in $secciones[$titulo]})
        if(-not $items.Count){continue}
        Write-Host ''
        Write-Host ('  --- '+$titulo+' '+('-'*(60-$titulo.Length))) -ForegroundColor Cyan
        foreach($r in $items){
            $marca=switch($r.Estado){'Completada'{'[OK]'} 'Omitida'{'[--]'} 'Con incidencias'{'[!!]'} 'Tiempo agotado'{'[!!]'} default{'[XX]'}}
            $color=switch($r.Estado){'Completada'{'Green'} 'Omitida'{'DarkYellow'} 'Con incidencias'{'Yellow'} 'Tiempo agotado'{'Yellow'} default{'Red'}}
            Write-Host ('  {0} {1,-42} {2}' -f $marca,$r.Nombre,$r.Estado) -ForegroundColor $color
            foreach($t in @(Lineas-Tarea $r)){
                $c=if($t -match '^\[(XX|FALTA)\]'){'Red'}elseif($t -match '^\[!!\]'){'Yellow'}elseif($t -match '^\[OK\]'){'Green'}else{'Gray'}
                Write-Host ('        '+$t) -ForegroundColor $c
            }
        }
    }
    $atencion=@($lista|Where-Object{$_.Estado -in @('Con incidencias','Error','Tiempo agotado','Cancelada')})
    Write-Host ''
    Write-Host $linea -ForegroundColor Cyan
    if($atencion.Count){
        Write-Host '  REQUIERE ATENCION' -ForegroundColor Yellow
        foreach($r in $atencion){
            $d=if($r.Detalle){($r.Detalle -split "`n")[0].Trim()}else{''}
            if($d.Length -gt 80){$d=$d.Substring(0,77)+'...'}
            Write-Host ('   - '+$r.Nombre+': '+$d) -ForegroundColor Yellow
        }
    }else{Write-Host '  Todo en orden: no hay tareas que requieran atencion.' -ForegroundColor Green}
    Write-Host $linea -ForegroundColor Cyan
}
function Abrir-EdgeAutomatico {
    $url='https://teams.public.onecdn.static.microsoft/evergreen-assets/safelinks/2/atp-safelinks.html'
    $candidatos=@()
    foreach($base in @(${env:ProgramFiles(x86)},$env:ProgramFiles)){
        if($base){$candidatos+=Join-Path $base 'Microsoft\Edge\Application\msedge.exe'}
    }
    $edge=@($candidatos | Where-Object {$_ -and (Test-Path -LiteralPath $_)} | Select-Object -First 1)
    try {
        if($edge.Count -gt 0){Start-Process -FilePath $edge[0] -ArgumentList $url -ErrorAction Stop | Out-Null}
        else{Start-Process -FilePath 'msedge.exe' -ArgumentList $url -ErrorAction Stop | Out-Null}
        Write-Host 'Pagina solicitada abierta en Microsoft Edge.' -ForegroundColor Cyan
    }catch{
        Write-Warning ('No se pudo abrir Edge: '+$_.Exception.Message)
    }
}
function Ejecutar-Tareas([string[]]$ids,[string]$tipo){
    $inicio=Get-Date
    $temp=Join-Path ([IO.Path]::GetTempPath()) ('InventarioTI_'+[guid]::NewGuid().ToString('N'))
    [void](New-Item -Path $temp -ItemType Directory -Force)
    $lista=@($ids|ForEach-Object {Nuevo-Registro $_ $temp})
    $mapa=@{};foreach($r in $lista){$mapa[$r.Id]=$r}
    $admin=$null
    $envio=[ordered]@{estado='No realizado';detalle='Modo manual: no se envia a SharePoint.'}
    try{
        if($tipo -eq 'Automatico'){Abrir-EdgeAutomatico}
        # RevocarAdmin es la ultima tarea: no debe correr en paralelo con tareas elevadas.
        foreach($r in $lista){
            if($tipo -eq 'Automatico' -and $r.Id -in @('Papelera','DistribucionAlmacenamiento','RevocarAdmin')){continue}
            Iniciar-Tarea $r
        }
        while($true){
            foreach($r in $lista){try{Actualizar-Tarea $r}catch{$r.Estado='Error';$r.Detalle=$_.Exception.Message;$r.Fin=Get-Date}}
            if($tipo -eq 'Automatico' -and $mapa.ContainsKey('Papelera') -and $mapa.ContainsKey('Temporales')){
                $p=$mapa['Papelera'];$t=$mapa['Temporales']
                if($p.Estado -eq 'Pendiente' -and $t.Estado -notin @('Pendiente','En ejecucion')){
                    if(((Get-Date)-$inicio).TotalSeconds -lt $maxSegundos){Iniciar-Tarea $p}
                }
            }
            # Analizar almacenamiento despues de las tareas de limpieza, para medir el estado final.
            if($tipo -eq 'Automatico' -and $mapa.ContainsKey('DistribucionAlmacenamiento')){
                $alm=$mapa['DistribucionAlmacenamiento']
                $otrasAlm=@($lista|Where-Object {$_.Id -notin @('DistribucionAlmacenamiento','RevocarAdmin') -and $_.Estado -in @('Pendiente','En ejecucion')})
                if($alm.Estado -eq 'Pendiente' -and $otrasAlm.Count -eq 0 -and ((Get-Date)-$inicio).TotalSeconds -lt $maxSegundos){
                    Iniciar-Tarea $alm
                }
            }
            # Iniciar la revocacion solo despues de terminar todas las otras tareas.
            # Asi no se pierden permisos mientras se realizan cambios o comprobaciones.
            if($tipo -eq 'Automatico' -and $mapa.ContainsKey('RevocarAdmin')){
                $rev=$mapa['RevocarAdmin']
                $otras=@($lista|Where-Object {$_.Id -ne 'RevocarAdmin' -and $_.Estado -in @('Pendiente','En ejecucion')})
                if($rev.Estado -eq 'Pendiente' -and $otras.Count -eq 0 -and ((Get-Date)-$inicio).TotalSeconds -lt $maxSegundos){
                    Iniciar-Tarea $rev
                }
            }
            Pintar-Pantalla $lista $inicio $tipo
            if(@($lista|Where-Object {$_.Estado -in @('Pendiente','En ejecucion')}).Count -eq 0){break}
            if($tipo -eq 'Automatico' -and ((Get-Date)-$inicio).TotalSeconds -ge $maxSegundos){
                foreach($r in $lista|Where-Object {$_.Estado -in @('Pendiente','En ejecucion')}){
                    if($r.Estado -eq 'En ejecucion'){Detener-Tarea $r}
                    $r.Estado='Tiempo agotado';$r.Detalle='No finalizo dentro del limite de 35 minutos.';$r.Fin=Get-Date
                };break
            }
            Start-Sleep -Milliseconds 800
        }
        Mostrar-Reporte $lista $inicio
        if($tipo -eq 'Automatico'){
            try{
                $admin=Solicitar-Administrativos $lista
                $payload=Crear-Payload $admin $lista
                $envio=Enviar-A-PowerAutomate $payload $lista
                Write-Host ('Power Automate: '+$envio.estado+' - '+$envio.detalle) -ForegroundColor Yellow
            }catch{
                $envio=[ordered]@{estado='Error';detalle='Fallo en datos administrativos o envio: '+$_.Exception.Message}
                Write-Host ('No se pudo completar integracion: '+$envio.detalle) -ForegroundColor Yellow
            }
        }
    }finally{
        foreach($r in $lista){if($r.Estado -eq 'En ejecucion'){Detener-Tarea $r;$r.Estado='Cancelada';$r.Detalle='Supervisor interrumpido.';$r.Fin=Get-Date}}
        try{Generar-ReporteSupervisado $lista $admin $envio $tipo $inicio $temp}
        catch{Write-Host ('ERROR AL GENERAR WORD: '+$_.Exception.Message) -ForegroundColor Red}
        foreach($r in $lista){if($r.Proceso){try{$r.Proceso.Dispose()}catch{}}}
        try{Remove-Item -LiteralPath $temp -Force -Recurse -ErrorAction Stop}catch{Write-Warning 'No se pudo borrar carpeta temporal de esta ejecucion.'}
    }
}
try{
    if($Modo -eq 'Automatico'){Ejecutar-Tareas @($catalogo.Keys) 'Automatico';return}
    if($Modo -eq 'Manual'){
        if(-not $catalogo.Contains($Tarea)){throw ('Tarea manual no reconocida: '+$Tarea)}
        Ejecutar-Tareas @($Tarea) 'Manual';return
    }
    while($true){
        Write-Host "`n=== INVENTARIO Y MANTENIMIENTO TI ===" -ForegroundColor Cyan
        Write-Host '1. Automatico: tareas paralelas, maximo 35 minutos; luego formulario y Power Automate'
        Write-Host '2. Manual: UNA tarea; solo Word, sin envio'
        Write-Host '0. Salir'
        $s=Read-Host 'Seleccion'
        if($s -eq '0'){break}
        if($s -eq '1'){Ejecutar-Tareas @($catalogo.Keys) 'Automatico';continue}
        if($s -eq '2'){
            $ids=@($catalogo.Keys)
            Write-Host "`n=== SELECCION DE TAREA ===" -ForegroundColor Cyan
            for($i=0;$i -lt $ids.Count;$i++){Write-Host ('{0,2}. {1}' -f ($i+1),$catalogo[$ids[$i]])}
            Write-Host ' 0. Volver'
            $numero=0;$s2=Read-Host 'Numero de tarea'
            if($s2 -eq '0'){continue}
            if([int]::TryParse($s2,[ref]$numero) -and $numero -ge 1 -and $numero -le $ids.Count){Ejecutar-Tareas @($ids[$numero-1]) 'Manual'}
            else{Write-Warning 'Seleccion no valida.'}
            continue
        }
        Write-Warning 'Seleccion no valida.'
    }
}catch{Write-Host ('ERROR DEL CONTROLADOR: '+$_.Exception.Message) -ForegroundColor Red}