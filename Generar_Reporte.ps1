#requires -version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Entrada,
    [Parameter(Mandatory=$true)][string]$Destino,
    [Parameter(Mandatory=$true)][string]$Plantilla
)
$ErrorActionPreference='Stop'
function Rellenar-Control($doc,[string]$tag,[string]$valor){
    if($null -eq $valor){$valor=''}
    try{
        $controles=$doc.SelectContentControlsByTag($tag)
        for($i=1;$i -le $controles.Count;$i++){
            try{$controles.Item($i).Range.Text=[string]$valor}
            catch{Write-Warning ('Campo Word '+$tag+': '+$_.Exception.Message)}
        }
    }catch{}
}
function Dato($obj,[string]$prop){
    if($null -eq $obj){return ''}
    if($obj.PSObject.Properties[$prop]){return [string]$obj.$prop}
    return ''
}
$word=$null;$documento=$null
try {
    if(-not (Test-Path -LiteralPath $Entrada)){throw 'No se encuentra archivo interno de datos del reporte.'}
    $carga=Get-Content -LiteralPath $Entrada -Raw -Encoding UTF8|ConvertFrom-Json
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false
    $word.DisplayAlerts=0
    try{$word.AutomationSecurity=3}catch{}
    if(Test-Path -LiteralPath $Plantilla){$documento=$word.Documents.Open($Plantilla,$false,$false)}
    else{$documento=$word.Documents.Add()}
    $ad=$carga.administrativo
    $reg=@($carga.registros)
    $inv=@($reg|Where-Object id -eq 'Inventario'|Select-Object -First 1)
    $hard=if($inv.Count){$inv[0].datos}else{$null}
    $usuario=if($ad){Dato $ad 'usuario'}else{Dato $hard 'usuario_sesion'}
    $departamento=Dato $ad 'departamento'
    $corp=Dato $ad 'corporacion'
    $correo=Dato $ad 'correo'
    $placa=Dato $ad 'placa'
    $fecha=[datetime]$carga.inicio
    Rellenar-Control $documento 'd1001' $fecha.ToString('dd/MM/yyyy')
    Rellenar-Control $documento 'd1040' $fecha.ToString('dd/MM/yyyy')
    Rellenar-Control $documento 't1002' $usuario
    Rellenar-Control $documento 't1003' $departamento
    Rellenar-Control $documento 't1004' $corp
    Rellenar-Control $documento 't1005' $correo
    Rellenar-Control $documento 't1006' (Dato $hard 'fabricante')
    Rellenar-Control $documento 't1007' (Dato $hard 'modelo')
    Rellenar-Control $documento 't1008' (Dato $hard 'serie')
    Rellenar-Control $documento 't1009' $placa
    Rellenar-Control $documento 't1010' (Dato $hard 'sistema_operativo')
    Rellenar-Control $documento 't1042' $(if($ad){Dato $ad 'tecnico'}else{[string]$env:USERNAME})
    $fallas=@($reg|Where-Object {$_.estado -in @('Error','Con incidencias','Tiempo agotado')})
    $tagsH=@('t1032','t1033','t1034')
    for($i=0;$i -lt $tagsH.Count;$i++){
        $v=if($i -lt $fallas.Count){$fallas[$i].nombre+': '+$fallas[$i].detalle}else{'Sin otras incidencias registradas.'}
        if($v.Length -gt 160){$v=$v.Substring(0,157)+'...'}
        Rellenar-Control $documento $tagsH[$i] $v
    }
    $recs=@('Revisar errores y omisiones del anexo tecnico antes del cierre.',
            'Validar con TI las cuentas administradas y politicas de seguridad.',
            'Verificar que el estado de Power Automate sea confirmado.',
            'Mantener registro de actualizaciones segun politica corporativa.',
            'Programar el siguiente mantenimiento preventivo.')
    $tagsR=@('t1035','t1036','t1037','t1038','t1039')
    for($i=0;$i -lt $tagsR.Count;$i++){Rellenar-Control $documento $tagsR[$i] $recs[$i]}

    $lineas=New-Object System.Collections.Generic.List[string]
    $lineas.Add('ANEXO TECNICO - INVENTARIO Y MANTENIMIENTO TI')
    $lineas.Add('Equipo: '+[string]$carga.equipo)
    $lineas.Add('Modo: '+[string]$carga.modo)
    $lineas.Add('Inicio: '+([datetime]$carga.inicio).ToString('yyyy-MM-dd HH:mm:ss'))
    $lineas.Add('Fin: '+([datetime]$carga.fin).ToString('yyyy-MM-dd HH:mm:ss'))
    $lineas.Add('El tiempo maximo de 35 minutos se aplica a las tareas automaticas.')
    $lineas.Add('')
    $lineas.Add('DATOS ADMINISTRATIVOS')
    if($ad){
        foreach($p in @('placa','usuario','correo','departamento','jefatura','corporacion','categoria','estado','garantia','accesorios','observaciones','tecnico')){
            $lineas.Add($p+': '+(Dato $ad $p))
        }
    }else{$lineas.Add('Modo manual: no se solicitaron datos administrativos.')}
    $lineas.Add('')
    $lineas.Add('INTEGRACION CON POWER AUTOMATE')
    $lineas.Add('Estado: '+(Dato $carga.power_automate 'estado'))
    $lineas.Add('Detalle: '+(Dato $carga.power_automate 'detalle'))
    $lineas.Add('No se incluye la URL firmada del flujo ni contrasenas.')
    $lineas.Add('')
    $lineas.Add('RESUMEN DE EJECUCION')
    foreach($r in $reg){
        $lineas.Add(([string]$r.nombre)+' - '+([string]$r.estado)+' - '+([int]$r.porcentaje)+'%')
        if($r.detalle){$lineas.Add('Detalle: '+[string]$r.detalle)}
    }
    $lineas.Add('')
    $lineas.Add('DETALLES POR TAREA')
    foreach($r in $reg){
        $lineas.Add('')
        $lineas.Add(([string]$r.nombre).ToUpperInvariant())
        $lineas.Add('Estado: '+[string]$r.estado)
        $lineas.Add('Progreso informado: '+[string]$r.porcentaje+'%')
        $lineas.Add('Detalle: '+[string]$r.detalle)
        if($null -ne $r.datos){
            $serializado=$r.datos|ConvertTo-Json -Depth 10
            if($serializado.Length -gt 8000){$serializado=$serializado.Substring(0,8000)+' [resumen truncado]'}
            foreach($l in ($serializado -split "`r?`n")){$lineas.Add($l)}
        }
    }
    $lineas.Add('')
    $lineas.Add('Seguridad: sin contrasenas en el reporte; cuentas locales omitidas si no se aprovisionaron.')
    $lineas.Add('Excluidos de esta version: OneDrive, Windows Update y Microsoft Update.')
    if(Test-Path -LiteralPath $Plantilla){
        $rango=$documento.Range($documento.Content.End-1,$documento.Content.End-1)
        $rango.InsertBreak(7)
        $rango=$documento.Range($documento.Content.End-1,$documento.Content.End-1)
        $rango.InsertAfter(($lineas -join "`r"))
    }else{$documento.Content.Text=($lineas -join "`r")}
    $documento.SaveAs2($Destino,16)
    if(-not (Test-Path -LiteralPath $Destino)){throw 'Word no genero el documento esperado.'}
    Write-Output ('OK: '+$Destino)
}catch{
    [Console]::Error.WriteLine('WORD ERROR: '+$_.Exception.Message)
    exit 1
}finally{
    if($documento){try{$documento.Close(0)}catch{};try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($documento)}catch{}}
    if($word){try{$word.Quit()}catch{};try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}catch{}}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
