#Requires -Version 5.1

<#
.SYNOPSIS
    Entra ID - Gestion de Extension Attributes en dispositivos
.DESCRIPTION
    Carga todos los dispositivos de Entra ID, permite seleccionar uno o varios
    mediante Out-GridView y asigna/borra cualquier extensionAttribute (1-15).
    Guarda log persistente y exporta CSV del inventario.
.NOTES
    Requiere: Microsoft.Graph >= 2.x
    Permisos:  Directory.AccessAsUser.All (delegado)
    Compatibilidad: PowerShell 5.1 y superior
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuracion global

$script:LogPath   = Join-Path $PSScriptRoot "EntraDeviceManager_$(Get-Date -Format 'yyyyMMdd').log"
$script:UserIndex = @{}   # hashtable userId -> UPN, cargado una sola vez al inicio
$script:Stats     = @{ OK = 0; Error = 0; Start = (Get-Date) }

#endregion

#region Logging

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    $line | Add-Content -Path $script:LogPath -ErrorAction SilentlyContinue

    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red    }
        'OK'    { Write-Host $line -ForegroundColor Green  }
        default { Write-Host $line -ForegroundColor Cyan   }
    }
}

#endregion

#region Modulo y conexion

function Assert-GraphModule {
    if (-not (Get-InstalledModule -Name Microsoft.Graph -ErrorAction SilentlyContinue)) {
        Write-Log "Modulo Microsoft.Graph no encontrado. Instalando..." WARN
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Write-Log "Modulo Microsoft.Graph cargado." OK
}

function Connect-ToGraph {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    Write-Log "Iniciando autenticacion en Microsoft Graph..."
    Connect-MgGraph -Scopes "Directory.AccessAsUser.All" -NoWelcome -ErrorAction Stop

    $ctx = Get-MgContext
    Write-Log "Conectado como: $($ctx.Account) | Tenant: $($ctx.TenantId)" OK
}

#endregion

#region Carga masiva de usuarios (una sola llamada a la API)

function Import-UserIndex {
    Write-Log "Cargando indice de usuarios..."
    $users = Get-MgUser -All -Property Id,UserPrincipalName -ErrorAction Stop
    $index = @{}
    foreach ($u in $users) {
        $index[$u.Id] = $u.UserPrincipalName
    }
    Write-Log "$($index.Count) usuarios indexados." OK
    return $index
}

#endregion

#region Helpers

function Get-UserIdFromPhysicalIds {
    param([string[]]$PhysicalIds)
    foreach ($id in $PhysicalIds) {
        if ($id.Length -ge 47 -and $id.StartsWith('[USER-GID]')) {
            return $id.Substring(11, 36)
        }
    }
    return $null
}

function Get-ExtAttr {
    # En Graph SDK v2 las propiedades no nativas llegan en AdditionalProperties
    param($Device, [int]$Num)
    if ($Device.AdditionalProperties -and
        $Device.AdditionalProperties.ContainsKey('extensionAttributes')) {
        $dict    = $Device.AdditionalProperties['extensionAttributes']
        $attrKey = "extensionAttribute$Num"
        if ($dict -and $dict.ContainsKey($attrKey)) { return $dict[$attrKey] }
    }
    return $null
}

#endregion

#region Inventario de dispositivos

function Get-DeviceInventory {
    param([hashtable]$UserIndex)

    Write-Log "Cargando inventario de dispositivos..."
    $properties = 'Id,DisplayName,OperatingSystem,OperatingSystemVersion,AccountEnabled,PhysicalIds,extensionAttributes'
    $devices    = Get-MgDevice -All -Property $properties -ErrorAction Stop
    Write-Log "$($devices.Count) dispositivos encontrados." OK

    $i   = 0
    $inv = foreach ($dev in $devices) {
        $i++
        if ($i % 100 -eq 0) {
            Write-Host "  Procesando $i / $($devices.Count)..." -ForegroundColor DarkGray
        }

        $userId = $null
        $upn    = $null

        if ($dev.PhysicalIds -and $dev.PhysicalIds.Count -gt 0) {
            $userId = Get-UserIdFromPhysicalIds -PhysicalIds $dev.PhysicalIds
            # Lookup local en el hashtable: sin llamada a la API
            if ($userId -and $UserIndex.ContainsKey($userId)) {
                $upn = $UserIndex[$userId]
            }
        }

        [PSCustomObject]@{
            Id                   = $dev.Id
            Nombre               = $dev.DisplayName
            SO                   = $dev.OperatingSystem
            Version              = $dev.OperatingSystemVersion
            Enabled              = $dev.AccountEnabled
            UsuarioId            = $userId
            'Nombre usuario'     = $upn
            ExtensionAttribute1  = Get-ExtAttr $dev 1
            ExtensionAttribute2  = Get-ExtAttr $dev 2
            ExtensionAttribute3  = Get-ExtAttr $dev 3
            ExtensionAttribute4  = Get-ExtAttr $dev 4
            ExtensionAttribute5  = Get-ExtAttr $dev 5
            ExtensionAttribute6  = Get-ExtAttr $dev 6
            ExtensionAttribute7  = Get-ExtAttr $dev 7
            ExtensionAttribute8  = Get-ExtAttr $dev 8
            ExtensionAttribute9  = Get-ExtAttr $dev 9
            ExtensionAttribute10 = Get-ExtAttr $dev 10
            ExtensionAttribute11 = Get-ExtAttr $dev 11
            ExtensionAttribute12 = Get-ExtAttr $dev 12
            ExtensionAttribute13 = Get-ExtAttr $dev 13
            ExtensionAttribute14 = Get-ExtAttr $dev 14
            ExtensionAttribute15 = Get-ExtAttr $dev 15
        }
    }
    return $inv
}

function Export-InventoryToCsv {
    param([PSCustomObject[]]$Inventory)
    $csvPath = Join-Path $PSScriptRoot "Inventario_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Inventory | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "Inventario exportado a: $csvPath" OK
}

#endregion

#region Actualizacion de atributos

function Invoke-AttributeUpdate {
    param(
        [PSCustomObject[]]$Devices,
        [int]$AttributeNumber,
        [string]$AttributeValue
    )

    $action = if ($AttributeValue -eq '') { 'BORRAR' } else { "= '$AttributeValue'" }
    Write-Host ""
    Write-Host "  Vas a modificar extensionAttribute$AttributeNumber $action" -ForegroundColor White
    Write-Host "  Afecta a $($Devices.Count) dispositivo(s):" -ForegroundColor White
    $Devices | ForEach-Object { Write-Host "    - $($_.Nombre)" -ForegroundColor Gray }

    $confirm = ''
    do { $confirm = (Read-Host "`n  Confirmar (S/N)").ToUpper().Trim() }
    while ($confirm -notin @('S','N'))

    if ($confirm -eq 'N') {
        Write-Log "Operacion cancelada por el usuario." WARN
        return
    }

    $body = @{
        extensionAttributes = @{
            "extensionAttribute$AttributeNumber" = $AttributeValue
        }
    } | ConvertTo-Json -Depth 3

    foreach ($dev in $Devices) {
        try {
            Update-MgDevice -DeviceId $dev.Id -BodyParameter $body -ErrorAction Stop
            Write-Log "OK | $($dev.Nombre) ($($dev.Id)) | attr$AttributeNumber $action" OK
            $script:Stats.OK++
        }
        catch {
            Write-Log "ERROR | $($dev.Nombre) ($($dev.Id)) | $_" ERROR
            $script:Stats.Error++
        }
    }
}

#endregion

#region Bucle principal

function Start-MainLoop {
    param(
        [PSCustomObject[]]$Inventory,
        [hashtable]$UserIndex
    )

    do {
        $seleccion = $Inventory | Out-GridView -Title "Selecciona uno o mas dispositivos -- $($Inventory.Count) total" -PassThru

        if (-not $seleccion) {
            Write-Log "No se selecciono ningun dispositivo." WARN
        }
        else {
            $numero = 0
            do {
                $inputNum = Read-Host "`nNumero de atributo a modificar (1-15)"
                $numero   = $inputNum -as [int]
            } while ($numero -notin 1..15)

            $atributo = Read-Host "Nuevo valor (dejar vacio para borrar el atributo)"

            Invoke-AttributeUpdate -Devices $seleccion -AttributeNumber $numero -AttributeValue $atributo

            Write-Log "Refrescando inventario..."
            $Inventory = Get-DeviceInventory -UserIndex $UserIndex
        }

        $continuar = ''
        do { $continuar = (Read-Host "`nOtra operacion? (S/N)").ToUpper().Trim() }
        while ($continuar -notin @('S','N'))

    } while ($continuar -eq 'S')
}

#endregion

#region Resumen final

function Show-Summary {
    $elapsed = (Get-Date) - $script:Stats.Start
    Write-Host ""
    Write-Host "-------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Resumen de sesion"                  -ForegroundColor White
    Write-Host "  Actualizaciones OK : $($script:Stats.OK)"    -ForegroundColor Green
    if ($script:Stats.Error -gt 0) {
        Write-Host "  Errores            : $($script:Stats.Error)" -ForegroundColor Red
    } else {
        Write-Host "  Errores            : 0"                      -ForegroundColor Gray
    }
    Write-Host "  Tiempo total       : $([math]::Round($elapsed.TotalSeconds,1))s" -ForegroundColor Gray
    Write-Host "  Log guardado en    : $script:LogPath"                            -ForegroundColor Gray
    Write-Host "-------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region Entry point

try {
    Assert-GraphModule
    Connect-ToGraph

    # Dos llamadas a la API en total para cargar todo el inventario
    $userIndex = Import-UserIndex
    $inventory = Get-DeviceInventory -UserIndex $userIndex

    $exportar = (Read-Host "`nExportar inventario a CSV antes de continuar? (S/N)").ToUpper().Trim()
    if ($exportar -eq 'S') {
        Export-InventoryToCsv -Inventory $inventory
    }

    Start-MainLoop -Inventory $inventory -UserIndex $userIndex
}
catch {
    Write-Log "Error critico: $_" ERROR
}
finally {
    Show-Summary
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Log "Sesion cerrada." INFO
}

#endregion
