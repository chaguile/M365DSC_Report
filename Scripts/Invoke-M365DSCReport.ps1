#Requires -Version 5.1
<#
    Microsoft365DSC Report - Orquestador unico
    ===========================================

    Un solo script que ejecuta, de principio a fin, el proceso de:
      1. Exportar la configuracion DSC de cada tenant de Microsoft 365.
      2. Generar un reporte HTML comparativo (baseline) entre tenants.

    Presenta un menu inicial que detecta automaticamente en que paso del
    proceso te encuentras y sugiere el siguiente. Cada paso operativo se
    lanza en una sesion de PowerShell nueva y limpia (proceso hijo), tal y
    como exige Microsoft365DSC para evitar conflictos de ensamblados entre
    Microsoft.Graph, Az y PnP.PowerShell.

    Autor : Christian Aguilera
    Web    : https://microsoft365dsc.com/  |  https://export.microsoft365dsc.com/

    USO
    ---
      # Menu interactivo (recomendado):
      .\Invoke-M365DSCReport.ps1

      # Ejecutar un paso concreto directamente (sin menu):
      .\Invoke-M365DSCReport.ps1 -Step Setup
      .\Invoke-M365DSCReport.ps1 -Step Provision
      .\Invoke-M365DSCReport.ps1 -Step Export
      .\Invoke-M365DSCReport.ps1 -Step Report
      .\Invoke-M365DSCReport.ps1 -Step Remove

      # Cambiar la carpeta raiz de trabajo (por defecto C:\M365DSC):
      .\Invoke-M365DSCReport.ps1 -Root "D:\M365DSC"
#>

param(
    [ValidateSet('Menu','Setup','Provision','Export','Report','Remove')]
    [string]$Step = 'Menu',

    [string]$Root = 'C:\M365DSC',

    # Idioma de la interfaz: ES (Espanol) o EN (English). Si se omite, el menu
    # lo pregunta al inicio. Los sub-pasos lo reciben como parametro.
    [ValidateSet('ES','EN','')]
    [string]$Lang = ''
)

# Idioma activo para toda la sesion (los procesos hijo lo reciben por -Lang)
$script:Lang = if ($Lang) { $Lang } else { 'ES' }

# Devuelve el texto en el idioma activo:  tr "en espanol" "in english"
function tr { param([string]$Es, [string]$En) if ($script:Lang -eq 'EN') { $En } else { $Es } }

# ============================================================
#  AYUDANTES COMPARTIDOS (usados por el menu)
# ============================================================
function Write-Step { param($m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  [X]  $m" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $yes = if ($script:Lang -eq 'EN') { 'Y' } else { 'S' }
    $d = if ($Default) { $yes } else { 'N' }
    while ($true) {
        $v = Read-Host "$Prompt ($yes/N) [$d]"
        if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
        switch ($v.Trim().ToUpper()) {
            'S' { return $true }
            'Y' { return $true }
            'N' { return $false }
            default { Write-Warn (tr "Responde S o N" "Answer Y or N") }
        }
    }
}

# Ruta del ejecutable de PowerShell que esta corriendo ahora mismo.
function Get-PsHost {
    try { $p = (Get-Process -Id $PID).Path; if ($p) { return $p } } catch { }
    if ($PSVersionTable.PSVersion.Major -ge 6) { return 'pwsh' } else { return 'powershell' }
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $pr = New-Object Security.Principal.WindowsPrincipal($id)
        return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

# Relanza este mismo script en un proceso PowerShell NUEVO y limpio para
# ejecutar un paso concreto. Comparte la consola actual (Read-Host funciona).
function Invoke-ChildStep {
    param([string]$Name)
    $exe = Get-PsHost
    $argline = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Step $Name -Root `"$Root`" -Lang $script:Lang"
    Write-Host ""
    Write-Host (tr "  Lanzando '$Name' en una sesion limpia de PowerShell..." "  Launching '$Name' in a clean PowerShell session...") -ForegroundColor DarkCyan
    Start-Process -FilePath $exe -ArgumentList $argline -NoNewWindow -Wait
}

# Pregunta el idioma al inicio (solo en modo menu)
function Select-Language {
    Clear-Host
    Write-Host ("=" * 68) -ForegroundColor Cyan
    Write-Host "  MICROSOFT 365 DSC - BASELINE REPORT" -ForegroundColor Cyan
    Write-Host "  Creado por / Created by: Christian Aguilera - FendariGroup" -ForegroundColor DarkCyan
    Write-Host ("=" * 68) -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   Selecciona idioma / Select language:" -ForegroundColor White
    Write-Host "     1) Espanol" -ForegroundColor Gray
    Write-Host "     2) English" -ForegroundColor Gray
    Write-Host ""
    $v = (Read-Host "   Opcion / Option [1]").Trim()
    if ($v -eq '2') { return 'EN' } else { return 'ES' }
}


# ============================================================
#  PASO 0 - PREPARACION DEL ENTORNO
# ============================================================
function Invoke-SetupStep {
    $ErrorActionPreference = 'Stop'

    Clear-Host
    Write-Host ("=" * 66) -ForegroundColor Cyan
    Write-Host (tr " PREPARACION DEL ENTORNO - MICROSOFT365DSC" " ENVIRONMENT SETUP - MICROSOFT365DSC") -ForegroundColor Cyan
    Write-Host ("=" * 66) -ForegroundColor Cyan

    # --- Carpetas ---
    Write-Step (tr "Creando la estructura de carpetas" "Creating the folder structure")
    $folders = @(
        (Join-Path $Root 'Scripts')
        (Join-Path $Root 'Export')
        (Join-Path $Root 'Reports')
        (Join-Path $Root 'Tenants\Baseline')
        (Join-Path $Root 'Tenants\SnapshotA')
        (Join-Path $Root 'Tenants\SnapshotB')
    )
    foreach ($f in $folders) {
        if (Test-Path $f) { Write-Ok "$f $(tr '(ya existe)' '(already exists)')" }
        else { New-Item -ItemType Directory -Force -Path $f | Out-Null; Write-Ok "$f $(tr '(creada)' '(created)')" }
    }

    # --- Modulo Microsoft365DSC ---
    Write-Step (tr "Modulo Microsoft365DSC" "Microsoft365DSC module")
    $mod = Get-Module -ListAvailable -Name Microsoft365DSC | Sort-Object Version -Descending | Select-Object -First 1
    if ($mod) {
        Write-Ok (tr "Instalado (version $($mod.Version))" "Installed (version $($mod.Version))")
        if (Read-YesNo (tr " Comprobar si hay una version mas reciente?" " Check for a newer version?") $false) {
            $scope = if (Test-IsAdmin) { 'AllUsers' } else { 'CurrentUser' }
            try { Install-Module Microsoft365DSC -Force -AllowClobber -Scope $scope; Write-Ok (tr "Actualizado" "Updated") }
            catch { Write-Warn (tr "No se pudo actualizar: $($_.Exception.Message)" "Could not update: $($_.Exception.Message)") }
        }
    } else {
        Write-Warn (tr "No esta instalado." "Not installed.")
        if (-not (Test-IsAdmin)) {
            Write-Warn (tr "No estas como Administrador. Se instalara en el ambito CurrentUser." "You are not running as Administrator. It will be installed in CurrentUser scope.")
            Write-Warn (tr "La guia recomienda -Scope AllUsers (requiere abrir PowerShell como Administrador)." "The guide recommends -Scope AllUsers (requires opening PowerShell as Administrator).")
        }
        if (Read-YesNo (tr " Instalar Microsoft365DSC ahora?" " Install Microsoft365DSC now?") $true) {
            $scope = if (Test-IsAdmin) { 'AllUsers' } else { 'CurrentUser' }
            try {
                Install-Module Microsoft365DSC -Force -AllowClobber -Scope $scope
                Write-Ok (tr "Instalado en el ambito $scope" "Installed in $scope scope")
            } catch { Write-Err (tr "Fallo la instalacion: $($_.Exception.Message)" "Installation failed: $($_.Exception.Message)") }
        }
    }

    # --- Dependencias ---
    if (Get-Module -ListAvailable -Name Microsoft365DSC) {
        Write-Step (tr "Dependencias de Microsoft365DSC" "Microsoft365DSC dependencies")
        Write-Host (tr " Update-M365DSCDependencies alinea las versiones de los submodulos" " Update-M365DSCDependencies aligns the submodule versions") -ForegroundColor Gray
        Write-Host (tr " (Graph, Exchange, PnP, etc.). Puede tardar varios minutos." " (Graph, Exchange, PnP, etc.). It may take several minutes.") -ForegroundColor Gray
        if (Read-YesNo (tr " Ejecutar Update-M365DSCDependencies ahora?" " Run Update-M365DSCDependencies now?") $true) {
            try {
                Import-Module Microsoft365DSC -ErrorAction Stop
                Update-M365DSCDependencies
                Write-Ok (tr "Dependencias actualizadas" "Dependencies updated")
            } catch { Write-Warn (tr "Fallo: $($_.Exception.Message)" "Failed: $($_.Exception.Message)") }
        }
    }

    Write-Host ""
    Write-Host ("=" * 66) -ForegroundColor Green
    Write-Host (tr " ENTORNO PREPARADO" " ENVIRONMENT READY") -ForegroundColor Green
    Write-Host ("=" * 66) -ForegroundColor Green
    Write-Host (tr " Siguiente: genera la consulta de export en" " Next: generate the export query at") -ForegroundColor Yellow
    Write-Host "   https://export.microsoft365dsc.com/" -ForegroundColor Yellow
    Write-Host (tr " y guardala como $((Join-Path $Root 'Scripts\ConfigurationFile.ps1'))" " and save it as $((Join-Path $Root 'Scripts\ConfigurationFile.ps1'))") -ForegroundColor Yellow
    Write-Host ""
}


# ============================================================
#  PASO - PROVISION
# ============================================================
function Invoke-ProvisionStep {
$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  [X]  $m" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $yes = if ($script:Lang -eq 'EN') { 'Y' } else { 'S' }
    $d = if ($Default) { $yes } else { 'N' }
    while ($true) {
        $v = Read-Host "$Prompt ($yes/N) [$d]"
        if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
        switch ($v.Trim().ToUpper()) {
            'S' { return $true }
            'Y' { return $true }
            'N' { return $false }
            default { Write-Warn (tr "Responde S o N" "Answer Y or N") }
        }
    }
}


# ============================================================
#  CATALOGO DE PERMISOS POR WORKLOAD
# ============================================================

$PermissionCatalog = [ordered]@{

    'Base' = @{
        Prefixes = @('')
        Graph    = @(
            'Organization.Read.All'
            'Directory.Read.All'
        )
    }

    'EntraID' = @{
        Prefixes = @('AAD')
        Graph    = @(
            'Application.Read.All'
            'AdministrativeUnit.Read.All'
            'AccessReview.Read.All'
            'Agreement.Read.All'
            'CustomSecAttributeDefinition.Read.All'
            'Device.Read.All'
            'Domain.Read.All'
            'EntitlementManagement.Read.All'
            'Group.Read.All'
            'GroupMember.Read.All'
            'IdentityProvider.Read.All'
            'IdentityRiskyUser.Read.All'
            'IdentityUserFlow.Read.All'
            'LifecycleWorkflows.Read.All'
            'NetworkAccessPolicy.Read.All'
            'OnPremDirectorySynchronization.Read.All'
            'Policy.Read.All'
            'Policy.Read.ConditionalAccess'
            'Policy.Read.PermissionGrant'
            'PrivilegedAccess.Read.AzureAD'
            'PrivilegedAccess.Read.AzureADGroup'
            'RoleEligibilitySchedule.Read.Directory'
            'RoleManagement.Read.All'
            'RoleManagement.Read.Directory'
            'User.Read.All'
            'UserAuthenticationMethod.Read.All'
        )
    }

    'Intune' = @{
        Prefixes = @('Intune')
        Graph    = @(
            'DeviceManagementApps.Read.All'
            'DeviceManagementConfiguration.Read.All'
            'DeviceManagementManagedDevices.Read.All'
            'DeviceManagementRBAC.Read.All'
            'DeviceManagementServiceConfig.Read.All'
            'DeviceManagementScripts.Read.All'          # scripts / remediaciones
            'DeviceManagementConfiguration.ReadWrite.All'
            'Policy.Read.All'
            'Group.Read.All'
            'CloudPC.Read.All'
        )
    }

    'Exchange' = @{
        Prefixes = @('EXO')
        Graph    = @()
        Other    = @{ 'Exchange' = @('Exchange.ManageAsApp') }
    }

    'Purview' = @{
        Prefixes = @('SC')
        Graph    = @(
            'InformationProtectionPolicy.Read.All'
            'SecurityEvents.Read.All'
        )
        Other    = @{ 'Exchange' = @('Exchange.ManageAsApp') }
    }

    'SharePoint' = @{
        Prefixes = @('SPO','ODSettings')
        Graph    = @(
            'Sites.Read.All'
            'Files.Read.All'
        )
        Other    = @{ 'SharePoint' = @('Sites.FullControl.All') }
    }

    'Teams' = @{
        Prefixes = @('Teams')
        Graph    = @(
            'Group.Read.All'
            'Channel.ReadBasic.All'
            'TeamSettings.Read.All'
            'TeamMember.Read.All'
            'TeamsTab.Read.All'
            'TeamsAppInstallation.ReadForTeam.All'
            'User.Read.All'
            'Policy.Read.All'
        )
    }

    'Office365' = @{
        Prefixes = @('O365')
        Graph    = @(
            'Group.Read.All'
            'ExternalConnection.Read.All'
            'SearchConfiguration.Read.All'
            'PeopleSettings.Read.All'                   # Copilot people settings
        )
    }

    'Defender' = @{
        Prefixes = @('Defender')
        Graph    = @(
            'SecurityEvents.Read.All'
            'ThreatHunting.Read.All'
        )
        Other    = @{ 'Exchange' = @('Exchange.ManageAsApp') }
    }

    'Planner' = @{
        Prefixes = @('Planner')
        Graph    = @(
            'Tasks.Read.All'
            'Group.Read.All'
        )
    }

    'PowerPlatform' = @{
        Prefixes = @('PP')
        Graph    = @('Directory.Read.All')
    }

    'Azure' = @{
        Prefixes = @('Azure')
        Graph    = @(
            'RoleManagement.Read.All'
            'Directory.Read.All'
        )
    }
}

$ResourceAppIds = @{
    'Graph'      = '00000003-0000-0000-c000-000000000000'
    'Exchange'   = '00000002-0000-0ff1-ce00-000000000000'
    'SharePoint' = '00000003-0000-0ff1-ce00-000000000000'
}

# Componentes conocidos como problematicos (se ofrece excluirlos)
$KnownProblematic = @{
    'AADVerifiedIdAuthority'                    = 'Verified ID no aprovisionado en la mayoria de tenants'
    'AADVerifiedIdAuthorityContract'            = 'Verified ID no aprovisionado en la mayoria de tenants'
    'AADUserFlowAttribute'                      = 'Recurso exclusivo de Azure AD B2C'
    'IntuneCustomizationBrandingProfile'         = 'Bug: falla si el perfil tiene DisplayName vacio'
    'AzureRoleDefinition'                       = 'Requiere rol RBAC sobre una suscripcion Azure'
    'AzureRoleAssignmentScheduleRequest'        = 'Requiere rol RBAC sobre una suscripcion Azure'
    'AzureRoleEligibilityScheduleRequest'       = 'Requiere rol RBAC sobre una suscripcion Azure'
    'AzureRoleEligibilityScheduleSettings'      = 'Requiere rol RBAC sobre una suscripcion Azure'
    'PlannerBucket'                             = 'Solo soporta autenticacion delegada (Credential)'
    'PlannerPlan'                               = 'Solo soporta autenticacion delegada (Credential)'
    'PlannerTask'                               = 'Solo soporta autenticacion delegada (Credential)'
}


# ============================================================
#  ENTRADA INTERACTIVA
# ============================================================
Clear-Host
Write-Host ("=" * 66) -ForegroundColor Cyan
Write-Host " APP REGISTRATION PARA MICROSOFT365DSC - MODO CERTIFICADO" -ForegroundColor Cyan
Write-Host ("=" * 66) -ForegroundColor Cyan

# --- Componentes ---
Write-Step "Componentes a exportar"
Write-Host "   1) Pegar la lista ahora" -ForegroundColor Gray
Write-Host "   2) Leer desde fichero de texto (uno por linea)" -ForegroundColor Gray
Write-Host "   3) Extraer desde un script de export existente (.ps1)" -ForegroundColor Gray

$modo = Read-WithDefault "Selecciona" "3"
$Components = @()

switch ($modo) {
    '2' {
        do {
            $path = (Read-Host " Ruta del fichero").Trim('"').Trim()
            if (-not (Test-Path $path)) { Write-Err "No existe: $path" }
        } until (Test-Path $path)
        $Components = Get-Content $path |
                      ForEach-Object { $_.Trim().Trim(',').Trim('"').Trim("'") } |
                      Where-Object { $_ -and $_ -notmatch '^#' }
    }
    '3' {
        do {
            $path = (Read-Host " Ruta del script .ps1").Trim('"').Trim()
            if (-not (Test-Path $path)) { Write-Err "No existe: $path" }
        } until (Test-Path $path)
        $raw = Get-Content $path -Raw
        if ($raw -match '(?s)-Components\s*@\((.*?)\)') {
            $Components = [regex]::Matches($Matches[1], '"([^"]+)"') |
                          ForEach-Object { $_.Groups[1].Value }
        } else { Write-Err "No se encontro un bloque -Components @(...)" }
    }
    default {
        Write-Host " Pega los componentes. Escribe FIN para terminar:" -ForegroundColor Gray
        $buffer = @()
        while ($true) {
            $line = Read-Host
            if ($line.Trim().ToUpper() -eq 'FIN') { break }
            $buffer += $line
        }
        $blob = $buffer -join "`n"
        if ($blob -match '"') {
            $Components = [regex]::Matches($blob, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        } else {
            $Components = $blob -split "`n" |
                          ForEach-Object { $_.Trim().Trim(',').Trim("'") } |
                          Where-Object { $_ -and $_ -notmatch '^#' }
        }
    }
}

$Components = @($Components | Select-Object -Unique)
if ($Components.Count -eq 0) { Write-Err "No se obtuvieron componentes. Abortando."; return }
Write-Ok "$($Components.Count) componentes cargados"

# --- Filtrado de problematicos ---
$found = @($Components | Where-Object { $KnownProblematic.ContainsKey($_) })
if ($found.Count -gt 0) {
    Write-Step "Componentes con problemas conocidos"
    foreach ($c in $found) {
        Write-Host ("    {0,-42} {1}" -f $c, $KnownProblematic[$c]) -ForegroundColor Yellow
    }
    if (Read-YesNo "`n Excluirlos del export generado?" $true) {
        $Components = @($Components | Where-Object { $KnownProblematic.Keys -notcontains $_ })
        Write-Ok "Quedan $($Components.Count) componentes"
    }
}

# --- App Registration ---
Write-Step "Datos de la App Registration"
$AppDisplayName = Read-WithDefault " Nombre de la aplicacion" "M365DSC-Export"
$AssignDirRoles = Read-YesNo " Asignar roles de directorio (Global Reader / Exchange Admin)?" $true

# --- Certificado ---
Write-Step "Certificado"
Write-Host "   1) Generar uno nuevo autofirmado" -ForegroundColor Gray
Write-Host "   2) Usar uno existente del almacen (por thumbprint)" -ForegroundColor Gray
$certMode = Read-WithDefault " Selecciona" "1"

$CertSubject = $null; $CertYears = 2; $ExistingThumb = $null
$CertStore = 'Cert:\CurrentUser\My'

if ($certMode -eq '2') {
    do {
        $ExistingThumb = (Read-Host " Thumbprint del certificado").Trim().Replace(' ','').ToUpper()
        $test = Get-ChildItem $CertStore | Where-Object { $_.Thumbprint -eq $ExistingThumb }
        if (-not $test) {
            $test = Get-ChildItem 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Thumbprint -eq $ExistingThumb }
            if ($test) { $CertStore = 'Cert:\LocalMachine\My' }
        }
        if (-not $test) { Write-Err "No encontrado en CurrentUser\My ni LocalMachine\My" }
    } until ($test)
    Write-Ok "Encontrado: $($test.Subject) (expira $($test.NotAfter.ToString('yyyy-MM-dd')))"
} else {
    $CertSubject = Read-WithDefault " Subject del certificado" "CN=$AppDisplayName"
    $CertYears = 0
    while ($CertYears -lt 1 -or $CertYears -gt 5) {
        $v = Read-WithDefault " Validez en anos (1-5)" "2"
        [int]::TryParse($v, [ref]$CertYears) | Out-Null
        if ($CertYears -lt 1 -or $CertYears -gt 5) { Write-Warn "Debe estar entre 1 y 5" }
    }
    $ExportPfx = Read-YesNo " Exportar tambien .pfx (para usar en otra maquina)?" $true
}

# --- Tenant ---
Write-Step "Tenant destino"
Write-Host " Deja vacio para usar el tenant del usuario que inicie sesion." -ForegroundColor Gray
$TenantHint = (Read-Host " TenantId o dominio (opcional)").Trim()

# --- Salida ---
Write-Step "Rutas de salida"
$ExportPath = Read-WithDefault " Carpeta destino del export M365DSC" "C:\M365DSC\Export"
$OutDir     = Read-WithDefault " Carpeta para el script y el certificado" $PWD.Path

if (-not (Test-Path $OutDir)) {
    if (Read-YesNo " La carpeta no existe. Crearla?" $true) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    } else { Write-Err "Abortando."; return }
}

# El export unificado (paso 4) aisla SharePoint en un proceso hijo por si mismo,
# por lo que el script principal incluye TODOS los componentes (SPO incluido).


# ============================================================
#  CALCULO DE PERMISOS
# ============================================================
Write-Step "Analizando componentes"

$activeWorkloads = @(); $graphPerms = @(); $otherPerms = @{}

foreach ($wlName in $PermissionCatalog.Keys) {
    $wl = $PermissionCatalog[$wlName]
    $match = $false
    foreach ($pfx in $wl.Prefixes) {
        if ($pfx -eq '') { $match = $true; break }
        if ($Components | Where-Object { $_ -like "$pfx*" }) { $match = $true; break }
    }
    if (-not $match) { continue }

    $activeWorkloads += $wlName
    $graphPerms      += $wl.Graph
    if ($wl.ContainsKey('Other')) {
        foreach ($resKey in $wl.Other.Keys) {
            if (-not $otherPerms.ContainsKey($resKey)) { $otherPerms[$resKey] = @() }
            $otherPerms[$resKey] += $wl.Other[$resKey]
        }
    }
}

$graphPerms = @($graphPerms | Sort-Object -Unique)
foreach ($k in @($otherPerms.Keys)) { $otherPerms[$k] = @($otherPerms[$k] | Sort-Object -Unique) }

Write-Host "  Workloads detectados:" -ForegroundColor Gray
foreach ($wlName in $activeWorkloads) {
    $count = 0
    foreach ($pfx in $PermissionCatalog[$wlName].Prefixes) {
        if ($pfx -eq '') { $count = $Components.Count; break }
        $count += @($Components | Where-Object { $_ -like "$pfx*" }).Count
    }
    Write-Host ("    {0,-16} {1,4} componentes" -f $wlName, $count) -ForegroundColor Gray
}

Write-Host "`n  Microsoft Graph ($($graphPerms.Count)):" -ForegroundColor Gray
$graphPerms | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }
foreach ($resKey in $otherPerms.Keys) {
    Write-Host "`n  ${resKey}:" -ForegroundColor Gray
    $otherPerms[$resKey] | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }
}

Write-Step "Resumen"
Write-Host "  Componentes    : $($Components.Count)"
Write-Host "  Workloads      : $($activeWorkloads -join ', ')"
Write-Host "  Permisos Graph : $($graphPerms.Count)"
Write-Host "  Aplicacion     : $AppDisplayName"
Write-Host "  Autenticacion  : Certificado"
Write-Host "  Roles dir.     : $(if($AssignDirRoles){'si'}else{'no'})"
Write-Host "  Tenant         : $(if($TenantHint){$TenantHint}else{'(el del usuario)'})"
Write-Host "  Export path    : $ExportPath"
Write-Host "  Salida         : $OutDir"

if (-not (Read-YesNo "`n Continuar?" $true)) { Write-Warn "Cancelado."; return }


# ============================================================
#  MODULOS DE GRAPH
# ============================================================
Write-Step "Preparando modulos de Microsoft Graph"

$graphModules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.Identity.DirectoryManagement'
)

foreach ($m in $graphModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Warn "Instalando $m ..."
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
}

$versionSets = foreach ($m in $graphModules) {
    ,@(Get-Module -ListAvailable -Name $m | Select-Object -ExpandProperty Version)
}
$commonVersions = $versionSets[0]
foreach ($set in $versionSets[1..($versionSets.Count - 1)]) {
    $commonVersions = $commonVersions | Where-Object { $set -contains $_ }
}
$targetVersion = $commonVersions | Sort-Object -Descending | Select-Object -First 1

if (-not $targetVersion) {
    $targetVersion = (Get-Module Microsoft.Graph.Authentication -ListAvailable |
                      Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Warn "Sin version comun. Forzando $targetVersion."
    foreach ($m in $graphModules) {
        if (-not (Get-Module -ListAvailable -Name $m | Where-Object { $_.Version -eq $targetVersion })) {
            Install-Module $m -RequiredVersion $targetVersion -Scope CurrentUser -Force -AllowClobber
        }
    }
}
Write-Ok "Version objetivo: $targetVersion"

Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
foreach ($m in $graphModules) {
    $mod = Get-Module -ListAvailable -Name $m |
           Where-Object { $_.Version -eq $targetVersion } | Select-Object -First 1
    if (-not $mod) { Write-Err "No se encuentra $m $targetVersion. Abortando."; return }
    Import-Module $mod.Path -Force -ErrorAction Stop
    Write-Ok "$m $targetVersion"
}


# ============================================================
#  CERTIFICADO
# ============================================================
Write-Step "Certificado"

if ($certMode -eq '2') {
    $cert = Get-ChildItem $CertStore | Where-Object { $_.Thumbprint -eq $ExistingThumb }
    Write-Ok "Usando existente: $($cert.Thumbprint)"
} else {
    $cert = New-SelfSignedCertificate `
        -Subject           $CertSubject `
        -CertStoreLocation $CertStore `
        -KeyExportPolicy   Exportable `
        -KeySpec           Signature `
        -KeyLength         2048 `
        -KeyAlgorithm      RSA `
        -HashAlgorithm     SHA256 `
        -NotAfter          (Get-Date).AddYears($CertYears)

    Write-Ok "Creado: $($cert.Subject)"
    Write-Ok "Thumbprint: $($cert.Thumbprint)"
    Write-Ok "Expira: $($cert.NotAfter.ToString('yyyy-MM-dd'))"
}

$cerPath = Join-Path $OutDir "$AppDisplayName.cer"
Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
Write-Ok "Clave publica exportada: $cerPath"

$pfxPath = $null
if ($certMode -ne '2' -and $ExportPfx) {
    Write-Host " Introduce una contrasena para proteger el .pfx:" -ForegroundColor Gray
    $pfxPwd = Read-Host " Contrasena" -AsSecureString
    $pfxPath = Join-Path $OutDir "$AppDisplayName.pfx"
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPwd -Force | Out-Null
    Write-Ok "Clave privada exportada: $pfxPath"
    Write-Warn "El .pfx contiene la clave privada. Protegelo como una contrasena."
}


# ============================================================
#  CONEXION
# ============================================================
Write-Step "Conectando a Microsoft Graph (se abrira el navegador)"

$connectArgs = @{
    Scopes = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'RoleManagement.ReadWrite.Directory'
        'Directory.ReadWrite.All'
    )
    NoWelcome = $true
}
if ($TenantHint) { $connectArgs['TenantId'] = $TenantHint }

Connect-MgGraph @connectArgs

$ctx = Get-MgContext
Write-Ok "Tenant : $($ctx.TenantId)"
Write-Ok "Usuario: $($ctx.Account)"

if (-not (Read-YesNo " Es el tenant correcto?" $true)) {
    Disconnect-MgGraph | Out-Null; Write-Warn "Cancelado."; return
}


# ============================================================
#  APP REGISTRATION
# ============================================================
Write-Step "App Registration: $AppDisplayName"

$app = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue |
       Select-Object -First 1

if ($app) {
    Write-Warn "Ya existe (AppId: $($app.AppId))"
    if (-not (Read-YesNo " Reutilizarla y actualizar permisos + certificado?" $true)) {
        Disconnect-MgGraph | Out-Null; Write-Warn "Cancelado."; return
    }
} else {
    $app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience 'AzureADMyOrg'
    Write-Ok "Creada (AppId: $($app.AppId))"
    Start-Sleep -Seconds 10
}

$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue |
      Select-Object -First 1
if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Ok "Service Principal creado"
    Start-Sleep -Seconds 10
} else { Write-Ok "Service Principal existente" }


# ============================================================
#  SUBIR EL CERTIFICADO A LA APP
# ============================================================
Write-Step "Subiendo el certificado a la aplicacion"

$app = Get-MgApplication -ApplicationId $app.Id
$existingKeys = @($app.KeyCredentials)

$alreadyThere = $existingKeys | Where-Object {
    $_.CustomKeyIdentifier -and
    ([System.BitConverter]::ToString($_.CustomKeyIdentifier).Replace('-','')) -eq $cert.Thumbprint
}

if ($alreadyThere) {
    Write-Ok "El certificado ya estaba asociado a la app"
} else {
    $newKey = @{
        Type                = 'AsymmetricX509Cert'
        Usage               = 'Verify'
        Key                 = $cert.GetRawCertData()
        DisplayName         = "CN=$AppDisplayName"
        StartDateTime       = $cert.NotBefore
        EndDateTime         = $cert.NotAfter
    }

    # Conservar los certificados previos que sigan vigentes
    $keep = @($existingKeys | Where-Object { $_.EndDateTime -gt (Get-Date) } | ForEach-Object {
        @{
            Type          = $_.Type
            Usage         = $_.Usage
            Key           = $_.Key
            DisplayName   = $_.DisplayName
            StartDateTime = $_.StartDateTime
            EndDateTime   = $_.EndDateTime
        }
    })

    Update-MgApplication -ApplicationId $app.Id -KeyCredentials (@($keep) + @($newKey))
    Write-Ok "Certificado asociado ($($cert.Thumbprint))"
    if ($keep.Count -gt 0) { Write-Ok "Conservados $($keep.Count) certificados previos vigentes" }
    Start-Sleep -Seconds 5
}


# ============================================================
#  MANIFIESTO
# ============================================================
Write-Step "Registrando permisos en el manifiesto"

function Get-ResourceEntry {
    param([string]$ResourceAppId, [string[]]$PermissionNames)

    $resSp = Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -ErrorAction SilentlyContinue |
             Select-Object -First 1
    if (-not $resSp) { Write-Warn "Service Principal no encontrado: $ResourceAppId"; return $null }

    $access = @()
    foreach ($p in $PermissionNames) {
        $role = $resSp.AppRoles | Where-Object {
            $_.Value -eq $p -and $_.AllowedMemberTypes -contains 'Application' -and $_.IsEnabled
        }
        if ($role) { $access += @{ id = $role.Id; type = 'Role' } }
        else       { Write-Warn "No disponible en este tenant, se omite: $p" }
    }
    if ($access.Count -eq 0) { return $null }

    return @{
        ResourceSp     = $resSp
        ResourceAccess = $access
        Entry          = @{ resourceAppId = $ResourceAppId; resourceAccess = $access }
    }
}

$entries = @()
$g = Get-ResourceEntry -ResourceAppId $ResourceAppIds['Graph'] -PermissionNames $graphPerms
if ($g) { $entries += $g }
foreach ($resKey in $otherPerms.Keys) {
    $r = Get-ResourceEntry -ResourceAppId $ResourceAppIds[$resKey] -PermissionNames $otherPerms[$resKey]
    if ($r) { $entries += $r }
}

if ($entries.Count -eq 0) { Write-Err "No se resolvio ningun permiso. Abortando."; return }

Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @($entries.Entry)
Write-Ok "Manifiesto actualizado ($($entries.Count) recursos)"


# ============================================================
#  ADMIN CONSENT
# ============================================================
Write-Step "Concediendo admin consent"

$granted = 0; $skipped = 0; $failed = 0
$existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All

foreach ($entry in $entries) {
    foreach ($ra in $entry.ResourceAccess) {
        if ($existing | Where-Object { $_.AppRoleId -eq $ra.id -and $_.ResourceId -eq $entry.ResourceSp.Id }) {
            $skipped++; continue
        }
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id -PrincipalId $sp.Id `
                -ResourceId $entry.ResourceSp.Id -AppRoleId $ra.id | Out-Null
            $granted++
        } catch { $failed++; Write-Warn "Fallo: $($_.Exception.Message)" }
    }
}
Write-Ok "$granted concedidos | $skipped ya existentes | $failed fallidos"


# ============================================================
#  ROLES DE DIRECTORIO
# ============================================================
if ($AssignDirRoles) {
    Write-Step "Asignando roles de directorio"

    $roles = @('Global Reader')
    if ($otherPerms.ContainsKey('Exchange'))     { $roles += 'Exchange Administrator' }
    if ($activeWorkloads -contains 'SharePoint') { $roles += 'SharePoint Administrator' }

    foreach ($roleName in $roles) {
        $role = Get-MgDirectoryRole -Filter "displayName eq '$roleName'" -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if (-not $role) {
            $tpl = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $roleName }
            if ($tpl) {
                try { $role = New-MgDirectoryRole -RoleTemplateId $tpl.Id; Start-Sleep -Seconds 5 }
                catch { Write-Warn "No se pudo activar ${roleName}" }
            }
        }
        if (-not $role) { Write-Warn "Rol no encontrado: $roleName"; continue }

        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
        if ($members.Id -contains $sp.Id) { Write-Ok "$roleName (ya asignado)"; continue }

        try {
            New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id `
                -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" }
            Write-Ok $roleName
        } catch { Write-Warn "Fallo al asignar ${roleName}: $($_.Exception.Message)" }
    }
}


# ============================================================
#  SALIDA
# ============================================================
$tenantDomain = (Get-MgOrganization).VerifiedDomains |
                Where-Object { $_.IsInitial } | Select-Object -ExpandProperty Name

Write-Host "`n"
Write-Host ("=" * 66) -ForegroundColor Cyan
Write-Host " DATOS DE CONEXION" -ForegroundColor Cyan
Write-Host ("=" * 66) -ForegroundColor Cyan
Write-Host " ApplicationId        : $($app.AppId)"
Write-Host " TenantId             : $tenantDomain"
Write-Host " CertificateThumbprint: $($cert.Thumbprint)"
Write-Host " Certificado expira   : $($cert.NotAfter.ToString('yyyy-MM-dd'))"
Write-Host ("=" * 66) -ForegroundColor Cyan

# El script principal incluye todos los componentes (SPO incluido); el paso 4
# se encarga de ejecutar SharePoint en un proceso aislado.
$mainComponents = $Components

function New-ExportScript {
    param(
        [string]$FilePath, [string[]]$Comps, [string]$Path,
        [string]$Url, [string]$Header
    )
    $list = ($Comps | ForEach-Object { "    `"$_`"" }) -join ",`n"
    $urlLine = if ($Url) { "    -Url                   `"$Url`" ``" + [Environment]::NewLine } else { "" }

@"
# Generado el $(Get-Date -Format 'yyyy-MM-dd HH:mm')
# App: $AppDisplayName  |  Certificado expira: $($cert.NotAfter.ToString('yyyy-MM-dd'))
$Header

`$AppId      = "$($app.AppId)"
`$TenantId   = "$tenantDomain"
`$Thumbprint = "$($cert.Thumbprint)"

`$Components = @(
$list
)

Export-M365DSCConfiguration ``
    -Components            `$Components ``
    -ApplicationId         `$AppId ``
    -TenantId              `$TenantId ``
    -CertificateThumbprint `$Thumbprint ``
$urlLine    -Path                  "$Path"
"@ | Out-File -FilePath $FilePath -Encoding UTF8
}

$mainFile = Join-Path $OutDir "M365DSC-Export-Main.ps1"
New-ExportScript -FilePath $mainFile -Comps $mainComponents -Path $ExportPath `
    -Header "# Ejecutar en una sesion de PowerShell limpia."
Write-Ok "Script principal: $mainFile  ($($mainComponents.Count) componentes)"

Disconnect-MgGraph | Out-Null

Write-Host "`n" -NoNewline
Write-Host "PROXIMOS PASOS" -ForegroundColor Yellow
Write-Host "  1. Espera 10-15 minutos a que propaguen los permisos." -ForegroundColor Yellow
Write-Host "  2. Ejecuta el paso 4 (Export) apuntando a $((Split-Path $mainFile -Leaf))." -ForegroundColor Yellow
if ($pfxPath) {
    Write-Host "  *  Para otra maquina: importa el .pfx con" -ForegroundColor Yellow
    Write-Host "     Import-PfxCertificate -FilePath '$pfxPath' -CertStoreLocation Cert:\CurrentUser\My" -ForegroundColor Yellow
}
Write-Host ""
}


# ============================================================
#  PASO - EXPORT
# ============================================================
function Invoke-ExportStep {
$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  [X]  $m" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $yes = if ($script:Lang -eq 'EN') { 'Y' } else { 'S' }
    $d = if ($Default) { $yes } else { 'N' }
    while ($true) {
        $v = Read-Host "$Prompt ($yes/N) [$d]"
        if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
        switch ($v.Trim().ToUpper()) {
            'S' { return $true }
            'Y' { return $true }
            'N' { return $false }
            default { Write-Warn (tr "Responde S o N" "Answer Y or N") }
        }
    }
}


# ============================================================
#  CLASIFICACION DE COMPONENTES SPO
# ============================================================

# Recursos que iteran sitio por sitio: coste O(n) sobre el numero de sitios
$SPOPerSiteComponents = @{
    'SPOSite'              = (tr 'Recorre TODOS los sitios del tenant' 'Iterates over ALL tenant sites')
    'SPOSiteAuditSettings' = (tr 'Abre cada sitio para leer su configuracion de auditoria' 'Opens each site to read its audit settings')
    'SPOSiteGroup'         = (tr 'Enumera los grupos de CADA sitio' 'Enumerates the groups of EACH site')
    'SPOPropertyBag'       = (tr 'Lee el property bag de CADA sitio' 'Reads the property bag of EACH site')
    'SPOUserProfileProperty' = (tr 'Recorre el perfil de CADA usuario del tenant' 'Iterates over EACH tenant user profile')
}

# Recursos SPO a nivel de tenant: rapidos, una sola llamada
$SPOTenantComponents = @(
    'SPOAccessControlSettings','SPOApp','SPOBrowserIdleSignout','SPOHomeSite'
    'SPOHubSite','SPOOrgAssetsLibrary','SPORetentionLabelsSettings'
    'SPOSearchManagedProperty','SPOSearchResultSource','SPOSharingSettings'
    'SPOSiteDesign','SPOSiteDesignRights','SPOSiteScript','SPOStorageEntity'
    'SPOTenantCdnEnabled','SPOTenantCdnPolicy','SPOTenantSettings','SPOTheme'
    'ODSettings'
)

function Test-IsSPOComponent {
    param([string]$Name)
    return ($Name -match '^(SPO|ODSettings)')
}


# ============================================================
#  ENTRADA
# ============================================================
Clear-Host
Write-Host ("=" * 68) -ForegroundColor Cyan
Write-Host (tr " EXPORT UNIFICADO MICROSOFT365DSC" " UNIFIED MICROSOFT365DSC EXPORT") -ForegroundColor Cyan
Write-Host ("=" * 68) -ForegroundColor Cyan

# --- Origen de los parametros ---
if (-not $ConfigFile -and -not $AppId) {
    Write-Step (tr "Origen de la configuracion" "Configuration source")
    Write-Host (tr "   1) Leer de un script de export generado (.ps1)" "   1) Read from a generated export script (.ps1)") -ForegroundColor Gray
    Write-Host (tr "   2) Introducir los datos manualmente" "   2) Enter the data manually") -ForegroundColor Gray
    $src = Read-WithDefault (tr " Selecciona" " Select") "1"

    if ($src -eq '1') {
        # Lo normal es el script generado por el paso 3 (Provisionar App):
        $defaultCfg = Join-Path $Root 'Export\M365DSC-Export-Main.ps1'
        do {
            $ConfigFile = (Read-WithDefault (tr " Ruta del script de export" " Path to the export script") $defaultCfg).Trim('"').Trim()
            if (-not (Test-Path $ConfigFile)) { Write-Err (tr "No existe: $ConfigFile" "Does not exist: $ConfigFile") }
        } until (Test-Path $ConfigFile)
    }
}

if ($ConfigFile) {
    Write-Step (tr "Leyendo parametros de $(Split-Path $ConfigFile -Leaf)" "Reading parameters from $(Split-Path $ConfigFile -Leaf)")
    $raw = Get-Content $ConfigFile -Raw

    if (-not $AppId      -and $raw -match '\$AppId\s*=\s*"([^"]+)"')      { $AppId      = $Matches[1] }
    if (-not $TenantId   -and $raw -match '\$TenantId\s*=\s*"([^"]+)"')   { $TenantId   = $Matches[1] }
    if (-not $Thumbprint -and $raw -match '\$Thumbprint\s*=\s*"([^"]+)"') { $Thumbprint = $Matches[1] }
    if (-not $SPOUrl     -and $raw -match '-Url\s+"([^"]+)"')             { $SPOUrl     = $Matches[1] }
    if (-not $OutputPath -and $raw -match '-Path\s+"([^"]+)"')            { $OutputPath = $Matches[1] }

    if (-not $Components -and $raw -match '(?s)\$Components\s*=\s*@\((.*?)\n\)') {
        $Components = [regex]::Matches($Matches[1], '"([^"]+)"') |
                      ForEach-Object { $_.Groups[1].Value }
    }

    Write-Ok "AppId      : $AppId"
    Write-Ok "TenantId   : $TenantId"
    Write-Ok "Thumbprint : $Thumbprint"
    Write-Ok (tr "Componentes: $($Components.Count)" "Components : $($Components.Count)")
}

if (-not $AppId)      { $AppId      = (Read-Host " ApplicationId").Trim() }
if (-not $TenantId)   { $TenantId   = (Read-Host (tr " TenantId (dominio)" " TenantId (domain)")).Trim() }
if (-not $Thumbprint) { $Thumbprint = (Read-Host " CertificateThumbprint").Trim().Replace(' ','').ToUpper() }

if (-not $Components -or $Components.Count -eq 0) {
    Write-Step (tr "Componentes" "Components")
    Write-Host (tr " Pega la lista. Escribe FIN para terminar:" " Paste the list. Type FIN to finish:") -ForegroundColor Gray
    $buffer = @()
    while ($true) {
        $line = Read-Host
        if ($line.Trim().ToUpper() -eq 'FIN') { break }
        $buffer += $line
    }
    $blob = $buffer -join "`n"
    if ($blob -match '"') {
        $Components = [regex]::Matches($blob, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
    } else {
        $Components = $blob -split "`n" |
                      ForEach-Object { $_.Trim().Trim(',').Trim("'") } |
                      Where-Object { $_ -and $_ -notmatch '^#' }
    }
}

$Components = @($Components | Select-Object -Unique)
if ($Components.Count -eq 0) { Write-Err (tr "Sin componentes. Abortando." "No components. Aborting."); return }

if (-not $OutputPath) {
    $OutputPath = Read-WithDefault (tr " Carpeta de salida" " Output folder") "C:\Microsoft365DSC\Export"
}


# ============================================================
#  ANALISIS DE COMPONENTES SPO
# ============================================================
$spoAll   = @($Components | Where-Object { Test-IsSPOComponent $_ })
$mainComp = @($Components | Where-Object { -not (Test-IsSPOComponent $_) })

$spoHeavy = @($spoAll | Where-Object { $SPOPerSiteComponents.ContainsKey($_) })
$spoLight = @($spoAll | Where-Object { -not $SPOPerSiteComponents.ContainsKey($_) })

if ($spoHeavy.Count -gt 0 -and -not $Force) {
    Write-Host ""
    Write-Host ("!" * 68) -ForegroundColor Yellow
    Write-Host (tr " ADVERTENCIA: COMPONENTES DE ALTO COSTE DETECTADOS" " WARNING: HIGH-COST COMPONENTS DETECTED") -ForegroundColor Yellow
    Write-Host ("!" * 68) -ForegroundColor Yellow
    Write-Host ""
    Write-Host (tr " Los siguientes recursos NO consultan una API de tenant: recorren" " The following resources do NOT query a tenant API: they iterate") -ForegroundColor Yellow
    Write-Host (tr " objeto por objeto y abren una conexion por cada uno." " object by object and open a connection for each one.") -ForegroundColor Yellow
    Write-Host ""

    foreach ($h in $spoHeavy) {
        Write-Host ("   {0,-24} {1}" -f $h, $SPOPerSiteComponents[$h]) -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host (tr " Tiempo aproximado (referencia real, varia mucho por tenant):" " Approximate time (real reference, varies a lot per tenant):") -ForegroundColor Gray
    Write-Host (tr "     50 sitios    ->  5-15 minutos" "     50 sites     ->  5-15 minutes") -ForegroundColor Gray
    Write-Host (tr "    200 sitios    -> 20-60 minutos" "    200 sites     -> 20-60 minutes") -ForegroundColor Gray
    Write-Host (tr "   1000 sitios    ->  3-8 horas" "   1000 sites     ->  3-8 hours") -ForegroundColor Gray
    Write-Host (tr "   5000+ sitios   -> puede no terminar en un dia" "   5000+ sites    -> may not finish in a day") -ForegroundColor Gray
    Write-Host ""
    Write-Host (tr " Ademas, SPOUserProfileProperty recorre CADA usuario, no cada sitio." " Also, SPOUserProfileProperty iterates over EACH user, not each site.") -ForegroundColor Gray
    Write-Host ""

    Write-Host (tr " Opciones:" " Options:") -ForegroundColor Cyan
    Write-Host (tr "   1) Excluirlos (recomendado para un primer export)" "   1) Exclude them (recommended for a first export)") -ForegroundColor Gray
    Write-Host (tr "   2) Incluirlos todos" "   2) Include them all") -ForegroundColor Gray
    Write-Host (tr "   3) Elegir uno a uno" "   3) Choose one by one") -ForegroundColor Gray

    $opt = Read-WithDefault (tr " Selecciona" " Select") "1"

    switch ($opt) {
        '2' {
            Write-Warn (tr "Se incluiran los $($spoHeavy.Count) componentes de alto coste." "The $($spoHeavy.Count) high-cost components will be included.")
            if (-not (Read-YesNo (tr " Confirmas? Esto puede tardar horas" " Confirm? This may take hours") $false)) {
                $spoHeavy = @()
                Write-Ok (tr "Excluidos" "Excluded")
            }
        }
        '3' {
            $keep = @()
            foreach ($h in $spoHeavy) {
                Write-Host ""
                Write-Host "   $h" -ForegroundColor White
                Write-Host "   $($SPOPerSiteComponents[$h])" -ForegroundColor DarkGray
                if (Read-YesNo (tr "   Incluir?" "   Include?") $false) { $keep += $h }
            }
            $spoHeavy = $keep
            Write-Ok (tr "Se incluiran $($spoHeavy.Count) de alto coste" "$($spoHeavy.Count) high-cost ones will be included")
        }
        default {
            $spoHeavy = @()
            Write-Ok (tr "Componentes de alto coste excluidos" "High-cost components excluded")
        }
    }
}

$spoComp = @($spoLight + $spoHeavy) | Select-Object -Unique

if ($spoComp.Count -gt 0 -and -not $SPOUrl) {
    $prefix = $TenantId -replace '\.onmicrosoft\.com$',''
    $SPOUrl = Read-WithDefault (tr " URL de admin de SharePoint" " SharePoint admin URL") "https://$prefix-admin.sharepoint.com"
}


# ============================================================
#  RESUMEN
# ============================================================
Write-Step (tr "Plan de ejecucion" "Execution plan")
Write-Host (tr "  Fase 1 (sesion actual)  : $($mainComp.Count) componentes generales" "  Phase 1 (current session): $($mainComp.Count) general components")
if ($spoComp.Count -gt 0) {
    Write-Host (tr "  Fase 2 (proceso hijo)   : $($spoComp.Count) componentes SharePoint" "  Phase 2 (child process)  : $($spoComp.Count) SharePoint components")
    Write-Host (tr "                            $($spoLight.Count) de tenant + $($spoHeavy.Count) por sitio" "                             $($spoLight.Count) tenant-level + $($spoHeavy.Count) per-site")
} else {
    Write-Host (tr "  Fase 2                  : omitida (sin componentes SPO)" "  Phase 2                  : skipped (no SPO components)")
}
Write-Host (tr "  Salida                  : $OutputPath\M365TenantConfig.ps1" "  Output                   : $OutputPath\M365TenantConfig.ps1")

if (-not $Force) {
    if (-not (Read-YesNo (tr "`n Iniciar?" "`n Start?") $true)) { Write-Warn (tr "Cancelado." "Cancelled."); return }
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$workDir = Join-Path $OutputPath "_parts"
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()


# ============================================================
#  FASE 1 - COMPONENTES GENERALES
# ============================================================
$mainOut = Join-Path $workDir "main"
$mainFile = $null

if ($mainComp.Count -gt 0) {
    Write-Step (tr "Fase 1: componentes generales ($($mainComp.Count))" "Phase 1: general components ($($mainComp.Count))")

    $t1 = [System.Diagnostics.Stopwatch]::StartNew()
    New-Item -ItemType Directory -Path $mainOut -Force | Out-Null

    try {
        Import-Module Microsoft365DSC -ErrorAction Stop

        Export-M365DSCConfiguration `
            -Components            $mainComp `
            -ApplicationId         $AppId `
            -TenantId              $TenantId `
            -CertificateThumbprint $Thumbprint `
            -Path                  $mainOut

        $mainFile = Get-ChildItem $mainOut -Filter "*.ps1" -Recurse |
                    Select-Object -First 1 -ExpandProperty FullName

        $t1.Stop()
        if ($mainFile) {
            Write-Ok (tr "Completada en $([math]::Round($t1.Elapsed.TotalMinutes,1)) min" "Completed in $([math]::Round($t1.Elapsed.TotalMinutes,1)) min")
        } else {
            Write-Warn (tr "No se genero fichero en la fase 1" "No file was generated in phase 1")
        }
    } catch {
        $t1.Stop()
        Write-Err (tr "Fallo en la fase 1: $($_.Exception.Message)" "Phase 1 failed: $($_.Exception.Message)")
    }
}


# ============================================================
#  FASE 2 - SHAREPOINT EN PROCESO AISLADO
# ============================================================
$spoOut  = Join-Path $workDir "spo"
$spoFile = $null

if ($spoComp.Count -gt 0) {
    Write-Step (tr "Fase 2: SharePoint en proceso aislado ($($spoComp.Count))" "Phase 2: SharePoint in an isolated process ($($spoComp.Count))")
    Write-Host (tr "  Se lanza un PowerShell hijo sin Graph/Az cargados." "  A child PowerShell is launched without Graph/Az loaded.") -ForegroundColor Gray

    if ($spoHeavy.Count -gt 0) {
        Write-Warn (tr "Incluye componentes por sitio. Puede tardar mucho." "Includes per-site components. It may take a long time.")
    }

    $t2 = [System.Diagnostics.Stopwatch]::StartNew()
    New-Item -ItemType Directory -Path $spoOut -Force | Out-Null

    $childScript = Join-Path $workDir "spo-child.ps1"
    $compLiteral = ($spoComp | ForEach-Object { "    `"$_`"" }) -join ",`n"

@"
`$ErrorActionPreference = 'Continue'
`$ProgressPreference    = 'SilentlyContinue'

try {
    Import-Module Microsoft365DSC -ErrorAction Stop
} catch {
    Write-Error "$(tr 'No se pudo cargar Microsoft365DSC' 'Could not load Microsoft365DSC'): `$(`$_.Exception.Message)"
    exit 1
}

`$Components = @(
$compLiteral
)

try {
    Export-M365DSCConfiguration ``
        -Components            `$Components ``
        -ApplicationId         "$AppId" ``
        -TenantId              "$TenantId" ``
        -CertificateThumbprint "$Thumbprint" ``
        -Path                  "$spoOut"
    exit 0
} catch {
    Write-Error "$(tr 'Export SPO fallo' 'SPO export failed'): `$(`$_.Exception.Message)"
    exit 2
}
"@ | Out-File -FilePath $childScript -Encoding UTF8

    $psExe = if ($PSVersionTable.PSVersion.Major -ge 6) { 'pwsh' } else { 'powershell' }

    $proc = Start-Process -FilePath $psExe `
        -ArgumentList @('-NoProfile','-NoLogo','-ExecutionPolicy','Bypass','-File',"`"$childScript`"") `
        -NoNewWindow -Wait -PassThru

    $t2.Stop()

    if ($proc.ExitCode -eq 0) {
        $spoFile = Get-ChildItem $spoOut -Filter "*.ps1" -Recurse -ErrorAction SilentlyContinue |
                   Select-Object -First 1 -ExpandProperty FullName
        if ($spoFile) {
            Write-Ok (tr "Completada en $([math]::Round($t2.Elapsed.TotalMinutes,1)) min" "Completed in $([math]::Round($t2.Elapsed.TotalMinutes,1)) min")
        } else {
            Write-Warn (tr "El proceso termino bien pero no genero fichero" "The process finished fine but generated no file")
        }
    } else {
        Write-Warn (tr "El proceso hijo termino con codigo $($proc.ExitCode)" "The child process exited with code $($proc.ExitCode)")
        $spoFile = Get-ChildItem $spoOut -Filter "*.ps1" -Recurse -ErrorAction SilentlyContinue |
                   Select-Object -First 1 -ExpandProperty FullName
        if ($spoFile) { Write-Warn (tr "Se encontro un fichero parcial, se fusionara igualmente" "A partial file was found; it will be merged anyway") }
    }
}


# ============================================================
#  FUSION
# ============================================================
Write-Step (tr "Fusionando resultados" "Merging results")

function Get-ConfigBlocks {
    <#  Extrae el cuerpo de bloques de recurso de un M365TenantConfig.ps1,
        descartando cabecera, Configuration/Node y el bloque de invocacion. #>
    param([string]$Path)

    if (-not $Path -or -not (Test-Path $Path)) { return @() }
    $lines = Get-Content $Path

    $blocks = @()
    $current = $null
    $depth = 0
    $inNode = $false

    foreach ($line in $lines) {

        if (-not $inNode) {
            if ($line -match '^\s*Node\s') { $inNode = $true }
            continue
        }

        # Inicio de un bloque de recurso:  ResourceName "InstanceName"
        if (-not $current -and $line -match '^\s{0,12}([A-Za-z][A-Za-z0-9]*)\s+"([^"]+)"\s*$') {
            $current = [System.Collections.Generic.List[string]]::new()
            $current.Add($line)
            $depth = 0
            continue
        }

        if ($current) {
            $current.Add($line)
            $depth += ([regex]::Matches($line, '\{')).Count
            $depth -= ([regex]::Matches($line, '\}')).Count
            if ($depth -le 0 -and $line -match '\}') {
                $blocks += ,(($current -join "`n"))
                $current = $null
            }
        }
    }

    return $blocks
}

$mainBlocks = @(Get-ConfigBlocks -Path $mainFile)
$spoBlocks  = @(Get-ConfigBlocks -Path $spoFile)

Write-Ok (tr "Bloques fase 1: $($mainBlocks.Count)" "Phase 1 blocks: $($mainBlocks.Count)")
Write-Ok (tr "Bloques fase 2: $($spoBlocks.Count)" "Phase 2 blocks: $($spoBlocks.Count)")

if ($mainBlocks.Count -eq 0 -and $spoBlocks.Count -eq 0) {
    Write-Err (tr "No se extrajo ningun bloque. Revisa los ficheros en $workDir" "No block was extracted. Check the files in $workDir")
    return
}

# Cabecera: reutilizar la de la fase 1 si existe, si no la de SPO
$headerSource = if ($mainFile) { $mainFile } else { $spoFile }
$headerLines = @()
if ($headerSource) {
    foreach ($line in (Get-Content $headerSource)) {
        if ($line -match '^\s*Node\s') { break }
        $headerLines += $line
    }
}

$allBlocks = $mainBlocks + $spoBlocks

$sb = [System.Text.StringBuilder]::new()
foreach ($h in $headerLines) { [void]$sb.AppendLine($h) }
[void]$sb.AppendLine('    Node localhost')
[void]$sb.AppendLine('    {')
foreach ($b in $allBlocks) {
    [void]$sb.AppendLine($b)
    [void]$sb.AppendLine('')
}
[void]$sb.AppendLine('    }')
[void]$sb.AppendLine('}')
[void]$sb.AppendLine('')
[void]$sb.AppendLine('M365TenantConfig -ConfigurationData $ConfigurationData')

$finalFile = Join-Path $OutputPath "M365TenantConfig.ps1"
$sb.ToString() | Out-File -FilePath $finalFile -Encoding UTF8

$stopwatch.Stop()

# Copiar ficheros auxiliares (psd1 de ConfigurationData) si los hubiera
foreach ($src in @($mainOut, $spoOut)) {
    if (Test-Path $src) {
        Get-ChildItem $src -Filter "*.psd1" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $dest = Join-Path $OutputPath $_.Name
            if (-not (Test-Path $dest)) { Copy-Item $_.FullName $dest -Force }
        }
    }
}


# ============================================================
#  RESUMEN FINAL
# ============================================================
$sizeKb = [math]::Round((Get-Item $finalFile).Length / 1KB, 1)

Write-Host ""
Write-Host ("=" * 68) -ForegroundColor Green
Write-Host (tr " EXPORT COMPLETADO" " EXPORT COMPLETED") -ForegroundColor Green
Write-Host ("=" * 68) -ForegroundColor Green
Write-Host (tr " Fichero    : $finalFile" " File       : $finalFile")
Write-Host (tr " Tamano     : $sizeKb KB" " Size       : $sizeKb KB")
Write-Host (tr " Recursos   : $($allBlocks.Count) bloques" " Resources  : $($allBlocks.Count) blocks")
Write-Host (tr " Duracion   : $([math]::Round($stopwatch.Elapsed.TotalMinutes,1)) minutos" " Duration   : $([math]::Round($stopwatch.Elapsed.TotalMinutes,1)) minutes")
Write-Host ("=" * 68) -ForegroundColor Green

if (Read-YesNo (tr "`n Conservar los ficheros intermedios en _parts?" "`n Keep the intermediate files in _parts?") $false) {
    Write-Ok (tr "Conservados en $workDir" "Kept in $workDir")
} else {
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Ok (tr "Intermedios eliminados" "Intermediate files removed")
}
}


# ============================================================
#  PASO - REPORT
# ============================================================
function Invoke-ReportStep {
$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  [X]  $m" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}


# ============================================================
#  CONFIGURACION DE NORMALIZACION
# ============================================================

$IgnoredProperties = @(
    # Conexion / autenticacion
    'ApplicationId','ApplicationSecret','TenantId','CertificateThumbprint'
    'CertificatePath','CertificatePassword','Credential','ManagedIdentity'
    'AccessTokens'
    # Metadatos DSC
    'Ensure','ResourceInstanceName','ResourceName'
    # Identificadores y timestamps
    'Id','ObjectId','CreatedDateTime','LastModifiedDateTime'
    'CreatedDate','WhenChanged','WhenCreated','Guid'
    # Especificas de comparacion entre tenants
    'ExchangeObjectId','ExternalDirectoryObjectId','OrganizationId'
    'DistinguishedName','LegacyExchangeDN','ExchangeGuid','ArchiveGuid'
    'ImmutableId','OnPremisesSecurityIdentifier','SiteId','WebId'
    'AppId','ServicePrincipalId','ClientId'
    'CreatedBy','ModifiedBy','Owner','LastModifiedBy'
    'Version','ObjectVersion','ChangeKey'
)

$KeyCandidates = @(
    'Identity','DisplayName','Name','Title','Url','UserPrincipalName'
    'EmailAddress','PolicyName','RoleName','GroupName','Domain','Key'
)

$WorkloadMap = [ordered]@{
    'AAD'        = 'Entra ID'
    'Azure'      = 'Azure'
    'Defender'   = 'Defender'
    'EXO'        = 'Exchange Online'
    'Intune'     = 'Intune'
    'O365'       = 'Office 365'
    'ODSettings' = 'OneDrive'
    'Planner'    = 'Planner'
    'PP'         = 'Power Platform'
    'SC'         = 'Purview'
    'SPO'        = 'SharePoint'
    'Teams'      = 'Teams'
}

function Get-Workload {
    param([string]$ResourceName)
    foreach ($pfx in $WorkloadMap.Keys) {
        if ($ResourceName -like "$pfx*") { return $WorkloadMap[$pfx] }
    }
    return 'Otros'
}


# ============================================================
#  NORMALIZACION ENTRE TENANTS
# ============================================================

$script:TenantDomains = @{}

function Get-TenantDomainsFromConfig {
    param([array]$Objects)

    $domains = @{}
    foreach ($obj in $Objects) {
        foreach ($k in $obj.Keys) {
            $v = "$($obj[$k])"
            if ([string]::IsNullOrWhiteSpace($v)) { continue }

            foreach ($m in [regex]::Matches($v, '(?i)[a-z0-9-]+\.onmicrosoft\.com')) {
                $d = $m.Value.ToLower()
                if (-not $domains.ContainsKey($d)) { $domains[$d] = 0 }
                $domains[$d]++
            }
            foreach ($m in [regex]::Matches($v, '(?i)@([a-z0-9.-]+\.[a-z]{2,})')) {
                $d = $m.Groups[1].Value.ToLower()
                if (-not $domains.ContainsKey($d)) { $domains[$d] = 0 }
                $domains[$d]++
            }
        }
    }

    return @($domains.GetEnumerator() |
             Sort-Object Value -Descending |
             Select-Object -First 8 -ExpandProperty Key)
}

function Remove-TenantSpecifics {
    param([string]$Text, [string[]]$Domains, [string]$TenantPrefix)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    $out = $Text

    foreach ($d in $Domains) {
        $out = [regex]::Replace($out, [regex]::Escape($d), '{TENANT-DOMAIN}', 'IgnoreCase')
    }
    if ($TenantPrefix) {
        $out = [regex]::Replace($out, [regex]::Escape($TenantPrefix), '{TENANT}', 'IgnoreCase')
    }
    $out = [regex]::Replace($out,
        '(?i)\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b', '{GUID}')

    return $out
}

function Normalize-Value {
    param($Value, [string[]]$Domains, [string]$TenantPrefix)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Collections.IDictionary]) {
        $sorted = [ordered]@{}
        foreach ($k in ($Value.Keys | Sort-Object)) {
            $sorted[$k] = Normalize-Value $Value[$k] $Domains $TenantPrefix
        }
        return $sorted
    }

    if ($Value -is [array] -or ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string])) {
        $items = @($Value | ForEach-Object { Normalize-Value $_ $Domains $TenantPrefix })
        if (@($items | Where-Object { $_ -is [System.Collections.IDictionary] }).Count -eq 0) {
            $items = @($items | Sort-Object { "$_" })
        }
        return $items
    }

    if ($Value -is [bool]) { return $Value }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return $Value }

    return (Remove-TenantSpecifics -Text "$Value" -Domains $Domains -TenantPrefix $TenantPrefix)
}

function Get-InstanceKey {
    param([hashtable]$Res, [string[]]$Domains, [string]$TenantPrefix)

    foreach ($cand in $KeyCandidates) {
        if ($Res.ContainsKey($cand) -and -not [string]::IsNullOrWhiteSpace("$($Res[$cand])")) {
            return (Remove-TenantSpecifics -Text "$($Res[$cand])" -Domains $Domains -TenantPrefix $TenantPrefix)
        }
    }
    if ($Res.ContainsKey('ResourceInstanceName')) { return "$($Res['ResourceInstanceName'])" }
    return '(singleton)'
}

function Test-ValueEqual {
    param($A, $B)
    if ($null -eq $A -and $null -eq $B) { return $true }
    if ($null -eq $A -or  $null -eq $B) { return $false }
    return (($A | ConvertTo-Json -Depth 12 -Compress) -eq ($B | ConvertTo-Json -Depth 12 -Compress))
}


# ============================================================
#  ENTRADA INTERACTIVA
# ============================================================
if (-not $ConfigPaths -or $ConfigPaths.Count -eq 0) {
    Clear-Host
    Write-Host ("=" * 68) -ForegroundColor Cyan
    Write-Host (tr " REPORTE DE BASELINE - MICROSOFT365DSC" " BASELINE REPORT - MICROSOFT365DSC") -ForegroundColor Cyan
    Write-Host ("=" * 68) -ForegroundColor Cyan

    Write-Step (tr "Configuraciones a comparar" "Configurations to compare")

    $ConfigPaths = @(); $Labels = @()

    # Auto-descubrir los M365TenantConfig.ps1 que dejaste bajo Tenants\
    $tenantsDir = Join-Path $Root 'Tenants'
    $detected = @()
    if (Test-Path $tenantsDir) {
        $detected = @(Get-ChildItem $tenantsDir -Recurse -Filter 'M365TenantConfig.ps1' -ErrorAction SilentlyContinue |
                      Sort-Object { $_.Directory.Name })
    }

    if ($detected.Count -ge 1) {
        Write-Host (tr " Configuraciones detectadas en $tenantsDir :" " Configurations detected in $tenantsDir :") -ForegroundColor Gray
        for ($k = 0; $k -lt $detected.Count; $k++) {
            $kb = [math]::Round($detected[$k].Length / 1KB, 1)
            Write-Host ("   [{0}] {1,-16} {2,8} KB   {3}" -f ($k + 1), $detected[$k].Directory.Name, $kb, $detected[$k].FullName) -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host (tr " Enter = todas  |  numeros separados por coma (ej. 1,3)  |  M = escribir rutas a mano" " Enter = all  |  numbers separated by comma (e.g. 1,3)  |  M = type paths manually") -ForegroundColor DarkGray
        $sel = (Read-Host (tr " Seleccion" " Selection")).Trim()

        if ($sel.ToUpper() -ne 'M') {
            $picks = @()
            if ([string]::IsNullOrWhiteSpace($sel)) {
                $picks = 0..($detected.Count - 1)
            } else {
                foreach ($tok in ($sel -split '[,\s]+')) {
                    $n = 0
                    if ([int]::TryParse($tok, [ref]$n) -and $n -ge 1 -and $n -le $detected.Count) { $picks += ($n - 1) }
                }
                $picks = @($picks | Select-Object -Unique)
            }
            foreach ($idx in $picks) {
                $ConfigPaths += $detected[$idx].FullName
                $Labels      += $detected[$idx].Directory.Name
            }
            if ($ConfigPaths.Count -ge 1) {
                Write-Ok (tr "$($ConfigPaths.Count) seleccionadas: $($Labels -join ', ')" "$($ConfigPaths.Count) selected: $($Labels -join ', ')")
            }
        }
    }

    # Modo manual: si no se detecto nada, elegiste M, o faltan configuraciones
    if ($ConfigPaths.Count -lt 2) {
        if ($detected.Count -eq 0) { Write-Warn (tr "No se detectaron configuraciones en $tenantsDir" "No configurations detected in $tenantsDir") }
        Write-Host (tr " Introduce la ruta de cada M365TenantConfig.ps1 (Enter vacio para terminar, minimo 2)." " Enter the path of each M365TenantConfig.ps1 (empty Enter to finish, minimum 2).") -ForegroundColor Gray
        $i = $ConfigPaths.Count + 1
        while ($true) {
            $p = (Read-Host " Config #$i").Trim('"').Trim()
            if ([string]::IsNullOrWhiteSpace($p)) {
                if ($ConfigPaths.Count -ge 2) { break }
                Write-Warn (tr "Se necesitan al menos 2" "At least 2 are needed")
                continue
            }
            if (-not (Test-Path $p)) { Write-Err (tr "No existe: $p" "Does not exist: $p"); continue }

            $defaultLabel = (Get-Item $p).Directory.Name
            $l = Read-WithDefault (tr "   Etiqueta para esta columna" "   Label for this column") $defaultLabel

            $ConfigPaths += $p
            $Labels      += $l
            $i++
        }
    }

    Write-Step (tr "Configuracion de referencia (baseline)" "Reference configuration (baseline)")
    Write-Host (tr " Las demas se compararan contra esta." " The others will be compared against this one.") -ForegroundColor Gray
    for ($k = 0; $k -lt $Labels.Count; $k++) {
        Write-Host ("   [{0}] {1}" -f ($k+1), $Labels[$k]) -ForegroundColor Gray
    }
    # Por defecto se sugiere la carpeta "Baseline" como baseline, si existe
    $defaultBase = 1
    for ($k = 0; $k -lt $Labels.Count; $k++) {
        if ($Labels[$k] -match '^(?i)baseline$') { $defaultBase = $k + 1; break }
    }
    $BaselineIndex = -1
    while ($BaselineIndex -lt 0 -or $BaselineIndex -ge $Labels.Count) {
        $v = Read-WithDefault (tr " Cual es la baseline" " Which one is the baseline") "$defaultBase"
        $n = 0; [int]::TryParse($v, [ref]$n) | Out-Null
        $BaselineIndex = $n - 1
    }
    Write-Ok "Baseline: $($Labels[$BaselineIndex])"

    Write-Step (tr "Salida" "Output")
    $defaultOut = if (Test-Path (Join-Path $Root 'Reports')) {
                      Join-Path $Root ("Reports\baseline-{0}.html" -f (Get-Date -Format 'yyyy-MM'))
                  } else {
                      Join-Path $PWD.Path "M365DSC-Baseline.html"
                  }
    $OutputPath  = Read-WithDefault (tr " Ruta del HTML" " HTML path") $defaultOut
    $ReportTitle = Read-WithDefault (tr " Titulo del reporte" " Report title") (tr "Comparacion de baseline Microsoft 365" "Microsoft 365 baseline comparison")
    $ClientName  = Read-WithDefault (tr " Cliente / organizacion (opcional)" " Client / organization (optional)") ""

    Write-Step (tr "Marca del reporte (opcional)" "Report branding (optional)")
    Write-Host (tr " Deja en blanco cualquier campo para omitirlo." " Leave any field blank to omit it.") -ForegroundColor Gray
    $BrandName = Read-WithDefault (tr " Nombre / organizacion (cabecera y pie)" " Name / organization (header and footer)") ""
    $Tagline   = Read-WithDefault (tr " Eslogan" " Tagline") ""

    # Se sugiere un logo.* que este junto al script, si existe
    $scriptDirBrand = if ($PSCommandPath) { Split-Path $PSCommandPath -Parent } else { $PWD.Path }
    $logoDefault = ''
    foreach ($cand in @('logo.svg','logo.png','logo.jpg','logo.jpeg','logo.gif','logo.webp')) {
        $t = Join-Path $scriptDirBrand $cand
        if (Test-Path $t) { $logoDefault = $t; break }
    }
    $LogoPath = (Read-WithDefault (tr " Ruta del logo (SVG/PNG/JPG, vacio = sin logo)" " Logo path (SVG/PNG/JPG, empty = no logo)") $logoDefault).Trim('"').Trim()
}

if ([string]::IsNullOrWhiteSpace($ReportTitle)) { $ReportTitle = tr "Comparacion de baseline Microsoft 365" "Microsoft 365 baseline comparison" }

if (-not $Labels -or $Labels.Count -ne $ConfigPaths.Count) {
    $Labels = 1..$ConfigPaths.Count | ForEach-Object { "Config $_" }
}
if ($BaselineIndex -lt 0 -or $BaselineIndex -ge $ConfigPaths.Count) { $BaselineIndex = 0 }
if (-not $OutputPath) { $OutputPath = Join-Path $PWD.Path "M365DSC-Baseline.html" }


# ============================================================
#  LOGO Y MARCA  (se preguntan por pantalla en el modo interactivo)
# ============================================================
if ($null -eq $BrandName) { $BrandName = '' }
if ($null -eq $Tagline)   { $Tagline   = '' }

$logoTag = ''
if ($LogoPath -and (Test-Path $LogoPath)) {
    try {
        $mime = switch ([IO.Path]::GetExtension($LogoPath).ToLowerInvariant()) {
            '.svg'  { 'image/svg+xml' }
            '.png'  { 'image/png' }
            '.jpg'  { 'image/jpeg' }
            '.jpeg' { 'image/jpeg' }
            '.gif'  { 'image/gif' }
            '.webp' { 'image/webp' }
            default { 'image/svg+xml' }
        }
        $logoB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($LogoPath))
        $altTxt  = if ($BrandName) { $BrandName } else { 'Logo' }
        $logoTag = "<img class=`"logo`" src=`"data:$mime;base64,$logoB64`" alt=`"$altTxt`">"
        Write-Ok (tr "Logo embebido desde $LogoPath" "Logo embedded from $LogoPath")
    } catch {
        Write-Warn (tr "No se pudo leer el logo: $($_.Exception.Message)" "Could not read the logo: $($_.Exception.Message)")
    }
} elseif ($LogoPath) {
    Write-Warn (tr "Logo no encontrado en: $LogoPath (el reporte saldra sin logo)" "Logo not found at: $LogoPath (the report will have no logo)")
} else {
    Write-Ok (tr "Reporte sin logo" "Report without logo")
}


# ============================================================
#  CATALOGO DE RECURSOS (descripciones + enlaces a documentacion)
# ============================================================
if (-not $CatalogPath) {
    $scriptDir2 = if ($PSCommandPath) { Split-Path $PSCommandPath -Parent } else { $PWD.Path }
    # Se busca en varias ubicaciones para que el catalogo se encuentre aunque el
    # script se ejecute desde otra carpeta (junto al script, en Scripts\, en $Root).
    $catCandidates = @(
        (Join-Path $scriptDir2 'catalog.json')
        (Join-Path $scriptDir2 'Scripts\catalog.json')
        (Join-Path $Root       'Scripts\catalog.json')
        (Join-Path $PWD.Path   'catalog.json')
        (Join-Path $PWD.Path   'Scripts\catalog.json')
    )
    $CatalogPath = $catCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $CatalogPath) { $CatalogPath = $catCandidates[0] }
}

$script:Catalog = @{}
if (Test-Path $CatalogPath) {
    try {
        $rawCat = Get-Content $CatalogPath -Raw -Encoding UTF8
        $catObj = $rawCat | ConvertFrom-Json
        foreach ($prop in $catObj.PSObject.Properties) {
            if ($prop.Name -eq '_meta') { continue }
            $script:Catalog[$prop.Name] = $prop.Value
        }
        Write-Ok (tr "Catalogo cargado: $($script:Catalog.Count) recursos documentados" "Catalog loaded: $($script:Catalog.Count) documented resources")
    } catch {
        Write-Warn (tr "No se pudo leer el catalogo: $($_.Exception.Message)" "Could not read the catalog: $($_.Exception.Message)")
    }
} else {
    Write-Warn (tr "Catalogo no encontrado en $CatalogPath" "Catalog not found at $CatalogPath")
    Write-Warn (tr "El reporte saldra sin descripciones ni enlaces. Usa -CatalogPath si esta en otra ruta." "The report will have no descriptions or links. Use -CatalogPath if it is elsewhere.")
}

function Get-CatalogEntry {
    param([string]$ResourceName)
    if ($script:Catalog.ContainsKey($ResourceName)) { return $script:Catalog[$ResourceName] }
    return $null
}


# ============================================================
#  CARGA DE MODULO
# ============================================================
Write-Step (tr "Cargando Microsoft365DSC" "Loading Microsoft365DSC")
try {
    Import-Module Microsoft365DSC -ErrorAction Stop
    Write-Ok (tr "Modulo cargado" "Module loaded")
} catch {
    Write-Err (tr "No se pudo cargar Microsoft365DSC: $($_.Exception.Message)" "Could not load Microsoft365DSC: $($_.Exception.Message)")
    Write-Err (tr "Ejecuta este script en una sesion limpia de PowerShell." "Run this script in a clean PowerShell session.")
    return
}


# ============================================================
#  PARSEO Y NORMALIZACION
# ============================================================
Write-Step (tr "Parseando configuraciones" "Parsing configurations")

$allConfigs = @()

for ($c = 0; $c -lt $ConfigPaths.Count; $c++) {
    $path  = $ConfigPaths[$c]
    $label = $Labels[$c]

    Write-Host "  [$($c+1)/$($ConfigPaths.Count)] $label ..." -ForegroundColor Gray -NoNewline

    try {
        $objects = ConvertTo-DSCObject -Path $path -ErrorAction Stop
    } catch {
        Write-Host ""
        Write-Err (tr "Fallo al parsear '$path': $($_.Exception.Message)" "Failed to parse '$path': $($_.Exception.Message)")
        return
    }

    $domains = Get-TenantDomainsFromConfig -Objects $objects
    $prefix  = $null
    $onmic   = $domains | Where-Object { $_ -like '*.onmicrosoft.com' } | Select-Object -First 1
    if ($onmic) { $prefix = $onmic -replace '\.onmicrosoft\.com$','' }

    $script:TenantDomains[$label] = @{ Domains = @($domains); Prefix = $prefix }

    $bucket = @{}
    foreach ($obj in $objects) {
        $resName = "$($obj.ResourceName)"
        if ([string]::IsNullOrWhiteSpace($resName)) { continue }

        $instKey = Get-InstanceKey -Res $obj -Domains $domains -TenantPrefix $prefix
        $fullKey = "$resName||$instKey"

        $props = [ordered]@{}
        foreach ($k in ($obj.Keys | Sort-Object)) {
            if ($IgnoredProperties -contains $k) { continue }
            $props[$k] = Normalize-Value $obj[$k] $domains $prefix
        }

        $suffix = 1; $tryKey = $fullKey
        while ($bucket.ContainsKey($tryKey)) { $suffix++; $tryKey = "$fullKey #$suffix" }

        $bucket[$tryKey] = @{
            ResourceName = $resName
            InstanceKey  = $instKey
            Workload     = Get-Workload -ResourceName $resName
            Properties   = $props
        }
    }

    $allConfigs += @{ Label = $label; Path = $path; Items = $bucket }
    Write-Host " $($bucket.Count) $(tr 'instancias' 'instances')" -ForegroundColor Green
    if ($prefix) { Write-Host "        $(tr 'tenant detectado' 'tenant detected'): $prefix" -ForegroundColor DarkGray }
}


# ============================================================
#  COMPARACION
# ============================================================
Write-Step (tr "Comparando contra la baseline" "Comparing against the baseline")

$allKeys = @()
foreach ($cfg in $allConfigs) { $allKeys += $cfg.Items.Keys }
$allKeys = @($allKeys | Select-Object -Unique | Sort-Object)

Write-Ok (tr "$($allKeys.Count) instancias unicas en total" "$($allKeys.Count) unique instances in total")

$report = @{
    Title     = $ReportTitle
    Client    = $ClientName
    Generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Labels    = $Labels
    Sources   = $ConfigPaths
    BaseIndex = $BaselineIndex
    Tenants   = $script:TenantDomains
    Workloads = @{}
    Stats     = @{ Total=0; Same=0; Diff=0; OnlyBase=0; MissingBase=0; Partial=0 }
}

foreach ($key in $allKeys) {

    $resName = $null; $instKey = $null; $workload = $null
    $presence = @()

    foreach ($cfg in $allConfigs) {
        if ($cfg.Items.ContainsKey($key)) {
            $it = $cfg.Items[$key]
            if (-not $resName) {
                $resName  = $it.ResourceName
                $instKey  = $it.InstanceKey
                $workload = $it.Workload
            }
            $presence += $true
        } else {
            $presence += $false
        }
    }

    $inBase       = $presence[$BaselineIndex]
    $presentCount = @($presence | Where-Object { $_ }).Count
    $allPresent   = ($presentCount -eq $allConfigs.Count)

    $propNames = @()
    foreach ($cfg in $allConfigs) {
        if ($cfg.Items.ContainsKey($key)) { $propNames += $cfg.Items[$key].Properties.Keys }
    }
    $propNames = @($propNames | Select-Object -Unique | Sort-Object)

    $rows = @(); $instHasDiff = $false

    foreach ($pn in $propNames) {
        $values = @()
        foreach ($cfg in $allConfigs) {
            if ($cfg.Items.ContainsKey($key) -and $cfg.Items[$key].Properties.Contains($pn)) {
                $values += ,$cfg.Items[$key].Properties[$pn]
            } else {
                $values += ,$null
            }
        }

        $same = $true
        $baseVal = $values[$BaselineIndex]
        for ($i = 0; $i -lt $values.Count; $i++) {
            if ($i -eq $BaselineIndex) { continue }
            if (-not $presence[$i])    { continue }
            if (-not (Test-ValueEqual $baseVal $values[$i])) { $same = $false; break }
        }
        if (-not $same) { $instHasDiff = $true }

        $display = @($values | ForEach-Object {
            if ($null -eq $_) { $null }
            elseif ($_ -is [System.Collections.IDictionary] -or $_ -is [array]) {
                ($_ | ConvertTo-Json -Depth 8)
            } else { "$_" }
        })

        $rows += @{ n = $pn; v = $display; d = (-not $same) }
    }

    $status =
        if     (-not $inBase)        { 'missingbase' }
        elseif ($presentCount -eq 1) { 'onlybase' }
        elseif (-not $allPresent)    { 'partial' }
        elseif ($instHasDiff)        { 'diff' }
        else                         { 'same' }

    $report.Stats.Total++
    switch ($status) {
        'same'        { $report.Stats.Same++ }
        'diff'        { $report.Stats.Diff++ }
        'onlybase'    { $report.Stats.OnlyBase++ }
        'missingbase' { $report.Stats.MissingBase++ }
        'partial'     { $report.Stats.Partial++ }
    }

    if (-not $report.Workloads.ContainsKey($workload)) {
        $report.Workloads[$workload] = @{ Name = $workload; Resources = @{} }
    }
    if (-not $report.Workloads[$workload].Resources.ContainsKey($resName)) {
        $report.Workloads[$workload].Resources[$resName] = @{ Name = $resName; Instances = @() }
    }

    $report.Workloads[$workload].Resources[$resName].Instances += @{
        key      = $instKey
        status   = $status
        presence = $presence
        rows     = $rows
    }
}

Write-Ok ((tr "Iguales: {0} | Difieren: {1} | Parciales: {2} | Solo baseline: {3} | Faltan en baseline: {4}" "Identical: {0} | Different: {1} | Partial: {2} | Baseline only: {3} | Missing in baseline: {4}") -f `
    $report.Stats.Same, $report.Stats.Diff, $report.Stats.Partial,
    $report.Stats.OnlyBase, $report.Stats.MissingBase)


# ============================================================
#  SERIALIZACION
# ============================================================
Write-Step (tr "Generando HTML" "Generating HTML")

$wlArray = @()
$propDocs = @{}   # "Recurso.Propiedad" -> { d = descripcion; v = valores validos }
$docHits  = 0

foreach ($wlName in ($report.Workloads.Keys | Sort-Object)) {
    $wl = $report.Workloads[$wlName]
    $resArray = @()

    foreach ($rName in ($wl.Resources.Keys | Sort-Object)) {
        $r   = $wl.Resources[$rName]
        $cat = Get-CatalogEntry -ResourceName $rName

        $resEntry = @{
            name      = $r.Name
            instances = @($r.Instances | Sort-Object { $_.key })
        }

        if ($cat) {
            $docHits++
            if ($cat.d) { $resEntry['desc']  = $cat.d }
            if ($cat.u) { $resEntry['doc']   = $cat.u }
            if ($cat.r) { $resEntry['roles'] = @($cat.r) }

            # Solo se incrustan las propiedades realmente presentes en el reporte
            if ($cat.p) {
                $seen = @{}
                foreach ($inst in $r.Instances) {
                    foreach ($row in $inst.rows) { $seen[$row.n] = $true }
                }
                foreach ($pn in $seen.Keys) {
                    $pEntry = $cat.p.PSObject.Properties | Where-Object { $_.Name -eq $pn } |
                              Select-Object -First 1
                    if ($pEntry) {
                        $rec = @{ d = $pEntry.Value.d }
                        if ($pEntry.Value.v) { $rec['v'] = @($pEntry.Value.v) }
                        if ($pEntry.Value.k) { $rec['k'] = 1 }
                        $propDocs["$rName.$pn"] = $rec
                    }
                }
            }
        }

        $resArray += $resEntry
    }
    $wlArray += @{ name = $wl.Name; resources = $resArray }
}

Write-Ok (tr "Recursos con documentacion: $docHits | propiedades documentadas: $($propDocs.Count)" "Resources with documentation: $docHits | documented properties: $($propDocs.Count)")

# ============================================================
#  TEXTOS DE LA INTERFAZ DEL REPORTE (bilingue EN/ES)
#  Nota: las descripciones de propiedades (catalog.json) van SIEMPRE en ingles.
# ============================================================
$ui = if ($script:Lang -eq 'EN') {
    @{
        statInstances='Instances'; statIdentical='Identical'; statDifferent='Different'
        statPartial='Partial'; statBaselineOnly='Baseline only'; statMissingBase='Missing in base'
        legIdentical='Identical to baseline'; legDifferent='Different values'
        legMissing='Missing in some tenant'; legOnlyBase='Only in the baseline'
        legElsewhere='Exists elsewhere, not in the baseline'
        filterPlaceholder='Filter by resource, instance or property...'
        btnOnlyDiff='Findings only'; btnExpand='Expand all'; btnCollapse='Collapse all'
        toggleHideEqual='Hide identical rows'; toggleDesc='Descriptions'
        btnCsvDetail='Export CSV detail'; btnCsvSummary='CSV summary'
        stGenerated='Generated '; stBaseline='Baseline: '; stConfigs='configurations'
        stInstAnalyzed='instances analyzed'; stDomains='Neutralized domains:'
        thProperty='Property'; lblBaseline='(baseline)'; lblAbsent='absent'
        lblNotSet='(not set)'; lblInstAbsent='(instance absent)'
        badgeFindings='findings'; badgeInst='inst.'; badgeDiff='diff.'
        emptyResults='No results with the current filters.'
        docsLabel='docs'; ariaDesc='Description of '; docTitle='Documentation of '
        roleRequired='Read role required: '; valuesLabel='Values: '; keyLabel='key'
        stIdentical='identical'; stDiffers='differs'; stPartial='partial'
        stOnlyBase='baseline only'; stNotBase='not in baseline'
        csvWorkload='Workload'; csvResource='Resource'; csvResourceDesc='Resource description'
        csvInstance='Instance'; csvStatus='Status'; csvProperty='Property'
        csvPropDesc='Property description'; csvDiffers='Differs'; csvPresentIn='Present in'
        csvDocumentation='Documentation'; csvDescription='Description'
        csvPropsWithDiff='Properties with differences'; csvTotalProps='Total properties'
        csvYes='YES'; csvNo='NO'; csvDetailFile='Detail'; csvSummaryFile='Summary'
        toastRowsExported=' rows exported'; toastNoRows='No rows to export with the current filters'
        toastInstExported=' instances exported'; toastNoInst='No instances to export with the current filters'
        footerReport='Baseline Report'; footerGeneratedBy='Generated by'
    }
} else {
    @{
        statInstances='Instancias'; statIdentical='Iguales'; statDifferent='Difieren'
        statPartial='Parciales'; statBaselineOnly='Solo baseline'; statMissingBase='Faltan en base'
        legIdentical='Identico a la baseline'; legDifferent='Valores distintos'
        legMissing='Falta en algun tenant'; legOnlyBase='Solo existe en la baseline'
        legElsewhere='Existe fuera, no en la baseline'
        filterPlaceholder='Filtrar por recurso, instancia o propiedad...'
        btnOnlyDiff='Solo hallazgos'; btnExpand='Expandir todo'; btnCollapse='Colapsar todo'
        toggleHideEqual='Ocultar filas iguales'; toggleDesc='Descripciones'
        btnCsvDetail='Exportar CSV detalle'; btnCsvSummary='CSV resumen'
        stGenerated='Generado '; stBaseline='Baseline: '; stConfigs='configuraciones'
        stInstAnalyzed='instancias analizadas'; stDomains='Dominios neutralizados:'
        thProperty='Propiedad'; lblBaseline='(baseline)'; lblAbsent='ausente'
        lblNotSet='(no definida)'; lblInstAbsent='(instancia ausente)'
        badgeFindings='hallazgos'; badgeInst='inst.'; badgeDiff='dif.'
        emptyResults='No hay resultados con los filtros actuales.'
        docsLabel='docs'; ariaDesc='Descripcion de '; docTitle='Documentacion de '
        roleRequired='Rol de lectura requerido: '; valuesLabel='Valores: '; keyLabel='clave'
        stIdentical='identico'; stDiffers='difiere'; stPartial='parcial'
        stOnlyBase='solo baseline'; stNotBase='no en baseline'
        csvWorkload='Workload'; csvResource='Recurso'; csvResourceDesc='Descripcion recurso'
        csvInstance='Instancia'; csvStatus='Estado'; csvProperty='Propiedad'
        csvPropDesc='Descripcion propiedad'; csvDiffers='Difiere'; csvPresentIn='Presente en'
        csvDocumentation='Documentacion'; csvDescription='Descripcion'
        csvPropsWithDiff='Propiedades con diferencia'; csvTotalProps='Total propiedades'
        csvYes='SI'; csvNo='NO'; csvDetailFile='Detalle'; csvSummaryFile='Resumen'
        toastRowsExported=' filas exportadas'; toastNoRows='No hay filas que exportar con los filtros actuales'
        toastInstExported=' instancias exportadas'; toastNoInst='No hay instancias que exportar con los filtros actuales'
        footerReport='Baseline Report'; footerGeneratedBy='Generado por'
    }
}

$payload = @{
    title     = $report.Title
    client    = $report.Client
    generated = $report.Generated
    labels    = @($report.Labels)
    sources   = @($report.Sources)
    baseIndex = $report.BaseIndex
    tenants   = $report.Tenants
    stats     = $report.Stats
    workloads = $wlArray
    propDocs  = $propDocs
    lang      = $script:Lang
    ui        = $ui
}

$json = $payload | ConvertTo-Json -Depth 20 -Compress
$json = $json.Replace('</script>', '<\/script>')

# Se usa el caracter real (no la entidad HTML) porque se inserta via textContent en JS
$clientLine = if ($report.Client) { " $([char]0x00B7) $($report.Client)" } else { "" }

# Fragmentos de marca para la plantilla (vacios si no hay branding)
$brandSuffix = if ($BrandName) { " | $BrandName" } else { "" }
$taglineTag  = if ($Tagline)   { "<div class=`"tagline`">$Tagline</div>" } else { "" }
$footerBrand = if ($BrandName) { "$($ui.footerGeneratedBy) <strong>$BrandName</strong> &middot; Microsoft365DSC $($ui.footerReport)" }
               else            { "<strong>Microsoft365DSC</strong> $($ui.footerReport)" }
$htmlLang    = if ($script:Lang -eq 'EN') { 'en' } else { 'es' }


# ============================================================
#  PLANTILLA HTML
# ============================================================

$html = @"
<!DOCTYPE html>
<html lang="$htmlLang">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$($report.Title)$brandSuffix</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {
    --brand:      #212d32;
    --brand-soft: #eef1f2;
    --brand-mid:  #48585f;

    --bg:        #f6f7f8;
    --panel:     #ffffff;
    --panel-2:   #f1f3f4;
    --border:    #e2e6e8;
    --border-2:  #cbd2d6;

    --text:      #212d32;
    --text-mid:  #4a585f;
    --text-dim:  #8794 9b;
    --text-dim:  #87949b;

    --accent:    #0f7a8c;
    --accent-sf: #e6f2f4;

    --ok:        #1c7a4a;
    --ok-sf:     #e8f4ee;
    --warn:      #a8700f;
    --warn-sf:   #fbf3e4;
    --err:       #b83227;
    --err-sf:    #fbeceb;
    --purple:    #5f3d94;
    --purple-sf: #efeaf7;
    --blue:      #1c5f96;
    --blue-sf:   #e8f0f7;

    --sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    --mono: 'JetBrains Mono', ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    --shadow:   0 1px 2px rgba(33,45,50,.05), 0 2px 8px rgba(33,45,50,.04);
    --shadow-2: 0 2px 4px rgba(33,45,50,.06), 0 8px 24px rgba(33,45,50,.06);
  }

  * { box-sizing: border-box; }
  html { scroll-behavior: smooth; }
  body {
    margin: 0; background: var(--bg); color: var(--text);
    font-family: var(--sans); font-size: 14px; line-height: 1.55;
    -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;
  }

  /* ---------- Cabecera ---------- */
  header {
    background: var(--panel); border-bottom: 1px solid var(--border);
    position: sticky; top: 0; z-index: 30; box-shadow: var(--shadow);
  }
  .head-top {
    display: flex; align-items: center; gap: 20px;
    padding: 15px 28px 13px; border-bottom: 1px solid var(--border);
  }
  .logo { height: 30px; width: auto; flex-shrink: 0; }
  .head-titles { flex: 1; min-width: 0; }
  h1 { margin: 0; font-size: 17px; font-weight: 600; letter-spacing: -.25px; color: var(--brand); }
  .meta { color: var(--text-dim); font-size: 12px; margin-top: 2px; }
  .tagline {
    font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px;
    color: var(--text-dim); font-weight: 500; white-space: nowrap;
    padding-left: 20px; border-left: 1px solid var(--border);
  }

  .head-body { padding: 15px 28px 17px; }

  /* ---------- Tarjetas de estado ---------- */
  .stats { display: flex; gap: 10px; flex-wrap: wrap; }
  .stat {
    background: var(--panel); border: 1px solid var(--border);
    border-left: 3px solid var(--border-2);
    border-radius: 8px; padding: 9px 16px; min-width: 108px;
    cursor: pointer; transition: transform .12s, box-shadow .12s;
    user-select: none;
  }
  .stat:hover { transform: translateY(-1px); box-shadow: var(--shadow-2); }
  .stat.sel { box-shadow: 0 0 0 2px var(--accent); }
  .stat .n { font-size: 21px; font-weight: 700; line-height: 1.15; letter-spacing: -.5px; }
  .stat .l {
    font-size: 9.5px; color: var(--text-dim); text-transform: uppercase;
    letter-spacing: .8px; font-weight: 500; margin-top: 1px;
  }
  .stat.tot  { border-left-color: var(--brand); cursor: default; }
  .stat.tot:hover { transform: none; box-shadow: none; }
  .stat.ok   { border-left-color: var(--ok);     background: var(--ok-sf); }
  .stat.ok   .n { color: var(--ok); }
  .stat.dif  { border-left-color: var(--err);    background: var(--err-sf); }
  .stat.dif  .n { color: var(--err); }
  .stat.par  { border-left-color: var(--warn);   background: var(--warn-sf); }
  .stat.par  .n { color: var(--warn); }
  .stat.only { border-left-color: var(--purple); background: var(--purple-sf); }
  .stat.only .n { color: var(--purple); }
  .stat.miss { border-left-color: var(--blue);   background: var(--blue-sf); }
  .stat.miss .n { color: var(--blue); }

  .legend {
    display: flex; gap: 18px; margin-top: 13px; flex-wrap: wrap;
    font-size: 11.5px; color: var(--text-mid);
  }
  .legend span { display: flex; align-items: center; gap: 7px; }

  .tenant-info {
    margin-top: 10px; font-size: 11px; color: var(--text-mid);
    font-family: var(--mono); padding: 7px 12px;
    background: var(--brand-soft); border: 1px solid var(--border);
    border-radius: 6px; display: inline-block;
  }

  /* ---------- Controles ---------- */
  .controls { display: flex; gap: 9px; margin-top: 15px; flex-wrap: wrap; align-items: center; }
  .controls input[type=text] {
    background: var(--panel); border: 1px solid var(--border-2); color: var(--text);
    padding: 8px 13px; border-radius: 7px; min-width: 270px;
    font-size: 13px; font-family: var(--sans);
    transition: border-color .12s, box-shadow .12s;
  }
  .controls input[type=text]::placeholder { color: var(--text-dim); }
  .controls input[type=text]:focus {
    outline: none; border-color: var(--accent);
    box-shadow: 0 0 0 3px var(--accent-sf);
  }
  .btn {
    background: var(--panel); border: 1px solid var(--border-2); color: var(--text-mid);
    padding: 8px 14px; border-radius: 7px; cursor: pointer;
    font-size: 12.5px; font-weight: 500; font-family: var(--sans);
    transition: all .12s; white-space: nowrap;
  }
  .btn:hover { border-color: var(--accent); color: var(--accent); background: var(--accent-sf); }
  .btn.active {
    background: var(--accent); border-color: var(--accent); color: #fff;
    box-shadow: 0 1px 3px rgba(15,122,140,.28);
  }
  .btn.primary {
    background: var(--brand); border-color: var(--brand); color: #fff;
  }
  .btn.primary:hover { background: #2d3d44; border-color: #2d3d44; color: #fff; }
  .sep { width: 1px; height: 22px; background: var(--border); margin: 0 3px; }
  .toggle {
    display: flex; align-items: center; gap: 7px;
    color: var(--text-mid); font-size: 12.5px; cursor: pointer; user-select: none;
  }
  .toggle input { accent-color: var(--accent); width: 15px; height: 15px; cursor: pointer; }

  /* ---------- Contenido ---------- */
  main { padding: 20px 28px 80px; max-width: 1900px; }

  .wl {
    margin-bottom: 12px; border: 1px solid var(--border);
    border-radius: 10px; overflow: hidden; background: var(--panel);
    box-shadow: var(--shadow);
  }
  .wl-head {
    padding: 13px 18px; cursor: pointer; background: var(--panel);
    display: flex; align-items: center; gap: 11px; user-select: none;
    transition: background .12s;
  }
  .wl-head:hover { background: var(--brand-soft); }
  .wl-title { font-weight: 600; font-size: 14.5px; flex: 1; letter-spacing: -.1px; color: var(--brand); }

  .res { border-top: 1px solid var(--border); }
  .res-head {
    background: var(--panel-2); padding: 9px 18px 9px 38px; cursor: pointer;
    display: flex; align-items: center; gap: 11px; user-select: none;
    font-family: var(--mono); font-size: 12.5px; font-weight: 500;
    transition: background .12s;
  }
  .res-head:hover { background: #e7eaec; }
  .res-title { flex: 1; color: var(--text-mid); }

  .inst { border-top: 1px solid var(--border); }
  .inst-head {
    padding: 8px 18px 8px 58px; cursor: pointer; background: var(--panel);
    display: flex; align-items: center; gap: 11px; user-select: none;
    font-size: 12.5px; transition: background .12s;
  }
  .inst-head:hover { background: var(--bg); }
  .inst-title { flex: 1; color: var(--text-mid); font-family: var(--mono); font-size: 11.5px; }

  .chev {
    color: var(--text-dim); font-size: 9px; width: 11px;
    transition: transform .16s ease; flex-shrink: 0;
  }
  .open > .chev, .open > .wl-head > .chev { transform: rotate(90deg); }

  .dot {
    width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0;
    display: inline-block; box-shadow: 0 0 0 2.5px rgba(255,255,255,.95);
  }
  .dot.same        { background: var(--ok); }
  .dot.diff        { background: var(--err); }
  .dot.partial     { background: var(--warn); }
  .dot.onlybase    { background: var(--purple); }
  .dot.missingbase { background: var(--blue); }

  .badge {
    font-size: 10.5px; padding: 2px 8px; border-radius: 11px; white-space: nowrap;
    background: var(--panel-2); border: 1px solid var(--border);
    color: var(--text-dim); font-weight: 500;
  }
  .badge.diff { background: var(--err-sf);    border-color: #f0cdca; color: var(--err); }
  .badge.par  { background: var(--warn-sf);   border-color: #eeddba; color: var(--warn); }
  .badge.only { background: var(--purple-sf); border-color: #d9cdee; color: var(--purple); }
  .badge.miss { background: var(--blue-sf);   border-color: #c4d8ea; color: var(--blue); }

  .body { display: none; }
  .open > .body { display: block; }

  /* ---------- Tablas ---------- */
  .tbl-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; background: var(--panel); }
  th, td {
    padding: 7px 14px; text-align: left; vertical-align: top;
    border-top: 1px solid var(--border); font-family: var(--mono);
  }
  th {
    background: var(--panel-2); color: var(--text-mid); font-weight: 600;
    font-size: 10px; text-transform: uppercase; letter-spacing: .7px;
    font-family: var(--sans); border-bottom: 1px solid var(--border-2);
  }
  th.baseline {
    border-bottom: 2px solid var(--accent); color: var(--accent);
    background: var(--accent-sf);
  }
  td.baseline { background: #fafcfc; }
  td.prop { color: var(--text-mid); width: 215px; white-space: nowrap; font-weight: 500; }
  td.val  { max-width: 380px; overflow-wrap: anywhere; white-space: pre-wrap; color: var(--text); }
  tr.row-diff td { background: var(--err-sf); }
  tr.row-diff td.prop { color: var(--err); }
  tr.row-diff td.baseline { background: #fbe8e6; }
  td.absent { color: var(--warn); font-style: italic; }

  /* ---------- Documentacion ---------- */
  .doc-link {
    color: var(--accent); text-decoration: none; font-size: 11px;
    padding: 2px 8px; border-radius: 11px; border: 1px solid var(--accent-sf);
    background: var(--accent-sf); font-weight: 500; white-space: nowrap;
    transition: all .12s; font-family: var(--sans);
  }
  .doc-link:hover { background: var(--accent); color: #fff; border-color: var(--accent); }

  .res-desc {
    padding: 9px 18px 11px 38px; background: var(--brand-soft);
    border-top: 1px solid var(--border); font-size: 12px;
    color: var(--text-mid); line-height: 1.5;
  }
  .res-desc .roles {
    display: block; margin-top: 5px; font-size: 11px;
    color: var(--text-dim); font-family: var(--mono);
  }

  .prop-wrap { display: flex; align-items: baseline; gap: 6px; }
  .info {
    display: inline-flex; align-items: center; justify-content: center;
    width: 14px; height: 14px; border-radius: 50%; flex-shrink: 0;
    background: var(--border); color: var(--panel);
    font-size: 9px; font-weight: 700; font-family: var(--sans);
    cursor: help; user-select: none; transition: background .12s;
  }
  .info:hover { background: var(--accent); }

  /* Tooltip flotante unico, fuera del flujo de la tabla */
  #tip {
    position: fixed; z-index: 9999; display: none;
    max-width: 380px; padding: 11px 14px; border-radius: 8px;
    background: var(--brand); color: #eef1f2; font-family: var(--sans);
    font-size: 11.5px; font-weight: 400; line-height: 1.5; text-align: left;
    box-shadow: 0 4px 12px rgba(33,45,50,.22), 0 12px 32px rgba(33,45,50,.18);
    pointer-events: none; overflow: hidden auto;
  }
  #tip.show { display: block; }
  #tip b { color: #fff; display: block; margin-bottom: 4px; font-size: 11px; }
  #tip .vals {
    display: block; margin-top: 7px; padding-top: 7px;
    border-top: 1px solid rgba(255,255,255,.18);
    font-family: var(--mono); font-size: 10.5px; color: #b9c4c8;
    word-break: break-word;
  }
  td.prop.is-key::after {
    content: '$($ui.keyLabel)'; margin-left: 7px; font-size: 8.5px; letter-spacing: .5px;
    text-transform: uppercase; color: var(--text-dim); font-family: var(--sans);
  }

  .empty {
    padding: 50px; text-align: center; color: var(--text-dim);
    background: var(--panel); border: 1px solid var(--border); border-radius: 10px;
  }
  .hidden { display: none !important; }

  /* ---------- Pie ---------- */
  footer {
    padding: 20px 28px 34px; color: var(--text-dim); font-size: 11.5px;
    border-top: 1px solid var(--border); margin-top: 30px;
    display: flex; justify-content: space-between; align-items: center;
    flex-wrap: wrap; gap: 10px;
  }
  footer strong { color: var(--brand-mid); font-weight: 600; }

  /* ---------- Aviso CSV ---------- */
  .toast {
    position: fixed; bottom: 26px; right: 26px; z-index: 100;
    background: var(--brand); color: #fff; padding: 12px 20px;
    border-radius: 8px; font-size: 13px; box-shadow: var(--shadow-2);
    opacity: 0; transform: translateY(10px); pointer-events: none;
    transition: opacity .2s, transform .2s;
  }
  .toast.show { opacity: 1; transform: translateY(0); }

  /* En ventanas bajas o estrechas la cabecera deja de ocupar
     casi toda la pantalla; si no, no queda sitio para el contenido. */
  @media (max-height: 760px), (max-width: 1100px) {
    header { position: static; }
  }
  @media (max-width: 900px) {
    .head-top, .head-body, main { padding-left: 14px; padding-right: 14px; }
    .stat { min-width: 88px; padding: 7px 11px; }
    .stat .n { font-size: 18px; }
    .legend { gap: 12px; font-size: 11px; }
    .controls input[type=text] { min-width: 100%; }
    td.prop { width: auto; min-width: 150px; }
    #tip { max-width: calc(100vw - 24px); }
  }

  @media print {
    header { position: static; box-shadow: none; }
    .controls, .toast { display: none !important; }
    .body { display: block !important; }
    .wl, table { box-shadow: none; page-break-inside: avoid; }
  }
</style>
</head>
<body>

<header>
  <div class="head-top">
    $logoTag
    <div class="head-titles">
      <h1 id="rTitle"></h1>
      <div class="meta" id="rMeta"></div>
    </div>
    $taglineTag
  </div>

  <div class="head-body">
    <div class="stats">
      <div class="stat tot">                        <div class="n" id="sTotal">0</div><div class="l">$($ui.statInstances)</div></div>
      <div class="stat ok"   data-f="same">         <div class="n" id="sSame">0</div> <div class="l">$($ui.statIdentical)</div></div>
      <div class="stat dif"  data-f="diff">         <div class="n" id="sDiff">0</div> <div class="l">$($ui.statDifferent)</div></div>
      <div class="stat par"  data-f="partial">      <div class="n" id="sPart">0</div> <div class="l">$($ui.statPartial)</div></div>
      <div class="stat only" data-f="onlybase">     <div class="n" id="sOnly">0</div> <div class="l">$($ui.statBaselineOnly)</div></div>
      <div class="stat miss" data-f="missingbase">  <div class="n" id="sMiss">0</div> <div class="l">$($ui.statMissingBase)</div></div>
    </div>

    <div class="legend">
      <span><i class="dot same"></i> $($ui.legIdentical)</span>
      <span><i class="dot diff"></i> $($ui.legDifferent)</span>
      <span><i class="dot partial"></i> $($ui.legMissing)</span>
      <span><i class="dot onlybase"></i> $($ui.legOnlyBase)</span>
      <span><i class="dot missingbase"></i> $($ui.legElsewhere)</span>
    </div>

    <div class="tenant-info" id="tInfo"></div>

    <div class="controls">
      <input type="text" id="filter" placeholder="$($ui.filterPlaceholder)">
      <button class="btn" id="bOnlyDiff">$($ui.btnOnlyDiff)</button>
      <div class="sep"></div>
      <button class="btn" id="bExpand">$($ui.btnExpand)</button>
      <button class="btn" id="bCollapse">$($ui.btnCollapse)</button>
      <label class="toggle"><input type="checkbox" id="cHideEqualRows"> $($ui.toggleHideEqual)</label>
      <label class="toggle"><input type="checkbox" id="cShowDesc" checked> $($ui.toggleDesc)</label>
      <div class="sep"></div>
      <button class="btn primary" id="bCsvDetail">$($ui.btnCsvDetail)</button>
      <button class="btn" id="bCsvSummary">$($ui.btnCsvSummary)</button>
    </div>
  </div>
</header>

<main id="root"></main>

<footer>
  <div>$footerBrand</div>
  <div id="fMeta"></div>
</footer>

<div class="toast" id="toast"></div>
<div id="tip"></div>

<script id="payload" type="application/json">$json</script>
<script>
(function () {
  const DATA = JSON.parse(document.getElementById('payload').textContent);
  const UI   = DATA.ui || {};
  const root = document.getElementById('root');
  const BASE = DATA.baseIndex || 0;

  const STATUS_LABEL = {
    same:        UI.stIdentical,
    diff:        UI.stDiffers,
    partial:     UI.stPartial,
    onlybase:    UI.stOnlyBase,
    missingbase: UI.stNotBase
  };
  const STATUS_BADGE = {
    diff: 'diff', partial: 'par', onlybase: 'only', missingbase: 'miss'
  };

  // ---------- Cabecera ----------
  document.getElementById('rTitle').textContent = DATA.title;
  document.getElementById('rMeta').textContent =
    UI.stGenerated + DATA.generated + '$clientLine' +
    '   |   ' + UI.stBaseline + DATA.labels[BASE] +
    '   |   ' + DATA.labels.length + ' ' + UI.stConfigs;
  document.getElementById('fMeta').textContent =
    DATA.stats.Total + ' ' + UI.stInstAnalyzed + '  ' + String.fromCharCode(183) + '  ' + DATA.generated;

  document.getElementById('sTotal').textContent = DATA.stats.Total;
  document.getElementById('sSame').textContent  = DATA.stats.Same;
  document.getElementById('sDiff').textContent  = DATA.stats.Diff;
  document.getElementById('sPart').textContent  = DATA.stats.Partial;
  document.getElementById('sOnly').textContent  = DATA.stats.OnlyBase;
  document.getElementById('sMiss').textContent  = DATA.stats.MissingBase;

  if (DATA.tenants) {
    const parts = [];
    DATA.labels.forEach(l => {
      const t = DATA.tenants[l];
      if (t && t.Prefix) parts.push(l + ' ' + String.fromCharCode(8594) + ' ' + t.Prefix);
    });
    const ti = document.getElementById('tInfo');
    if (parts.length) {
      ti.textContent = UI.stDomains + '   ' + parts.join('     |     ');
    } else {
      ti.style.display = 'none';
    }
  }

  // ---------- Estado ----------
  let onlyDiff = false;
  let hideEqualRows = false;
  let filterText = '';
  let statusFilter = null;
  let showDesc = true;

  const PDOCS = DATA.propDocs || {};

  function el(tag, cls, txt) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (txt !== undefined) e.textContent = txt;
    return e;
  }

  function toast(msg) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 2600);
  }

  // ---------- Filtro compartido ----------
  function instPasses(resName, inst) {
    if (onlyDiff && inst.status === 'same') return false;
    if (statusFilter && inst.status !== statusFilter) return false;
    if (filterText) {
      const hay = (resName + ' ' + inst.key + ' ' +
                   inst.rows.map(r => r.n).join(' ')).toLowerCase();
      if (hay.indexOf(filterText) === -1) return false;
    }
    return true;
  }

  // ---------- Render ----------
  function instTable(inst, resName) {
    const wrap = el('div', 'tbl-wrap');
    const t = el('table');
    const thead = el('thead');
    const htr = el('tr');
    htr.appendChild(el('th', null, UI.thProperty));
    DATA.labels.forEach((l, i) => {
      const isBase = (i === BASE);
      let label = l;
      if (isBase) label += '  ' + UI.lblBaseline;
      if (!inst.presence[i]) label += '  ' + String.fromCharCode(8212) + ' ' + UI.lblAbsent;
      htr.appendChild(el('th', isBase ? 'baseline' : null, label));
    });
    thead.appendChild(htr);
    t.appendChild(thead);

    const tb = el('tbody');
    inst.rows.forEach(r => {
      const tr = el('tr', r.d ? 'row-diff' : 'row-same');
      if (hideEqualRows && !r.d) tr.classList.add('hidden');
      const pd = PDOCS[resName + '.' + r.n];
      const tdProp = el('td', 'prop' + (pd && pd.k ? ' is-key' : ''));
      const wrap = el('span', 'prop-wrap');
      wrap.appendChild(el('span', null, r.n));
      if (pd && showDesc) {
        const ic = el('span', 'info', 'i');
        ic.tabIndex = 0;
        ic.setAttribute('role', 'button');
        ic.setAttribute('aria-label', UI.ariaDesc + r.n);
        ic.dataset.n = r.n;
        ic.dataset.d = pd.d;
        if (pd.v && pd.v.length) ic.dataset.v = pd.v.join(' | ');
        wrap.appendChild(ic);
      }
      tdProp.appendChild(wrap);
      tr.appendChild(tdProp);
      r.v.forEach((v, i) => {
        const td = el('td', 'val' + (i === BASE ? ' baseline' : ''));
        if (v === null || v === undefined) {
          td.classList.add('absent');
          td.textContent = inst.presence[i] ? UI.lblNotSet : UI.lblInstAbsent;
        } else {
          td.textContent = v;
        }
        tr.appendChild(td);
      });
      tb.appendChild(tr);
    });
    t.appendChild(tb);
    wrap.appendChild(t);
    return wrap;
  }

  function worstStatus(list) {
    const order = ['same','partial','diff','onlybase','missingbase'];
    let worst = 'same';
    list.forEach(s => { if (order.indexOf(s) > order.indexOf(worst)) worst = s; });
    return worst;
  }

  function build() {
    root.innerHTML = '';
    let anyVisible = false;

    DATA.workloads.forEach(wl => {
      const wlStatuses = [];
      let wlTotal = 0, wlFindings = 0;
      wl.resources.forEach(r => {
        r.instances.forEach(i => {
          wlStatuses.push(i.status);
          wlTotal++;
          if (i.status !== 'same') wlFindings++;
        });
      });

      const wlBox  = el('div', 'wl');
      const wlHead = el('div', 'wl-head');
      wlHead.appendChild(el('span', 'chev', String.fromCharCode(9654)));
      wlHead.appendChild(el('span', 'dot ' + worstStatus(wlStatuses)));
      wlHead.appendChild(el('span', 'wl-title', wl.name));
      wlHead.appendChild(el('span', 'badge', wlTotal + ' ' + UI.badgeInst));
      if (wlFindings) wlHead.appendChild(el('span', 'badge diff', wlFindings + ' ' + UI.badgeFindings));
      wlBox.appendChild(wlHead);

      const wlBody = el('div', 'body');
      let wlVisible = false;

      wl.resources.forEach(res => {
        const resStatuses = res.instances.map(i => i.status);
        const resFindings = res.instances.filter(i => i.status !== 'same').length;

        const resBox  = el('div', 'res');
        const resHead = el('div', 'res-head');
        resHead.appendChild(el('span', 'chev', String.fromCharCode(9654)));
        resHead.appendChild(el('span', 'dot ' + worstStatus(resStatuses)));
        resHead.appendChild(el('span', 'res-title', res.name));
        resHead.appendChild(el('span', 'badge', String(res.instances.length)));
        if (resFindings) resHead.appendChild(el('span', 'badge diff', String(resFindings)));
        if (res.doc) {
          const a = el('a', 'doc-link', UI.docsLabel + ' ' + String.fromCharCode(8599));
          a.href = res.doc; a.target = '_blank'; a.rel = 'noopener';
          a.title = UI.docTitle + res.name;
          a.addEventListener('click', ev => ev.stopPropagation());
          resHead.appendChild(a);
        }
        resBox.appendChild(resHead);

        const resBody = el('div', 'body');

        if (showDesc && (res.desc || (res.roles && res.roles.length))) {
          const dd = el('div', 'res-desc');
          if (res.desc) dd.appendChild(document.createTextNode(res.desc));
          if (res.roles && res.roles.length) {
            dd.appendChild(el('span', 'roles',
              UI.roleRequired + res.roles.join(', ')));
          }
          resBody.appendChild(dd);
        }
        let resVisible = false;

        res.instances.forEach(inst => {
          if (!instPasses(res.name, inst)) return;

          const iBox  = el('div', 'inst');
          const iHead = el('div', 'inst-head');
          iHead.appendChild(el('span', 'chev', String.fromCharCode(9654)));
          iHead.appendChild(el('span', 'dot ' + inst.status));
          iHead.appendChild(el('span', 'inst-title', inst.key));

          const nd = inst.rows.filter(r => r.d).length;
          if (nd) iHead.appendChild(el('span', 'badge diff', nd + ' ' + UI.badgeDiff));
          if (STATUS_BADGE[inst.status]) {
            iHead.appendChild(el('span', 'badge ' + STATUS_BADGE[inst.status],
                                 STATUS_LABEL[inst.status]));
          }
          iBox.appendChild(iHead);

          const iBody = el('div', 'body');
          iBody.appendChild(instTable(inst, res.name));
          iBox.appendChild(iBody);

          iHead.addEventListener('click', e => {
            e.stopPropagation();
            iBox.classList.toggle('open');
          });

          resBody.appendChild(iBox);
          resVisible = true;
        });

        if (!resVisible) return;
        resBox.appendChild(resBody);
        resHead.addEventListener('click', e => {
          e.stopPropagation();
          resBox.classList.toggle('open');
        });
        wlBody.appendChild(resBox);
        wlVisible = true;
      });

      if (!wlVisible) return;
      wlBox.appendChild(wlBody);
      wlHead.addEventListener('click', () => wlBox.classList.toggle('open'));
      root.appendChild(wlBox);
      anyVisible = true;
    });

    if (!anyVisible) {
      root.appendChild(el('div', 'empty', UI.emptyResults));
    }
  }

  // ---------- Exportacion CSV ----------
  function csvCell(v) {
    if (v === null || v === undefined) return '';
    let s = String(v);
    s = s.replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\r/g, ' ');
    if (s.indexOf('"') !== -1 || s.indexOf(';') !== -1 || s.indexOf(',') !== -1) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function downloadCsv(rows, filename) {
    // Separador coma (CSV estandar)
    const text = rows.map(r => r.map(csvCell).join(',')).join('\r\n');
    // BOM UTF-8 para que Excel respete los acentos
    const blob = new Blob([String.fromCharCode(0xFEFF) + text],
                          { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function stamp() {
    const d = new Date();
    const p = n => String(n).padStart(2, '0');
    return d.getFullYear() + p(d.getMonth()+1) + p(d.getDate()) + '-' +
           p(d.getHours()) + p(d.getMinutes());
  }

  function exportDetail() {
    const rows = [];
    const head = [UI.csvWorkload,UI.csvResource,UI.csvResourceDesc,UI.csvInstance,UI.csvStatus,
                  UI.csvProperty,UI.csvPropDesc,UI.csvDiffers];
    DATA.labels.forEach((l, i) => head.push(l + (i === BASE ? ' ' + UI.lblBaseline : '')));
    head.push(UI.csvPresentIn); head.push(UI.csvDocumentation);
    rows.push(head);

    let n = 0;
    DATA.workloads.forEach(wl => {
      wl.resources.forEach(res => {
        res.instances.forEach(inst => {
          if (!instPasses(res.name, inst)) return;

          const presentIn = DATA.labels
            .filter((l, i) => inst.presence[i]).join(' | ');

          inst.rows.forEach(r => {
            if (hideEqualRows && !r.d) return;
            const pdoc = PDOCS[res.name + '.' + r.n];
            const row = [
              wl.name, res.name, res.desc || '', inst.key,
              STATUS_LABEL[inst.status] || inst.status,
              r.n, pdoc ? pdoc.d : '', r.d ? UI.csvYes : UI.csvNo
            ];
            r.v.forEach((v, i) => {
              if (v === null || v === undefined) {
                row.push(inst.presence[i] ? UI.lblNotSet : UI.lblInstAbsent);
              } else { row.push(v); }
            });
            row.push(presentIn);
            row.push(res.doc || '');
            rows.push(row);
            n++;
          });
        });
      });
    });

    if (n === 0) { toast(UI.toastNoRows); return; }
    downloadCsv(rows, 'M365DSC-' + UI.csvDetailFile + '-' + stamp() + '.csv');
    toast(n + UI.toastRowsExported);
  }

  function exportSummary() {
    const rows = [];
    const head = [UI.csvWorkload,UI.csvResource,UI.csvDescription,UI.csvInstance,UI.csvStatus,
                  UI.csvPropsWithDiff,UI.csvTotalProps];
    DATA.labels.forEach(l => head.push(UI.csvPresentIn + ' ' + l));
    head.push(UI.csvDocumentation);
    rows.push(head);

    let n = 0;
    DATA.workloads.forEach(wl => {
      wl.resources.forEach(res => {
        res.instances.forEach(inst => {
          if (!instPasses(res.name, inst)) return;
          const row = [
            wl.name, res.name, res.desc || '', inst.key,
            STATUS_LABEL[inst.status] || inst.status,
            inst.rows.filter(r => r.d).length,
            inst.rows.length
          ];
          inst.presence.forEach(p => row.push(p ? UI.csvYes : UI.csvNo));
          row.push(res.doc || '');
          rows.push(row);
          n++;
        });
      });
    });

    if (n === 0) { toast(UI.toastNoInst); return; }
    downloadCsv(rows, 'M365DSC-' + UI.csvSummaryFile + '-' + stamp() + '.csv');
    toast(n + UI.toastInstExported);
  }

  // ---------- Eventos ----------
  let filterTimer = null;
  document.getElementById('filter').addEventListener('input', e => {
    clearTimeout(filterTimer);
    const val = e.target.value.toLowerCase().trim();
    filterTimer = setTimeout(() => { filterText = val; build(); }, 200);
  });

  document.getElementById('bOnlyDiff').addEventListener('click', e => {
    onlyDiff = !onlyDiff;
    e.target.classList.toggle('active', onlyDiff);
    if (onlyDiff) {
      statusFilter = null;
      document.querySelectorAll('.stat.sel').forEach(s => s.classList.remove('sel'));
    }
    build();
  });

  document.querySelectorAll('.stat[data-f]').forEach(card => {
    card.addEventListener('click', () => {
      const f = card.getAttribute('data-f');
      const wasSel = card.classList.contains('sel');
      document.querySelectorAll('.stat.sel').forEach(s => s.classList.remove('sel'));
      if (wasSel) {
        statusFilter = null;
      } else {
        statusFilter = f;
        card.classList.add('sel');
        onlyDiff = false;
        document.getElementById('bOnlyDiff').classList.remove('active');
      }
      build();
    });
  });

  document.getElementById('cHideEqualRows').addEventListener('change', e => {
    hideEqualRows = e.target.checked;
    build();
  });

  document.getElementById('cShowDesc').addEventListener('change', e => {
    showDesc = e.target.checked;
    build();
  });

  document.getElementById('bExpand').addEventListener('click', () => {
    document.querySelectorAll('.wl, .res, .inst').forEach(n => n.classList.add('open'));
  });

  document.getElementById('bCollapse').addEventListener('click', () => {
    document.querySelectorAll('.wl, .res, .inst').forEach(n => n.classList.remove('open'));
  });

  document.getElementById('bCsvDetail').addEventListener('click', exportDetail);
  document.getElementById('bCsvSummary').addEventListener('click', exportSummary);

  // ---------- Tooltip flotante ----------
  const TIP = document.getElementById('tip');

  function showTip(ic) {
    TIP.innerHTML = '';
    TIP.appendChild(el('b', null, ic.dataset.n));
    TIP.appendChild(document.createTextNode(ic.dataset.d));
    if (ic.dataset.v) {
      TIP.appendChild(el('span', 'vals', UI.valuesLabel + ic.dataset.v));
    }

    const M  = 10;
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    // Medir fuera de pantalla, con la altura ya acotada al viewport
    TIP.style.maxHeight = (vh - M * 2) + 'px';
    TIP.style.left = '-9999px';
    TIP.style.top  = '0px';
    TIP.classList.add('show');

    const r  = ic.getBoundingClientRect();
    const tw = TIP.offsetWidth;
    const th = TIP.offsetHeight;

    // Horizontal: derecha del icono; si no cabe, izquierda; si tampoco, pegado al borde
    let x = r.right + M;
    if (x + tw > vw - M) x = r.left - tw - M;
    if (x < M) x = Math.max(M, vw - tw - M);

    // Vertical: centrado respecto al icono y contenido en la ventana
    let y = r.top + r.height / 2 - th / 2;
    if (y + th > vh - M) y = vh - th - M;
    if (y < M) y = M;

    TIP.style.left = Math.round(x) + 'px';
    TIP.style.top  = Math.round(y) + 'px';
  }

  function hideTip() { TIP.classList.remove('show'); }

  // Delegacion: sobrevive a cada re-render de build()
  document.addEventListener('mouseover', e => {
    const ic = e.target.closest('.info');
    if (ic) showTip(ic);
  });
  document.addEventListener('mouseout', e => {
    if (e.target.closest('.info')) hideTip();
  });
  // Accesible por teclado y utilizable en tactil
  document.addEventListener('focusin', e => {
    const ic = e.target.closest('.info');
    if (ic) showTip(ic);
  });
  document.addEventListener('focusout', e => {
    if (e.target.closest('.info')) hideTip();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') hideTip();
  });
  window.addEventListener('scroll', hideTip, true);
  window.addEventListener('resize', hideTip);

  build();
})();
</script>

</body>
</html>
"@


# ============================================================
#  ESCRITURA
# ============================================================
$outDir = Split-Path $OutputPath -Parent
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$html | Out-File -FilePath $OutputPath -Encoding UTF8

$sizeMb = [math]::Round((Get-Item $OutputPath).Length / 1MB, 2)
Write-Ok (tr "Reporte generado: $OutputPath ($sizeMb MB)" "Report generated: $OutputPath ($sizeMb MB)")

if ($sizeMb -gt 25) {
    Write-Warn (tr "El fichero es grande. El render inicial puede tardar unos segundos." "The file is large. Initial rendering may take a few seconds.")
    Write-Warn (tr "Considera generar reportes separados por workload." "Consider generating separate reports per workload.")
}

$openYes = if ($script:Lang -eq 'EN') { 'Y' } else { 'S' }
if ((Read-WithDefault (tr "`n Abrir en el navegador? (S/N)" "`n Open in the browser? (Y/N)") $openYes).ToUpper() -eq $openYes) {
    Start-Process $OutputPath
}
}


# ============================================================
#  PASO - REMOVE
# ============================================================
function Invoke-RemoveStep {
$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "  [X]  $m" -ForegroundColor Red }

function Read-WithDefault {
    param([string]$Prompt, [string]$Default)
    $v = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $false)
    $d = if ($Default) { 'S' } else { 'N' }
    while ($true) {
        $v = Read-Host "$Prompt (S/N) [$d]"
        if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
        switch ($v.Trim().ToUpper()) {
            'S' { return $true }
            'Y' { return $true }
            'N' { return $false }
            default { Write-Warn "Responde S o N" }
        }
    }
}


# ============================================================
#  ENTRADA
# ============================================================
Clear-Host
Write-Host ("=" * 66) -ForegroundColor Red
Write-Host " DESMANTELAR APP REGISTRATION DE MICROSOFT365DSC" -ForegroundColor Red
Write-Host ("=" * 66) -ForegroundColor Red

Write-Step "Identificar la aplicacion"
Write-Host "   1) Por nombre" -ForegroundColor Gray
Write-Host "   2) Por Application (client) ID" -ForegroundColor Gray
$modo = Read-WithDefault " Selecciona" "1"

$AppDisplayName = $null; $AppIdInput = $null
if ($modo -eq '2') {
    $AppIdInput = (Read-Host " Application ID").Trim()
} else {
    $AppDisplayName = Read-WithDefault " Nombre de la aplicacion" "M365DSC-Export"
}

$TenantHint = (Read-Host " TenantId o dominio (opcional, vacio = el del usuario)").Trim()


# ============================================================
#  MODULOS
# ============================================================
Write-Step "Preparando modulos de Microsoft Graph"

$graphModules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.Identity.DirectoryManagement'
)

foreach ($m in $graphModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Warn "Instalando $m ..."
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
}

$versionSets = foreach ($m in $graphModules) {
    ,@(Get-Module -ListAvailable -Name $m | Select-Object -ExpandProperty Version)
}
$commonVersions = $versionSets[0]
foreach ($set in $versionSets[1..($versionSets.Count - 1)]) {
    $commonVersions = $commonVersions | Where-Object { $set -contains $_ }
}
$targetVersion = $commonVersions | Sort-Object -Descending | Select-Object -First 1

if (-not $targetVersion) {
    $targetVersion = (Get-Module Microsoft.Graph.Authentication -ListAvailable |
                      Sort-Object Version -Descending | Select-Object -First 1).Version
    foreach ($m in $graphModules) {
        if (-not (Get-Module -ListAvailable -Name $m | Where-Object { $_.Version -eq $targetVersion })) {
            Install-Module $m -RequiredVersion $targetVersion -Scope CurrentUser -Force -AllowClobber
        }
    }
}

Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
foreach ($m in $graphModules) {
    $mod = Get-Module -ListAvailable -Name $m |
           Where-Object { $_.Version -eq $targetVersion } | Select-Object -First 1
    if (-not $mod) { Write-Err "No se encuentra $m $targetVersion. Abortando."; return }
    Import-Module $mod.Path -Force -ErrorAction Stop
}
Write-Ok "Modulos cargados ($targetVersion)"


# ============================================================
#  CONEXION
# ============================================================
Write-Step "Conectando a Microsoft Graph"

$connectArgs = @{
    Scopes = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'RoleManagement.ReadWrite.Directory'
        'Directory.ReadWrite.All'
    )
    NoWelcome = $true
}
if ($TenantHint) { $connectArgs['TenantId'] = $TenantHint }

Connect-MgGraph @connectArgs

$ctx = Get-MgContext
Write-Ok "Tenant : $($ctx.TenantId)"
Write-Ok "Usuario: $($ctx.Account)"


# ============================================================
#  LOCALIZAR LA APP
# ============================================================
Write-Step "Buscando la aplicacion"

if ($AppIdInput) {
    $app = Get-MgApplication -Filter "appId eq '$AppIdInput'" -ErrorAction SilentlyContinue |
           Select-Object -First 1
} else {
    $matches = @(Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue)
    if ($matches.Count -gt 1) {
        Write-Warn "Hay $($matches.Count) apps con ese nombre:"
        for ($i = 0; $i -lt $matches.Count; $i++) {
            Write-Host ("    [{0}] {1}  (creada {2})" -f ($i+1), $matches[$i].AppId,
                        $matches[$i].CreatedDateTime) -ForegroundColor Gray
        }
        $sel = 0
        while ($sel -lt 1 -or $sel -gt $matches.Count) {
            $v = Read-Host " Cual eliminar (1-$($matches.Count))"
            [int]::TryParse($v, [ref]$sel) | Out-Null
        }
        $app = $matches[$sel - 1]
    } else {
        $app = $matches | Select-Object -First 1
    }
}

if (-not $app) {
    Write-Err "No se encontro la aplicacion."
    Disconnect-MgGraph | Out-Null
    return
}

$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue |
      Select-Object -First 1


# ============================================================
#  INVENTARIO
# ============================================================
Write-Step "Inventario de lo que se va a eliminar"

Write-Host "  Aplicacion    : $($app.DisplayName)"
Write-Host "  AppId         : $($app.AppId)"
Write-Host "  ObjectId      : $($app.Id)"
Write-Host "  Creada        : $($app.CreatedDateTime)"

$certCount   = @($app.KeyCredentials).Count
$secretCount = @($app.PasswordCredentials).Count
Write-Host "  Certificados  : $certCount"
Write-Host "  Secrets       : $secretCount"

$thumbprints = @()
foreach ($kc in $app.KeyCredentials) {
    if ($kc.CustomKeyIdentifier) {
        $tp = [System.BitConverter]::ToString($kc.CustomKeyIdentifier).Replace('-','')
        $thumbprints += $tp
        Write-Host "     - $tp (expira $($kc.EndDateTime.ToString('yyyy-MM-dd')))" -ForegroundColor DarkGray
    }
}

$roleMemberships = @()
$appRoleAssigns  = @()

if ($sp) {
    Write-Host "  ServicePrincipal: $($sp.Id)"

    # Roles de directorio
    $allRoles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue
    foreach ($role in $allRoles) {
        try {
            $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop)
            foreach ($mem in $members) {
                $memId = $mem.Id
                if (-not $memId -and $mem.AdditionalProperties) { $memId = $mem.AdditionalProperties['id'] }
                if ($memId -eq $sp.Id) {
                    $roleMemberships += [pscustomobject]@{ Id = $role.Id; Name = $role.DisplayName }
                    break
                }
            }
        } catch { }
    }

    Write-Host "  Roles de directorio: $($roleMemberships.Count)"
    $roleMemberships | ForEach-Object { Write-Host "     - $($_.Name)" -ForegroundColor DarkGray }

    # Permisos concedidos
    $appRoleAssigns = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue)
    Write-Host "  Permisos concedidos: $($appRoleAssigns.Count)"
} else {
    Write-Warn "No existe Service Principal para esta app"
}

Write-Host ""
Write-Host ("!" * 66) -ForegroundColor Red
Write-Host " ESTA OPERACION ES DESTRUCTIVA" -ForegroundColor Red
Write-Host ("!" * 66) -ForegroundColor Red

$confirm = Read-Host "`n Escribe el nombre exacto de la app para confirmar"
if ($confirm -ne $app.DisplayName) {
    Write-Warn "El nombre no coincide. Cancelado."
    Disconnect-MgGraph | Out-Null
    return
}


# ============================================================
#  1. ROLES DE DIRECTORIO
# ============================================================
if ($roleMemberships.Count -gt 0) {
    Write-Step "Quitando membresias de roles de directorio"

    foreach ($r in $roleMemberships) {
        $err = $null
        try {
            Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $r.Id -DirectoryObjectId $sp.Id `
                -ErrorAction Stop -ErrorVariable err | Out-Null
        } catch { $err = $_ }

        if ($err) {
            $msg = if ($err -is [System.Management.Automation.ErrorRecord]) { $err.Exception.Message }
                   else { ($err | Select-Object -First 1).ToString() }
            if ($msg -match 'does not exist|Request_ResourceNotFound') {
                Write-Ok "$($r.Name) (ya no era miembro)"
            } else {
                Write-Warn "Fallo al quitar '$($r.Name)': $msg"
            }
        } else {
            Write-Ok "$($r.Name) (quitado)"
        }
    }
}


# ============================================================
#  2. PERMISOS CONCEDIDOS
# ============================================================
if ($appRoleAssigns.Count -gt 0) {
    Write-Step "Revocando permisos concedidos"

    $removed = 0; $failed = 0
    foreach ($a in $appRoleAssigns) {
        try {
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -AppRoleAssignmentId $a.Id -ErrorAction Stop | Out-Null
            $removed++
        } catch {
            if ($_.Exception.Message -match 'does not exist|ResourceNotFound') { $removed++ }
            else { $failed++; Write-Warn "Fallo en $($a.Id): $($_.Exception.Message)" }
        }
    }
    Write-Ok "$removed revocados | $failed fallidos"
}


# ============================================================
#  3. OAUTH2 GRANTS (consentimientos delegados, si los hubiera)
# ============================================================
if ($sp) {
    # Via REST (Invoke-MgGraphRequest) para no depender del modulo
    # Microsoft.Graph.Identity.SignIns (donde viven los cmdlets Oauth2PermissionGrant).
    $grants = @()
    try {
        $resp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/oauth2PermissionGrants" `
            -ErrorAction Stop
        if ($resp.value) { $grants = @($resp.value) }
    } catch {
        Write-Warn "No se pudieron consultar los consentimientos delegados: $($_.Exception.Message)"
    }
    if ($grants.Count -gt 0) {
        Write-Step "Eliminando consentimientos delegados ($($grants.Count))"
        foreach ($grant in $grants) {
            try {
                Invoke-MgGraphRequest -Method DELETE `
                    -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" -ErrorAction Stop
                Write-Ok "Grant $($grant.id)"
            } catch {
                Write-Warn "Fallo: $($_.Exception.Message)"
            }
        }
    }
}


# ============================================================
#  4. CREDENCIALES DE LA APP
# ============================================================
if ($certCount -gt 0 -or $secretCount -gt 0) {
    Write-Step "Eliminando credenciales de la aplicacion"
    try {
        Update-MgApplication -ApplicationId $app.Id `
            -KeyCredentials @() -PasswordCredentials @() -ErrorAction Stop
        Write-Ok "$certCount certificados y $secretCount secrets eliminados"
    } catch {
        Write-Warn "Fallo al limpiar credenciales: $($_.Exception.Message)"
    }
}


# ============================================================
#  5. SERVICE PRINCIPAL
# ============================================================
if ($sp) {
    Write-Step "Eliminando Service Principal"
    try {
        Remove-MgServicePrincipal -ServicePrincipalId $sp.Id -ErrorAction Stop
        Write-Ok "Service Principal eliminado ($($sp.Id))"
    } catch {
        Write-Warn "Fallo: $($_.Exception.Message)"
    }
}


# ============================================================
#  6. APP REGISTRATION
# ============================================================
Write-Step "Eliminando App Registration"
try {
    Remove-MgApplication -ApplicationId $app.Id -ErrorAction Stop
    Write-Ok "Aplicacion eliminada ($($app.AppId))"
    Write-Warn "Queda en 'Deleted applications' 30 dias. Se puede restaurar o purgar desde el portal."
} catch {
    Write-Err "Fallo al eliminar la aplicacion: $($_.Exception.Message)"
}

# Purga definitiva opcional
if (Read-YesNo "`n Purgar definitivamente (sin periodo de restauracion)?" $false) {
    try {
        Start-Sleep -Seconds 5
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/$($app.Id)" -ErrorAction Stop
        Write-Ok "Purgada definitivamente"
    } catch {
        Write-Warn "No se pudo purgar: $($_.Exception.Message)"
        Write-Warn "Hazlo desde Entra ID > App registrations > Deleted applications"
    }
}

Disconnect-MgGraph | Out-Null


# ============================================================
#  7. CERTIFICADO LOCAL
# ============================================================
if ($thumbprints.Count -gt 0) {
    Write-Step "Certificado local"

    $localCerts = @()
    foreach ($store in @('Cert:\CurrentUser\My','Cert:\LocalMachine\My')) {
        foreach ($tp in $thumbprints) {
            $c = Get-ChildItem $store -ErrorAction SilentlyContinue |
                 Where-Object { $_.Thumbprint -eq $tp }
            if ($c) { $localCerts += [pscustomobject]@{ Store = $store; Cert = $c } }
        }
    }

    if ($localCerts.Count -eq 0) {
        Write-Ok "No hay certificados asociados en el almacen local"
    } else {
        foreach ($lc in $localCerts) {
            Write-Host "    $($lc.Store)  $($lc.Cert.Thumbprint)  $($lc.Cert.Subject)" -ForegroundColor Gray
        }
        if (Read-YesNo " Eliminarlos del almacen local?" $true) {
            foreach ($lc in $localCerts) {
                try {
                    Remove-Item -Path (Join-Path $lc.Store $lc.Cert.Thumbprint) -Force -ErrorAction Stop
                    Write-Ok "Eliminado $($lc.Cert.Thumbprint) de $($lc.Store)"
                } catch {
                    Write-Warn "Fallo (puede requerir admin para LocalMachine): $($_.Exception.Message)"
                }
            }
        }
    }
}


# ============================================================
#  8. FICHEROS GENERADOS
# ============================================================
Write-Step "Ficheros generados"

$searchDir = Read-WithDefault " Carpeta donde buscar ficheros generados (vacio = omitir)" $PWD.Path

if ($searchDir -and (Test-Path $searchDir)) {
    $patterns = @(
        "$($app.DisplayName).cer"
        "$($app.DisplayName).pfx"
        "M365DSC-Export-Ready.ps1"
        "M365DSC-Export-Main.ps1"
        "M365DSC-Export-SPO.ps1"
    )

    $files = @()
    foreach ($p in $patterns) {
        $f = Get-ChildItem -Path $searchDir -Filter $p -File -ErrorAction SilentlyContinue
        if ($f) { $files += $f }
    }

    if ($files.Count -eq 0) {
        Write-Ok "No se encontraron ficheros generados"
    } else {
        $files | ForEach-Object { Write-Host "    $($_.FullName)" -ForegroundColor Gray }
        if (Read-YesNo " Eliminarlos?" $false) {
            foreach ($f in $files) {
                try {
                    Remove-Item $f.FullName -Force -ErrorAction Stop
                    Write-Ok "Eliminado $($f.Name)"
                } catch { Write-Warn "Fallo: $($_.Exception.Message)" }
            }
        }
    }
}

Write-Host "`n" -NoNewline
Write-Host ("=" * 66) -ForegroundColor Green
Write-Host " DESMANTELAMIENTO COMPLETADO" -ForegroundColor Green
Write-Host ("=" * 66) -ForegroundColor Green
Write-Host ""
}


# ============================================================
#  ACCIONES INLINE DEL MENU (no requieren sesion limpia)
# ============================================================
function Invoke-QueryStep {
    $cfg = Join-Path $Root 'Scripts\ConfigurationFile.ps1'
    Clear-Host
    Write-Host ("=" * 66) -ForegroundColor Cyan
    Write-Host (tr " GENERAR LA CONSULTA DE EXPORT" " GENERATE THE EXPORT QUERY") -ForegroundColor Cyan
    Write-Host ("=" * 66) -ForegroundColor Cyan
    Write-Host ""
    Write-Host (tr " 1. Abre el generador de consultas de Microsoft365DSC:" " 1. Open the Microsoft365DSC query generator:") -ForegroundColor Gray
    Write-Host "      https://export.microsoft365dsc.com/" -ForegroundColor White
    Write-Host (tr " 2. Selecciona los componentes/workloads a exportar." " 2. Select the components/workloads to export.") -ForegroundColor Gray
    Write-Host (tr " 3. Descarga o copia el script y guardalo como:" " 3. Download or copy the script and save it as:") -ForegroundColor Gray
    Write-Host "      $cfg" -ForegroundColor White
    Write-Host ""
    Write-Host (tr " Recordatorios:" " Reminders:") -ForegroundColor DarkCyan
    Write-Host (tr "   - Los componentes Fabric no usan certificado y fallaran." "   - Fabric components do not use a certificate and will fail.") -ForegroundColor DarkGray
    Write-Host (tr "   - Los sitios de SharePoint pueden tardar mucho (excluyelos si" "   - SharePoint sites can take a long time (exclude them if") -ForegroundColor DarkGray
    Write-Host (tr "     no necesitas un export completo del tenant)." "     you don't need a full tenant export).") -ForegroundColor DarkGray
    Write-Host ""

    if (Read-YesNo (tr " Abrir ahora el sitio en el navegador?" " Open the site in the browser now?") $true) {
        try { Start-Process "https://export.microsoft365dsc.com/" } catch { Write-Warn (tr "No se pudo abrir el navegador." "Could not open the browser.") }
    }

    Write-Host ""
    if (Test-Path $cfg) { Write-Ok (tr "Detectado: $cfg" "Detected: $cfg") }
    else { Write-Warn (tr "Aun no existe $cfg. Guardalo ahi cuando termines." "$cfg does not exist yet. Save it there when you finish.") }
}

function Invoke-VerifyTenantsStep {
    $tenants = Join-Path $Root 'Tenants'
    Clear-Host
    Write-Host ("=" * 66) -ForegroundColor Cyan
    Write-Host (tr " VERIFICAR CONFIGURACIONES DE TENANTS" " VERIFY TENANT CONFIGURATIONS") -ForegroundColor Cyan
    Write-Host ("=" * 66) -ForegroundColor Cyan
    Write-Host ""
    Write-Host (tr " Coloca el M365TenantConfig.ps1 de cada tenant en su carpeta:" " Place each tenant's M365TenantConfig.ps1 in its folder:") -ForegroundColor Gray
    Write-Host "   $tenants\Baseline\M365TenantConfig.ps1     $(tr '(el tenant mas completo / de referencia)' '(the most complete / reference tenant)')" -ForegroundColor DarkGray
    Write-Host "   $tenants\SnapshotA\M365TenantConfig.ps1" -ForegroundColor DarkGray
    Write-Host "   $tenants\SnapshotB\M365TenantConfig.ps1" -ForegroundColor DarkGray
    Write-Host ""

    if (-not (Test-Path $tenants)) {
        Write-Warn (tr "No existe la carpeta $tenants. Ejecuta primero 'Preparar entorno'." "Folder $tenants does not exist. Run 'Prepare environment' first.")
        return
    }

    $files = @(Get-ChildItem $tenants -Recurse -Filter 'M365TenantConfig.ps1' -ErrorAction SilentlyContinue)
    if ($files.Count -eq 0) {
        Write-Warn (tr "No se encontro ningun M365TenantConfig.ps1 todavia." "No M365TenantConfig.ps1 found yet.")
        return
    }

    Write-Step (tr "Archivos encontrados" "Files found")
    $files | ForEach-Object {
        $kb = [math]::Round($_.Length / 1KB, 1)
        $flag = if ($kb -lt 1) { tr '[!] muy pequeno' '[!] very small' } else { '' }
        Write-Host ("    {0,-16} {1,8} KB  {2}" -f $_.Directory.Name, $kb, $flag) -ForegroundColor Gray
    }
    Write-Host ""
    if ($files.Count -ge 2) { Write-Ok (tr "$($files.Count) configuraciones listas para comparar" "$($files.Count) configurations ready to compare") }
    else { Write-Warn (tr "Se necesitan al menos 2 configuraciones para comparar" "At least 2 configurations are needed to compare") }
}


# ============================================================
#  ESTADO Y MENU
# ============================================================
function Get-State {
    $scripts  = Join-Path $Root 'Scripts'
    $export   = Join-Path $Root 'Export'
    $tenants  = Join-Path $Root 'Tenants'
    $reportes = Join-Path $Root 'Reports'

    $moduleOk = [bool](Get-Module -ListAvailable -Name Microsoft365DSC)
    $foldersOk = (Test-Path $scripts) -and (Test-Path $export) -and (Test-Path $tenants) -and (Test-Path $reportes)

    $tenantFiles = @()
    if (Test-Path $tenants) {
        $tenantFiles = @(Get-ChildItem $tenants -Recurse -Filter 'M365TenantConfig.ps1' -ErrorAction SilentlyContinue)
    }
    $reportFiles = @()
    if (Test-Path $reportes) {
        $reportFiles = @(Get-ChildItem $reportes -Filter '*.html' -ErrorAction SilentlyContinue)
    }

    return [ordered]@{
        Setup      = ($moduleOk -and $foldersOk)
        Query      = (Test-Path (Join-Path $scripts 'ConfigurationFile.ps1'))
        Provision  = ((Test-Path (Join-Path $export 'M365DSC-Export-Main.ps1')) -or
                      @(Get-ChildItem $export -Filter '*.cer' -ErrorAction SilentlyContinue).Count -gt 0)
        Export     = (Test-Path (Join-Path $export 'M365TenantConfig.ps1'))
        Tenants    = ($tenantFiles.Count -ge 2)
        Report     = ($reportFiles.Count -gt 0)
    }
}

function Show-Menu {
    param($State)

    # Determinar el siguiente paso pendiente (orden logico principal)
    $order = @('Setup','Query','Provision','Export','Tenants','Report')
    $next = $null
    foreach ($k in $order) { if (-not $State[$k]) { $next = $k; break } }

    function Mark { param($key)
        if ($State[$key]) { return '[OK]' } else { return '[  ]' }
    }
    $nextTag = tr '  <-- SIGUIENTE' '  <-- NEXT'
    function Tag { param($key)
        if ($next -eq $key) { return $nextTag } else { return '' }
    }

    Clear-Host
    Write-Host ("=" * 68) -ForegroundColor Cyan
    Write-Host "  MICROSOFT 365 DSC - BASELINE REPORT" -ForegroundColor Cyan
    Write-Host "  Creado por / Created by: Christian Aguilera - FendariGroup" -ForegroundColor DarkCyan
    Write-Host ("=" * 68) -ForegroundColor Cyan
    Write-Host ""
    Write-Host (tr "  PROCESO 1 - EXPORTAR LA CONFIGURACION DE CADA TENANT" "  PROCESS 1 - EXPORT EACH TENANT'S CONFIGURATION") -ForegroundColor White
    Write-Host ("   1) {0} {1}{2}" -f (Mark 'Setup'),     (tr 'Preparar entorno (carpetas, modulo, dependencias)' 'Prepare environment (folders, module, dependencies)'), (Tag 'Setup'))     -ForegroundColor Gray
    Write-Host ("   2) {0} {1}{2}" -f (Mark 'Query'),     (tr 'Generar consulta de export -> ConfigurationFile.ps1' 'Generate export query -> ConfigurationFile.ps1'),   (Tag 'Query'))    -ForegroundColor Gray
    Write-Host ("   3) {0} {1}{2}" -f (Mark 'Provision'), (tr 'Provisionar App Registration (certificado)' 'Provision App Registration (certificate)'),               (Tag 'Provision')) -ForegroundColor Gray
    Write-Host ("   4) {0} {1}{2}" -f (Mark 'Export'),    (tr 'Exportar el tenant -> M365TenantConfig.ps1' 'Export the tenant -> M365TenantConfig.ps1'),               (Tag 'Export'))    -ForegroundColor Gray
    Write-Host ("   5) [  ] {0}" -f (tr 'Eliminar App Registration (limpieza tras exportar)' 'Delete App Registration (cleanup after export)')) -ForegroundColor Gray
    Write-Host ""
    Write-Host (tr "  PROCESO 2 - REPORTE COMPARATIVO ENTRE TENANTS" "  PROCESS 2 - COMPARATIVE REPORT ACROSS TENANTS") -ForegroundColor White
    Write-Host ("   6) {0} {1}{2}" -f (Mark 'Tenants'), (tr 'Verificar los M365TenantConfig.ps1 en Tenants\' 'Verify the M365TenantConfig.ps1 files in Tenants\'), (Tag 'Tenants'))   -ForegroundColor Gray
    Write-Host ("   7) {0} {1}{2}" -f (Mark 'Report'),  (tr 'Generar el reporte HTML de baseline' 'Generate the HTML baseline report'),                     (Tag 'Report'))    -ForegroundColor Gray
    Write-Host ""
    Write-Host ("   Q) {0}" -f (tr 'Salir' 'Quit')) -ForegroundColor Gray
    Write-Host ""
    Write-Host ("-" * 68) -ForegroundColor DarkGray
    if ($next) {
        $labels = if ($script:Lang -eq 'EN') {
            @{ Setup='Prepare the environment'; Query='Generate the export query';
               Provision='Provision the App Registration'; Export='Export the tenant';
               Tenants='Place the tenant configurations'; Report='Generate the report' }
        } else {
            @{ Setup='Preparar el entorno'; Query='Generar la consulta de export';
               Provision='Provisionar la App Registration'; Export='Exportar el tenant';
               Tenants='Colocar las configuraciones de los tenants'; Report='Generar el reporte' }
        }
        Write-Host ("  {0}: {1}." -f (tr 'Sugerencia' 'Suggestion'), $labels[$next]) -ForegroundColor Yellow
    } else {
        Write-Host (tr "  Todos los pasos completados. Puedes regenerar el reporte (7)." "  All steps completed. You can regenerate the report (7).") -ForegroundColor Green
    }
    Write-Host ""
}

function Start-MenuLoop {
    while ($true) {
        $state = Get-State
        Show-Menu -State $state

        $choice = (Read-Host (tr "  Elige una opcion" "  Choose an option")).Trim().ToUpper()
        switch ($choice) {
            '1' { Invoke-ChildStep 'Setup' }
            '2' { Invoke-QueryStep }
            '3' { Invoke-ChildStep 'Provision' }
            '4' { Invoke-ChildStep 'Export' }
            '5' { Invoke-ChildStep 'Remove' }
            '6' { Invoke-VerifyTenantsStep }
            '7' { Invoke-ChildStep 'Report' }
            'Q' { Write-Host (tr "  Hasta luego." "  Goodbye.") -ForegroundColor Cyan; return }
            ''  { }
            default { Write-Warn (tr "Opcion no valida: $choice" "Invalid option: $choice") }
        }

        if ($choice -ne 'Q') {
            Write-Host ""
            Read-Host (tr "  Presiona Enter para volver al menu" "  Press Enter to return to the menu") | Out-Null
        }
    }
}


# ============================================================
#  PUNTO DE ENTRADA
# ============================================================
switch ($Step) {
    'Setup'     { Invoke-SetupStep }
    'Provision' { Invoke-ProvisionStep }
    'Export'    { Invoke-ExportStep }
    'Report'    { Invoke-ReportStep }
    'Remove'    { Invoke-RemoveStep }
    default     {
        # Modo menu: preguntar idioma si no se indico por parametro
        if (-not $Lang) { $script:Lang = Select-Language }
        Start-MenuLoop
    }
}
