# M365DSC Report

Orquestador en PowerShell para **auditar y comparar la configuración de varios
tenants de Microsoft 365** usando [Microsoft365DSC](https://microsoft365dsc.com/).

Un **único script con menú** (`Invoke-M365DSCReport.ps1`) recorre todo el
proceso —de la exportación de cada tenant al reporte HTML comparativo— y
**detecta automáticamente en qué paso te encuentras**, sugiriendo el siguiente.

El resultado final es un **reporte HTML autocontenido** que compara N tenants
contra una configuración de referencia (*baseline*), resaltando diferencias por
workload, recurso e instancia, con búsqueda, filtros y exportación a CSV.

---

## Características

- **Un solo script para todo el flujo**, con menú de estado que marca cada paso
  como `[OK]` o `[  ]` y apunta al siguiente pendiente.
- **Sesiones aisladas**: cada paso operativo se lanza en un PowerShell nuevo y
  limpio para evitar los conflictos de ensamblados entre `Microsoft.Graph`, `Az`
  y `PnP.PowerShell` que sufre Microsoft365DSC.
- **App Registration por certificado** provisionada automáticamente, con
  cálculo de los permisos de Microsoft Graph / Exchange / SharePoint según los
  componentes que vayas a exportar, admin consent y roles de directorio.
- **Export unificado** que ejecuta los componentes generales y los de SharePoint
  (proceso aislado) y los fusiona en un único `M365TenantConfig.ps1`, avisando de
  los recursos de alto coste (recorrido sitio por sitio).
- **Reporte HTML comparativo** con neutralización de dominios/GUID entre tenants,
  descripciones y enlaces a documentación por recurso, y exportación a CSV.
- **Limpieza segura** de la App Registration, permisos, certificados y ficheros
  generados cuando terminas.

---

## Requisitos

- **Windows PowerShell 5.1** (el recomendado por Microsoft365DSC).
- Módulo **Microsoft365DSC** (el propio menú lo instala en el paso 1).
- Para provisionar la App Registration: cuenta **Global Administrator** del
  tenant.
- Conexión a Internet (descarga de módulos, admin consent y documentación).

---

## Inicio rápido

```powershell
# 1. Clona o descarga el repositorio en C:\M365DSC (ruta por defecto)
#    y sitúa Invoke-M365DSCReport.ps1 en C:\M365DSC\Scripts

# 2. Abre PowerShell (como Administrador para instalar módulos en AllUsers)
cd C:\M365DSC\Scripts
.\Invoke-M365DSCReport.ps1
```

Aparecerá el menú. Sigue el paso marcado como `<-- SIGUIENTE`.

---

## El menú

```
  PROCESO 1 - EXPORTAR LA CONFIGURACION DE CADA TENANT
   1) [  ] Preparar entorno (carpetas, modulo, dependencias)   <-- SIGUIENTE
   2) [  ] Generar consulta de export -> ConfigurationFile.ps1
   3) [  ] Provisionar App Registration (certificado)
   4) [  ] Exportar el tenant -> M365TenantConfig.ps1
   5) [  ] Eliminar App Registration (limpieza tras exportar)

  PROCESO 2 - REPORTE COMPARATIVO ENTRE TENANTS
   6) [  ] Verificar los M365TenantConfig.ps1 en Tenants\
   7) [  ] Generar el reporte HTML de baseline

   Q) Salir
```

| Paso | Qué hace |
|------|----------|
| **1. Preparar entorno** | Crea la estructura de carpetas, instala Microsoft365DSC y ejecuta `Update-M365DSCDependencies`. |
| **2. Generar consulta** | Te guía a [export.microsoft365dsc.com](https://export.microsoft365dsc.com/) para generar `ConfigurationFile.ps1` con los componentes a exportar. |
| **3. Provisionar App** | Crea la App Registration con certificado autofirmado, asigna los permisos según los componentes, concede admin consent y roles de directorio. Genera el script de export listo. *(Requiere Global Admin.)* |
| **4. Exportar tenant** | Ejecuta el export (SharePoint en proceso aislado) y fusiona todo en `M365TenantConfig.ps1`. Repite este bloque para cada tenant a comparar. |
| **5. Eliminar App** | Desmantela la App Registration, permisos, certificados y ficheros generados. |
| **6. Verificar tenants** | Comprueba que cada `M365TenantConfig.ps1` está en su carpeta bajo `Tenants\` y tiene datos. |
| **7. Generar reporte** | **Auto-detecta** los `M365TenantConfig.ps1` de `Tenants\` y te deja elegir cuáles comparar (o todas). Sugiere `Baseline` como *baseline* y guarda el HTML en `Reports\`. |

---

## Ejecutar un paso directamente (sin menú)

```powershell
.\Invoke-M365DSCReport.ps1 -Step Setup
.\Invoke-M365DSCReport.ps1 -Step Provision
.\Invoke-M365DSCReport.ps1 -Step Export
.\Invoke-M365DSCReport.ps1 -Step Report
.\Invoke-M365DSCReport.ps1 -Step Remove
```

Cambiar la carpeta raíz de trabajo (por defecto `C:\M365DSC`):

```powershell
.\Invoke-M365DSCReport.ps1 -Root "D:\M365DSC"
```

---

## Estructura de carpetas

```
C:\M365DSC\
├─ Scripts\
│  ├─ Invoke-M365DSCReport.ps1      # el orquestador (este script)
│  ├─ ConfigurationFile.ps1         # consulta generada en export.microsoft365dsc.com
│  ├─ catalog.json                  # descripciones y enlaces por recurso (reporte)
│  └─ logo.svg                      # tu logo embebido en el reporte (opcional)
├─ Export\                          # salida del export por tenant
├─ Tenants\
│  ├─ Baseline\  M365TenantConfig.ps1   # tenant de referencia (baseline)
│  ├─ SnapshotA\ M365TenantConfig.ps1
│  └─ SnapshotB\ M365TenantConfig.ps1
└─ Reports\                         # reportes HTML generados
```

---

## Flujo completo

1. **Por cada tenant** a comparar: pasos 1 → 4 (y 5 para limpiar al terminar).
   Copia cada `M365TenantConfig.ps1` resultante a su carpeta bajo `Tenants\`.
   El tenant más completo se usa como *baseline* (carpeta `Baseline`).
2. **Una vez** tengas 2 o más configuraciones: pasos 6 → 7 para generar el
   reporte comparativo.

> **Nota sobre SharePoint:** los recursos que recorren sitio por sitio
> (`SPOSite`, `SPOSiteGroup`, `SPOPropertyBag`, `SPOUserProfileProperty`, ...)
> pueden tardar horas en tenants grandes. El paso de export los detecta y ofrece
> excluirlos.

---

## Personalización (marca y logo)

El reporte **no lleva ninguna marca hardcodeada**. El paso 7 (Generar reporte)
**pregunta la marca por pantalla** cada vez, para que no se te olvide:

- **Nombre / organización** → aparece en el `<title>` y en el pie de página.
- **Eslogan** → aparece en la cabecera (se oculta si lo dejas en blanco).
- **Ruta del logo** → SVG/PNG/JPG/GIF/WEBP, se incrusta en el HTML. Si dejas un
  archivo `logo.svg` (o `logo.png`, `logo.jpg`) junto al script, se ofrece como
  valor por defecto.

Cualquier campo que dejes en blanco simplemente se omite (el reporte sale
genérico: "Microsoft365DSC Baseline Report", sin logo ni eslogan).

> El `logo.*` es personal: considera añadirlo a `.gitignore` si no quieres
> publicar tu marca en el repositorio (ya viene ignorado por defecto).

---

## Licencia

Distribuido bajo la **[PolyForm Noncommercial License 1.0.0](LICENSE)**.

Uso libre y colaborativo para fines **no comerciales**. El uso comercial no está
permitido bajo esta licencia. Consulta el archivo [`LICENSE`](LICENSE) para el
texto completo.

---

## Autor

**Christian Aguilera - FendariGroup**
```
Https://www.fendarigroup.com
