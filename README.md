================================================================================
  M365 Security Assessment Toolkit
  Automated Microsoft 365 Security & Licensing Report
================================================================================

DESCRIPCION
-----------
Toolkit de PowerShell que genera un reporte HTML ejecutivo de seguridad y
licenciamiento para tenants de Microsoft 365. Diseñado para consultores y
equipos de seguridad que necesitan evaluar el estado de un tenant sin
realizar ningun cambio.

El reporte es un archivo HTML unico (single-file, sin dependencias externas)
con tema oscuro, graficas interactivas y datos exportables en JSON embebido.

100% READ-ONLY — No modifica, crea ni elimina nada en el tenant.


QUE ANALIZA
-----------
1. LICENCIAMIENTO
   - SKUs comprados vs asignados vs sin usar
   - Adoption por producto: Entra ID, MDE, MDO, MDA, MDI, Intune, Purview
   - Desperdicio: cuentas deshabilitadas con licencia, sin sign-in 90+ dias
   - Duplicados: licencias overlap (ej: E5 + standalone MDE)
   - Capacidad por departamento
   - Breakdown Members vs Guests

2. SEGURIDAD (ADOPTION REAL)
   - MFA: Capacidad de registro, metodos por tipo
   - Conditional Access: Politicas activas, cobertura
   - Entra ID P2: Risky Users, PIM, Access Reviews
   - MDE: Dispositivos onboarded, cobertura, alertas (via Advanced Hunting)
   - MDO: Emails procesados, phishing/malware bloqueado (via Advanced Hunting)
   - MDA: Apps cloud descubiertas, eventos (via Advanced Hunting)
   - MDI: Domain Controllers monitoreados, actividad (via Advanced Hunting)
   - Intune: Dispositivos enrolled, compliance
   - Copilot: Uso real vs licencias

3. SECURE SCORE
   - Score actual vs maximo por categoria
   - Top 20 recomendaciones ordenadas por impacto
   - Estado de implementacion por control


SCRIPTS
-------
Ejecutar en orden:

  Script 1: Get-MSLicensingReport.ps1    (~1160 lineas)
            Inventario de licencias, SKUs, usuarios, adoption, desperdicio
            Salida: report_data.json + 7 CSVs

  Script 2: Get-MSSecurityAdoption.ps1   (~1240 lineas)
            Medicion de uso real de productos de seguridad
            Salida: security_adoption.json

  Script 3: Get-MSSecureScore.ps1        (~370 lineas)
            Secure Score + recomendaciones
            Salida: secure_score.json

  Script 4: Generate-FullReport.ps1      (~1790 lineas)
            Genera reporte HTML unico combinando los 3 JSONs
            Salida: Full_Security_Report.html


REQUISITOS
----------
Sistema Operativo:
  - Windows 10/11, Windows Server 2019+
  - PowerShell 5.1+ o PowerShell 7+

Modulos de PowerShell:
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
  Install-Module Microsoft.Graph.Users -Scope CurrentUser

Permisos (Rol minimo):
  - Global Reader en el tenant de destino

Permisos de Microsoft Graph (delegados):
  - Directory.Read.All             (SKUs, usuarios)
  - Organization.Read.All          (datos del tenant)
  - User.Read.All                  (usuarios, sign-in)
  - Policy.Read.All                (Conditional Access)
  - UserAuthenticationMethod.Read.All  (MFA)
  - Reports.Read.All               (Copilot usage)
  - DeviceManagementManagedDevices.Read.All  (Intune)
  - RoleManagement.Read.All        (PIM)
  - IdentityRiskyUser.Read.All     (Risky Users)
  - ThreatHunting.Read.All         (Advanced Hunting: MDE/MDO/MDA/MDI)
  - SecurityEvents.Read.All        (Secure Score)


USO RAPIDO
----------
  cd C:\ruta\al\proyecto

  # Script 1 - Licenciamiento (obligatorio, siempre primero)
  .\Get-MSLicensingReport.ps1

  # Script 2 - Seguridad (requiere output del Script 1)
  .\Get-MSSecurityAdoption.ps1

  # Script 3 - Secure Score
  .\Get-MSSecureScore.ps1

  # Script 4 - Generar reporte HTML
  .\Generate-FullReport.ps1

Con TenantId especifico:
  .\Get-MSLicensingReport.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

Todos los archivos de salida se generan en .\output\ con timestamp.
El reporte HTML se abre automaticamente en el navegador.


SALIDA GENERADA
---------------
  output/
    YYYYMMDD_HHMM_report_data.json         <- Datos de licenciamiento
    YYYYMMDD_HHMM_security_adoption.json   <- Datos de adopcion de seguridad
    YYYYMMDD_HHMM_secure_score.json        <- Secure Score y recomendaciones
    YYYYMMDD_HHMM_Full_Security_Report.html <- Reporte HTML ejecutivo
    YYYYMMDD_HHMM_00_UnmappedPlans.csv     <- Plans no clasificados (referencia)
    YYYYMMDD_HHMM_01_SKUs.csv              <- Inventario de SKUs
    YYYYMMDD_HHMM_02_Users.csv             <- Detalle por usuario
    YYYYMMDD_HHMM_03_Adoption.csv          <- Adoption por categoria
    YYYYMMDD_HHMM_04_Waste.csv             <- Desperdicio detectado
    YYYYMMDD_HHMM_06_Departments.csv       <- Asignacion por departamento
    YYYYMMDD_HHMM_07_Capacity.csv          <- Capacidad disponible


ESCALABILIDAD
-------------
El toolkit esta optimizado para tenants de cualquier tamaño:

  1,000 usuarios   -> ~2 minutos, ~30 API calls
  50,000 usuarios  -> ~5 minutos, ~53 API calls
  100,000 usuarios -> ~10 minutos, ~53 API calls

El 95% del tiempo es descarga de datos via API.
El procesamiento local usa HashSet O(1) y es practicamente instantaneo.
Sin riesgo de throttling (usa <0.5% de los limites de Graph API).


SEGURIDAD
---------
- NO ejecuta comandos de escritura (Set-, New-, Remove-, Update-)
- NO modifica configuraciones, politicas ni usuarios
- NO almacena credenciales
- Las queries KQL de Advanced Hunting son consultas de lectura sobre
  telemetria existente
- Equivalente a abrir el portal de Azure y mirar dashboards
- Es seguro ejecutar en tenants de clientes en produccion


ESTRUCTURA DEL REPORTE HTML
----------------------------
El reporte HTML generado incluye:

  1. Dashboard        - KPIs principales con indicadores visuales
  2. Licensing        - SKUs, asignaciones, porcentajes de uso
  3. Adoption         - Uso real por categoria de seguridad con barras
  4. Waste            - Desperdicio detectado con recomendaciones
  5. Duplicates       - Licencias overlap entre SKUs
  6. MFA              - Estado de MFA, metodos de autenticacion
  7. Conditional Access - Politicas activas y configuracion
  8. Entra ID P2      - Risky Users, PIM, Access Reviews
  9. MDE              - Dispositivos, cobertura, alertas
  10. MDO             - Correos, phishing, malware, Safe Links
  11. MDA             - Apps cloud, eventos, shadow IT
  12. MDI             - Domain Controllers, actividad de identidad
  13. Intune          - Devices enrolled, compliance, plataformas
  14. Purview         - Information Protection, DLP, Audit
  15. Copilot         - Uso real vs licencias asignadas
  16. Secure Score    - Score por categoria, recomendaciones top
  17. Users           - Tabla interactiva con busqueda y filtros

Caracteristicas del HTML:
  - Single-file (sin dependencias externas, sin CDN)
  - Tema oscuro profesional con CSS variables
  - Graficas de barras HTML/CSS puro
  - Datos JSON embebidos para export
  - Navegacion por tabs
  - Tabla de usuarios con busqueda y paginacion JS
  - Optimizado para impresion (print styles)


MAPEO DE LICENCIAS
------------------
El script clasifica automaticamente los ServicePlans de Microsoft en categorias:

  ~95 ServicePlanNames mapeados a categorias de seguridad y productividad
  ~70 SkuPartNumbers con nombres amigables

Los planes no reconocidos se exportan a UnmappedPlans.csv para revision.
El mapeo usa NOMBRES de ServicePlan (no GUIDs), con descubrimiento dinamico
de IDs desde los SKUs reales del tenant.

Categorias cubiertas:
  Seguridad:     Entra ID P1/P2, MDE P1/P2, MDO P1/P2, MDA, MDI
  Compliance:    Purview AIP, MIP, DLP, Audit, eDiscovery, Insider Risk,
                 Encryption, Lockbox, PAM, Comm Compliance, Data Lifecycle
  Productividad: Exchange, SharePoint, Teams, M365 Apps, Power BI,
                 PowerApps, PowerAutomate, Copilot
  Gestion:       Intune P1/P2

NOTAS
-----
- Primera ejecucion: aparecera un popup de consentimiento de permisos de Graph.
  Es normal — son permisos de lectura delegados.
- El Script 2 detecta automaticamente que productos tiene licenciados el tenant
  y ejecuta solo los modulos relevantes.
- Si un producto no tiene licencia (ej: no hay MDE), esa seccion aparece vacia
  en el reporte con nota explicativa.
- Compatible con tenants GCC, Education y Commercial.
