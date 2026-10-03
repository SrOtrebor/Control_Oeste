# Plan de Arquitectura y Migración v2
# Plataforma Multi-Sede de Control de Acceso y Fichaje Biométrico

**Proyecto**: Control de Acceso & Tiempo y Asistencia  
**Última actualización**: 09/09/2026  
**Versión**: 2.0  
**Premisa**: El sistema actual en Python queda 100% intacto y operativo como respaldo.

---

## 1. Arquitectura Híbrida (Python + Power Apps + SharePoint)

```
                    ┌─────────────────────────┐
                    │   SHAREPOINT ONLINE      │
                    │   (Base de Datos Central) │
                    └──────────┬──────────────┘
                               │
                ┌──────────────┼──────────────┐
                │              │              │
                ▼              ▼              ▼
        ┌──────────┐   ┌───────────┐   ┌───────────┐
        │ App Python│   │ Power Apps │   │ Power BI  │
        │ (Garita)  │   │ (Consola) │   │(Dashboard)│
        └──────────┘   └───────────┘   └───────────┘
```

- **Python (garita)**: Velocidad instantánea, funciona offline, integra biometría USB.
- **Power Apps (oficina)**: Consola Master, gestión de nóminas, alertas, monitoreo multi-sede.
- **SharePoint**: Base de datos compartida, cifrada, auditada, gestionada por IT.

---

## 2. Stack Tecnológico

| Componente | Tecnología |
| :--- | :--- |
| Frontend Garita | Python (Flask + Pywebview) — actual, mejorado |
| Consola Master | Power Apps (Canvas App) |
| Base de Datos | Listas de SharePoint Online (Graph API) |
| Ingesta IRSA | **Python + API REST directa** (sin scraping) |
| Fichaje Biométrico | SDK del lector USB + Python (sin bridge) |
| Orquestación | Power Automate (alertas, archivado, reportes) |
| Identidad | Microsoft Entra ID (Azure AD) |

---

## 3. API de IRSA Conexión Locatarios (Descubierta 09/09/2026)

> **HALLAZGO CRÍTICO**: El portal `locatarios.irsa.com.ar` expone una API REST
> completa en JSON. No se requiere scraping, RPA ni robots de navegación.
> Toda la información se obtiene mediante llamadas HTTP directas.

### 3.1 Autenticación

- **Método**: Cookie de sesión (`ASP.NET_SessionId` + cookie `Locatarios`).
- **Servidor**: Microsoft IIS/10.0 + ASP.NET.
- **Obtención de sesión**: Login estándar con credenciales corporativas del usuario autorizado.

### 3.2 Endpoints Descubiertos

#### ENDPOINT 1 — Listar FAOs

```
POST https://locatarios.irsa.com.ar/api/faos/listar
Content-Type: application/json

Body:
{
  "start": 0,                    // Offset de paginación
  "limit": 500,                  // Máx. resultados (default del portal: 30)
  "shopping": ["61709"],         // Array de IDs de shopping (multi-sede)
  "desde": "2026-07-09",         // Fecha desde
  "hasta": "",                   // Fecha hasta
  "personal": "",                // Buscar por DNI directamente
  "marca": "",                   // Filtrar por empresa
  "cuit": "",                    // Filtrar por CUIT
  "id": "",                      // Filtrar por Nro. de FAO
  "estado": null,                // Filtrar por estado
  "estadoMensaje": "-1",
  "estados": null,
  "local": "",
  "porAprobar": false
}

Response:
{
  "results": [
    {
      "id": 1264326,
      "marca": "ASCENSORES SCHINDLER S.A.",
      "idShopping": "61709",
      "shopping": "Al Oeste Shopping",
      "fechaInicio": "2026-09-09T00:00:00",
      "fechaFin": "2026-09-19T00:00:00",
      "usuario": "30654687591/ASCENSORESSCHINDLERS.A.",
      "estado": "4",
      "areaSolicitante": "TIS",
      "aprobacionAutomatica": "NO"
    },
    // ... más FAOs
  ],
  "totalCount": 86
}
```

**IDs de Shopping conocidos:**
- `61709` = Al Oeste Shopping / Oeste Outlet

**Códigos de estado:**
- `0` = Ingresado (recién creado)
- `3` = En trámite (parcialmente aprobado)
- `4` = Aprobado
- `6` = Rechazado (probable, a confirmar)

---

#### ENDPOINT 2 — Detalle de FAO (Personal + Documentos)

```
GET https://locatarios.irsa.com.ar/api/faos/editar/{id}

Response completa:
{
  "id": "1263905",
  "marca": "ASCENSORES SCHINDLER S.A.",
  "emprendimiento": "30654687591",           // CUIT de la empresa
  "idShopping": "61709",
  "shopping": "Oeste Outlet",
  "fechaInicio": "2026-09-09T00:00:00",      // Vigencia desde
  "fechaFin": "2026-09-19T00:00:00",         // Vigencia hasta
  "lugarTrabajo": "Al Oeste Shopping",
  "detalleTrabajo": "Modernización de escaleras",
  "estado": "4",
  "horaInicio": "07:00:00",
  "horaFin": "10:00:00",
  "esHorarioNocturno": false,
  "acompaniamientoSeguridad": false,
  "ingresoCondicional": false,

  "personal": [                              // ← PERSONAS AUTORIZADAS
    {
      "id": 2698447,
      "nombre": "SERGIO",
      "apellido": "EHRAIJE",
      "tipoDocumento": "D.N.I.",
      "numeroDocumento": "24288384",          // ← DNI LIMPIO
      "nombreART": "EXPERTA ART",
      "telefonoART": "0800-777-7278",
      "activo": true
    }
    // ... más personas
  ],

  "documentosAdjuntos": [                    // ← PDFs ADJUNTOS
    {
      "id": 9147065,
      "nombre": "Poliza RC - ASSA Elevators(1).pdf",
      "ruta": "1263905\\20260908...guid.pdf"  // ← Ruta de descarga
    }
    // ... más documentos
  ],

  "firmas": [                               // ← ESTADO DE APROBACIONES
    {
      "roleName": "FAO Aprobacion Comercial",
      "userFullName": "Franco Daniel Curti",
      "vigente": true,
      "fechaFirma": "2026-09-08T09:46:31.063"
    }
  ],

  "rolesFirmantesPendientes": [              // ← FIRMAS QUE FALTAN
    {
      "roleName": "FAO Aprobacion TIS"
    }
  ],

  "comentarios": [],
  "tiposTrabajo": [119, 120]
}
```

---

#### ENDPOINT 3 — Descarga de PDF adjunto

```
GET https://locatarios.irsa.com.ar/api/faos/descargar/documento?ruta={ruta_url_encoded}

Ejemplo:
GET /api/faos/descargar/documento?ruta=1263905%5C20260908...guid.pdf

Response: Archivo PDF binario
```

---

#### ENDPOINT 4 — Listar FAPs (equivalente a FAOs)

```
POST https://locatarios.irsa.com.ar/api/faps/listar
Content-Type: application/json

Body: (misma estructura que FAOs)
{
  "start": 0,
  "limit": 500,
  "shopping": ["61709"],
  "desde": "...",
  "hasta": "...",
  // ... mismos filtros
}
```

> **CONFIRMACIÓN (09/09/2026)**: Se verificaron en vivo los endpoints de FAP.
> La estructura de respuesta de `/api/faps/editar/{id}` es **exactamente idéntica** a la de FAOs, compartiendo los mismos nodos `personal` (con DNIs) y `documentosAdjuntos`. 
> Esto permite utilizar el mismo código de ingesta para ambos tipos de trámites.

---

### 3.3 Estrategia de Ingesta Automatizada

```
┌──────────────────────────────────────────────────────────┐
│              FLUJO DE INGESTA IRSA                       │
│                                                          │
│  1. Autenticarse → obtener cookie de sesión              │
│                                                          │
│  2. POST /api/faos/listar (limit=500, shopping=sede)     │
│     POST /api/faps/listar (limit=500, shopping=sede)     │
│     → Obtener IDs de todos los FAOs/FAPs vigentes        │
│                                                          │
│  3. Para cada FAO/FAP:                                   │
│     GET /api/faos/editar/{id}                            │
│     → Extraer personal[] (DNI, nombre, apellido, ART)    │
│     → Registrar en Personas_Autorizadas                  │
│                                                          │
│  4. Filtro inteligente de documentosAdjuntos[]:           │
│     ⛔ Ignorar: *F931*, *Libre Deuda*, *DDJJ*, *ARCA*    │
│     ✅ Parsear: *Poliza*, *Cobertura*, *SVO*, *CNR*      │
│     → Extraer DNIs adicionales de pólizas ART            │
│     → Agregar a Personas_Autorizadas como origen PDF     │
│                                                          │
│  5. Log de ejecución en Log_Ingesta_IRSA                 │
│                                                          │
│  Frecuencia: Programable (ej. 6:00 AM y 13:00 PM)       │
│  Tiempo estimado: < 1 minuto por shopping                │
└──────────────────────────────────────────────────────────┘
```

#### Filtro de PDFs por nombre de archivo

| Patrón en el nombre | Acción | Motivo |
|:---------------------|:-------|:-------|
| `*F931*` | ⛔ Ignorar | Formulario impositivo AFIP, no tiene DNIs útiles |
| `*Libre Deuda*` | ⛔ Ignorar | Certificado de libre deuda, sin nómina |
| `*DDJJ*`, `*ARCA*`, `*Pago*` | ⛔ Ignorar | Comprobantes impositivos |
| `*Poliza*`, `*RC*` | ✅ Descargar y parsear | Póliza de responsabilidad civil, puede listar personas |
| `*Constancia Cobertura*`, `*SVO*` | ✅ Descargar y parsear | Constancia de seguro con nómina de trabajadores |
| `*CNR*`, `*Clausula*` | 🟡 Revisar | Cláusula de No Repetición, puede tener datos relevantes |
| `*asociart*`, `*experta*`, `*galeno*`, `*prevencion*` | ✅ Descargar y parsear | Documentos de ART con nóminas |

---

## 4. Modelo de Datos (Listas de SharePoint)

> Todas las listas llevan `Sede_ID` indexado. `DNI` indexado donde aplique.
> Archivado mensual automático para listas transaccionales.

### Listas Maestras
- `Sedes` — Config de cada shopping (horarios diurno/nocturno, activo)
- `Empresas_Proveedores` — Nombre normalizado, CUIT, rubro
- `Puestos_Objetivos` — Puestos de fichaje por sede
- `Roles_Usuarios` — Admin Master / Admin Sede / Operador

### Listas Operativas
- `Personas_Autorizadas` — Nómina maestra unificada (DNI, tipo permiso, vigencia, origen)
- `Registros_Accesos` — Movimientos de garita (entrada/salida consolidados)
- `Fichajes_Personal` — Fichadas con cálculo de horas diurnas/nocturnas
- `Alertas_Seguimiento` — Personas de interés con acción y notificación

### Listas de Auditoría
- `Log_Ingesta_IRSA` — Registro de cada sincronización
- `Huellas_Biometricas` — Templates biométricos (Fase 5)

---

## 5. Plan de Ejecución por Fases

```
FASE 1 — Cimientos de Datos (SharePoint)
├── Crear sitio y listas con índices
├── Dar de alta sede ALOE (Al Oeste Shopping, idShopping: 61709)
├── Flujo de archivado mensual
└── Entregable: Estructura lista para consumir

FASE 2 — App de Garita (MVP Python → SharePoint)
├── Conectar app Python actual a SharePoint vía Graph API
├── Caché local en RAM para velocidad instantánea
├── Modo offline con cola de sincronización
├── Validación DNI contra Personas_Autorizadas
├── Feedback visual (verde/rojo/gris/amarillo FAO)
└── Entregable: Garita operando con base centralizada

FASE 3 — Consola Master (Power Apps)
├── Gestión de nóminas (carga manual + parser ART)
├── ABM de excepciones
├── Dashboard en tiempo real
├── Selector multi-sede para Admin Master
├── Exportación y reportes
└── Entregable: Administración centralizada desde oficina

FASE 4 — Ingesta Automatizada IRSA (Python + API REST)
├── Módulo de autenticación y sesión
├── Sincronización de FAOs: /api/faos/listar + /editar/{id}
├── Sincronización de FAPs: /api/faps/listar + /editar/{id}
├── Filtro y parseo de PDFs de ART (solo los relevantes)
├── Inyección a Personas_Autorizadas
├── Log de auditoría en Log_Ingesta_IRSA
├── Programación horaria (6 AM / 1 PM) + botón manual
└── Entregable: Permisos actualizados sin intervención humana

FASE 5 — Fichaje Biométrico y Cálculo Horario
├── POC con lector de muestra (ZKTeco/SecuGen/DigitalPersona)
├── Integración SDK USB directo con Python
├── Popup de selección de puesto
├── Motor de cálculo diurno/nocturno
├── Reporte mensual para liquidación
├── Si POC falla: Plan B (PIN + DNI)
└── Entregable: Sistema de fichaje operativo

FASE 6 — Multi-Sede, Alertas y Cierre
├── Alertas con notificación (Teams/Email)
├── Alta de segunda sede de prueba
├── Validación de aislamiento de datos
├── Homologación en Al Oeste Shopping
├── Capacitación de operadores
└── Entregable: Plataforma multi-sede en producción
```

---

## 6. Criterios de Go/No-Go (Producción)

| Criterio | Go ✅ | No-Go 🛑 |
|:---------|:------|:---------|
| Respuesta en garita | < 1.5s promedio (100 escaneos) | > 3s o > 5% timeouts |
| Precisión | 100% match con nóminas | Cualquier falso negativo |
| Offline | Opera con caché, sincroniza al volver | Se congela sin internet |
| Multi-sede | Sede A no ve datos de Sede B | Cualquier fuga |
| Ingesta IRSA | Extrae 100% del personal[] de los FAOs/FAPs | Pierde registros |

---

## 7. Contingencia

- Sistema Python actual intacto como respaldo inmediato.
- Si la ingesta IRSA falla → carga manual mejorada vía Consola Master.
- Si biometría no es viable → PIN + DNI (sin impacto en arquitectura).
- Si Power Apps no rinde → backend Python con base SharePoint híbrida.
