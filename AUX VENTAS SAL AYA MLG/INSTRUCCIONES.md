# Auditoría: Cruce Ventas vs SIIGO - JRP

## Objetivo

Verificar que los ingresos recibidos para terceros en el **Reporte de Ventas** coincidan con los registros del auxiliar contable **SIIGO** (cuenta 28050501 "De clientes").

## Empresas Activas

| Empresa | Prefijo SIIGO | Proyectos |
|---------|---------------|-----------|
| **AYAMONTE CONSTRUCCIONES SAS** | AYA | URB. Ayamonte, Reserva Ayamonte, URB. Salamanca, Balcones Valencia |
| **CONSTRUCTORA SALERNO SAS** | SAL | URB. Nápoles, Conjunto Sorrento |
| **PROMOTORA MALAGA SAS** | MLG | URB. Málaga 1, URB. Málaga 2, Portal Málaga, Portofino, Proyecto Málaga |

> Las demás empresas (GML Ganadería, Reservas Andalucía, Sierra Nevada) se agregarán conforme se obtengan sus auxiliares SIIGO.

## Lógica del Cálculo

### Reporte de Ventas

| Tipo | Significado | Tratamiento |
|------|-------------|-------------|
| **CXPT** | Consignación del tercero | Se suma como ingreso |
| **CXPT con INSUMO = "Traslado"** | Movimiento entre clientes | **NO suma** como ingreso nuevo |
| **CXC** | Devolución / desistimiento | **Se resta** del ingreso |
| **AP** | Abono a proyecto | Registro contable |
| **SEG** | Seguimiento | Sin valor monetario |

```
Ingreso Neto Ventas = CXPT (sin traslados) - CXC
```

**Ejemplo:** Si José Hilario aportó $100M (CXPT), se le devolvió $10M (CXC) y se trasladó $5M a otro cliente (Traslado):
- CXPT real = $100M (los $5M de traslado NO se suman)
- CXC = $10M
- Ingreso Neto = $100M - $10M = **$90M**

### Auxiliar SIIGO (Cuenta 28050501)

- **Crédito** = pagos recibidos del tercero
- **Débito** = devoluciones / ajustes
- **Centro de costo** = equivale al proyecto en Ventas

```
Neto SIIGO = Crédito - Débito
```

### Cruce

Se compara por **NIT + Empresa**:
```
Diferencia = Ingreso Neto Ventas - Neto SIIGO
```

| Estado | Significado |
|--------|-------------|
| **OK** | Diferencia ≤ $1 (coincide) |
| **DIFERENCIA** | Los valores no coinciden, revisar |
| **SOLO EN VENTAS** | El tercero aparece en Ventas pero no en SIIGO |
| **SOLO EN SIIGO** | El tercero aparece en SIIGO pero no en Ventas |

## Cómo Actualizar la Información

### Paso 1: Actualizar el Reporte de Ventas

Reemplace el archivo `REPORTE VENTAS JRP.xlsx` en la raíz del repositorio con la versión más reciente.

### Paso 2: Actualizar Auxiliares SIIGO

Coloque los nuevos auxiliares en esta carpeta (`AUX VENTAS SAL AYA MLG/`) con la nomenclatura:

```
{PREFIJO}_{AÑO}_auxiliar.xlsx
```

Ejemplos:
- `AYA_2024_auxiliar.xlsx` (Ayamonte 2024)
- `SAL_2025_auxiliar.xlsx` (Salerno 2025)
- `MLG_2026_auxiliar.xlsx` (Málaga 2026)

Si tiene un nuevo año o empresa, agregue el archivo y actualice la lista `SIIGO_FILES` en `procesador_auditoria.py`.

### Paso 3: Generar el Reporte Excel

```bash
cd "AUX VENTAS SAL AYA MLG"
pip install openpyxl pandas   # solo la primera vez
python procesador_auditoria.py
```

El reporte se genera en `reportes/auditoria_ingresos_FECHA.xlsx` con hojas:
- `Por_Tercero` - Resumen por cada tercero con inmuebles
- `Por_Proyecto` - Resumen por proyecto
- `Por_Empresa` - Resumen por empresa
- `Cruce_SIIGO` - Comparación completa Ventas vs SIIGO
- `Diferencias` - Solo los terceros donde no coincide

### Paso 4: Actualizar el Dashboard Web

Abra una sesión de Claude Code y pídale:
> "Actualiza el dashboard de auditoría con los datos nuevos"

Claude leerá los archivos actualizados y regenerará el artefacto web.

## Cómo Agregar una Nueva Empresa

1. Obtenga el auxiliar SIIGO de la nueva empresa
2. Guárdelo como `{PREFIJO}_{AÑO}_auxiliar.xlsx` en esta carpeta
3. Edite `procesador_auditoria.py`:
   - Agregue el nombre de la empresa a `EMPRESAS_ACTIVAS`
   - Agregue los archivos a `SIIGO_FILES`
4. Ejecute el script para verificar

## Estructura de Archivos

```
empresasjrp/
├── REPORTE VENTAS JRP.xlsx              # Fuente principal
└── AUX VENTAS SAL AYA MLG/
    ├── INSTRUCCIONES.md                 # Este archivo
    ├── procesador_auditoria.py          # Script de procesamiento
    ├── AYA_2024_auxiliar.xlsx           # Auxiliar SIIGO Ayamonte 2024
    ├── AYA_2025_auxiliar.xlsx           # Auxiliar SIIGO Ayamonte 2025
    ├── AYA_2026_auxiliar.xlsx           # Auxiliar SIIGO Ayamonte 2026
    ├── SAL_2024_auxiliar.xlsx           # Auxiliar SIIGO Salerno 2024
    ├── SAL_2025_auxiliar.xlsx           # Auxiliar SIIGO Salerno 2025
    ├── SAL_2026_auxiliar.xlsx           # Auxiliar SIIGO Salerno 2026
    ├── MLG_2024_auxiliar.xlsx           # Auxiliar SIIGO Málaga 2024
    ├── MLG_2025_auxiliar.xlsx           # Auxiliar SIIGO Málaga 2025
    ├── MLG_2026_auxiliar.xlsx           # Auxiliar SIIGO Málaga 2026
    └── reportes/                        # Reportes generados
        └── auditoria_ingresos_*.xlsx
```

## Datos del Reporte Actual

- **61,170** registros de ventas (3 empresas)
- **17,531** registros SIIGO
- **1,854** terceros únicos
- **15** proyectos
- **Cruce:** 672 OK | 986 Diferencias | 76 Solo Ventas | 164 Solo SIIGO
