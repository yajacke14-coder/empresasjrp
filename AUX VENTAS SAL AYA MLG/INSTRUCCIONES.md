# Auditoría de Ingresos Recibidos para Terceros - JRP

## Descripción

Sistema de auditoría para revisar los ingresos recibidos para terceros de las empresas a cargo:

- **AYAMONTE CONSTRUCCIONES SAS** (URB. Ayamonte, URB. Salamanca, Reserva Ayamonte, Balcones Valencia)
- **CONSTRUCTORA SALERNO SAS** (URB. Nápoles, Conjunto Sorrento)
- **PROMOTORA MALAGA SAS** (URB. Málaga 1, URB. Málaga 2, Portal Málaga, Portofino)
- **PROMOTORA RESERVAS DE ANDALUCIA SAS** (Senderos Andalucía)
- **GML GANADERIA SAS ZOMAC** (Ganadería, Proyecto Máquinas, Proyecto Prefabricados)
- **PROMOTORA SIERRA NEVADA SAS**

## Conceptos Clave

### Tipos de Transacción
| Código | Significado | Tratamiento |
|--------|-------------|-------------|
| **CXPT** | Cuentas por pagar terceros - Valor total consignado | Ingreso recibido |
| **CXC** | Cuentas por cobrar - Devoluciones/desistimientos | Se resta del ingreso |
| **AP** | Abonos a proyecto | Registro contable del proyecto |
| **SEG** | Seguimiento | Solo control, sin valor monetario |

### Traslados (CXPT con INSUMO = "Traslado")
Los traslados entre clientes **NO son ingresos nuevos**. Representan movimientos de dinero de un cliente a otro dentro del mismo proyecto. El sistema los identifica y los separa automáticamente del cálculo de ingreso real.

**Fórmula de Ingreso Neto:**
```
Ingreso Neto = CXPT (sin traslados) - CXC (devoluciones/desistimientos)
```

## Herramientas

### 1. Artefacto Web (Dashboard Interactivo)
Acceda al dashboard en el enlace proporcionado. Permite:
- **Resumen General**: Vista consolidada de todas las empresas
- **Por Tercero**: Búsqueda y filtrado por cliente, NIT, empresa o proyecto
- **Por Proyecto**: Agrupación por proyecto con totales
- **Por Empresa**: Desglose por cada empresa
- **Alertas**: Detección automática de anomalías
- **Detalle Cliente**: Drill-down completo de un tercero específico
- **Cruce SIIGO**: Carga de auxiliares para comparación

### 2. Script Python (`procesador_auditoria.py`)

**Requisitos:**
```bash
pip install openpyxl pandas
```

**Uso básico (solo reporte de ventas):**
```bash
python procesador_auditoria.py
```

**Con auxiliar SIIGO:**
```bash
python procesador_auditoria.py auxiliar_siigo.xlsx
```

O coloque el archivo como `auxiliar_siigo.xlsx` en esta misma carpeta.

**Salida:** Genera un archivo Excel en `reportes/auditoria_ingresos_FECHA.xlsx` con hojas:
- `Por_Tercero` - Resumen por cada tercero con alertas
- `Por_Proyecto` - Resumen por proyecto
- `Por_Empresa` - Resumen por empresa
- `Alertas` - Solo los terceros con alertas detectadas
- `Cruce_SIIGO` - Diferencias encontradas (si se cargó auxiliar)

## Alertas Automáticas

| Alerta | Condición | Significado |
|--------|-----------|-------------|
| **ALTO DESISTIMIENTO** | CXC > 50% del CXPT real | Devoluciones excesivas |
| **INGRESO NEGATIVO** | Ingreso neto < 0 | Más devoluciones que ingresos |
| **ALTO TRASLADO** | Traslados > 30% del CXPT | Movimientos inusuales entre clientes |
| **MULTI-EMPRESA** | Tercero en >1 empresa | Verificar que proyectos coincidan |

## Flujo de Actualización

1. Actualice el archivo `REPORTE VENTAS JRP.xlsx` en la raíz del repositorio
2. Ejecute el script o suba la información al dashboard
3. Si tiene auxiliares SIIGO nuevos, cárguelos en la sección "Cruce SIIGO" del dashboard o páselos como argumento al script
4. Revise las alertas y diferencias detectadas
5. Para detalle de un tercero específico, búsquelo en "Detalle Cliente"

## Estructura de Archivos

```
empresasjrp/
├── REPORTE VENTAS JRP.xlsx          # Archivo fuente principal
└── AUX VENTAS SAL AYA MLG/
    ├── INSTRUCCIONES.md             # Este archivo
    ├── procesador_auditoria.py      # Script de procesamiento
    ├── auxiliar_siigo.xlsx           # (Colocar aquí los auxiliares)
    └── reportes/                    # Reportes generados
        └── auditoria_ingresos_*.xlsx
```

## Datos del Reporte Actual

- **61,512** registros totales
- **1,883** terceros únicos
- **6** empresas
- **15** proyectos
- **16,915** transacciones CXPT
- **1,385** transacciones CXC (devoluciones)
- **730** traslados identificados (no contados como ingreso)
