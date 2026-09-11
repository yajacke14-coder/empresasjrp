#!/usr/bin/env python3
"""
Procesador de Auditoría - Cruce Ventas vs SIIGO
================================================
Compara el Reporte de Ventas JRP contra los auxiliares contables de SIIGO
(cuenta 28050501) para las empresas Ayamonte, Salerno y Málaga.

Fórmula Ventas:  Ingreso Neto = CXPT (sin traslados) - CXC
Fórmula SIIGO:   Neto = Crédito - Débito

Uso:
    python procesador_auditoria.py

Requisitos:
    pip install openpyxl pandas
"""
import pandas as pd
import os
import sys
from datetime import datetime

VENTAS_FILE = os.path.join(os.path.dirname(__file__), '..', 'REPORTE VENTAS JRP.xlsx')
AUX_DIR = os.path.dirname(__file__)
OUTPUT_DIR = os.path.join(AUX_DIR, 'reportes')

EMPRESAS_ACTIVAS = [
    'AYAMONTE CONSTRUCCIONES SAS',
    'CONSTRUCTORA SALERNO SAS',
    'PROMOTORA MALAGA SAS',
]

SIIGO_FILES = {
    'AYAMONTE CONSTRUCCIONES SAS': [
        'AYA_2024_auxiliar.xlsx', 'AYA_2025_auxiliar.xlsx', 'AYA_2026_auxiliar.xlsx',
    ],
    'CONSTRUCTORA SALERNO SAS': [
        'SAL_2024_auxiliar.xlsx', 'SAL_2025_auxiliar.xlsx', 'SAL_2026_auxiliar.xlsx',
    ],
    'PROMOTORA MALAGA SAS': [
        'MLG_2024_auxiliar.xlsx', 'MLG_2025_auxiliar.xlsx', 'MLG_2026_auxiliar.xlsx',
    ],
}


def cargar_ventas(path=VENTAS_FILE):
    df = pd.read_excel(path)
    for col in df.columns:
        if col != 'TOTAL':
            df[col] = df[col].fillna('').astype(str)
    df['NIT TERCERO'] = df['NIT TERCERO'].apply(
        lambda x: x.split('.')[0] if '.' in str(x) else str(x)
    ).replace('', 'SIN NIT')
    df['TOTAL'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
    df = df[df['COMPRADOR'].isin(EMPRESAS_ACTIVAS)]
    return df


def cargar_siigo(path):
    try:
        test = pd.read_excel(path, header=None, nrows=15)
        header_row = None
        for i in range(min(15, len(test))):
            vals = [str(v).strip().lower() for v in test.iloc[i].values if pd.notna(v)]
            if any('contable' in v or 'tercero' in v for v in vals):
                header_row = i
                break
        s = pd.read_excel(path, header=header_row) if header_row is not None else pd.read_excel(path)
        if 'Código contable' in s.columns:
            s = s[s['Código contable'].notna()]
            s = s[~s['Código contable'].astype(str).str.contains('Cuenta contable', na=False)]
        return s
    except Exception as e:
        print(f"  Error cargando {path}: {e}")
        return pd.DataFrame()


def cargar_todos_siigo():
    parts = []
    for empresa, files in SIIGO_FILES.items():
        for fname in files:
            fpath = os.path.join(AUX_DIR, fname)
            if not os.path.exists(fpath):
                continue
            s = cargar_siigo(fpath)
            if len(s) > 0:
                s['_empresa'] = empresa
                s['_archivo'] = fname
                parts.append(s)
                print(f"  {fname}: {len(s)} registros")
    if not parts:
        return pd.DataFrame()
    df = pd.concat(parts, ignore_index=True)
    df['Identificación'] = df['Identificación'].fillna('').astype(str).apply(
        lambda x: x.split('.')[0]
    )
    df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
    df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
    df['Centro de costo'] = df['Centro de costo'].fillna('').astype(str)
    return df


def clasificar_transacciones(df):
    """
    CXPT real = consignaciones (sin traslados) -> ingreso
    Traslados = movimientos entre clientes, NO suman como ingreso nuevo
    CXC = devoluciones/desistimientos -> se restan del ingreso
    AP = abonos a proyecto
    """
    cxpt_real = df[(df['TRANSACCION'] == 'CXPT') & (df['INSUMO'] != 'Traslado')]
    cxpt_traslados = df[(df['TRANSACCION'] == 'CXPT') & (df['INSUMO'] == 'Traslado')]
    cxc = df[df['TRANSACCION'] == 'CXC']
    ap = df[df['TRANSACCION'] == 'AP']
    return cxpt_real, cxpt_traslados, cxc, ap


def safe_mode(series):
    m = series.mode()
    return str(m.iloc[0]) if len(m) > 0 else ''


def resumen_por_tercero(df):
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []
    for (tercero, nit), group in monetary.groupby(['TERCERO', 'NIT TERCERO']):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        inmuebles = [x for x in group['INMUEBLE'].unique() if x and x != 'SIN INMUEBLE']
        results.append({
            'TERCERO': tercero, 'NIT': nit,
            'EMPRESAS': ' | '.join(str(x) for x in group['COMPRADOR'].unique()),
            'PROYECTOS': ' | '.join(str(x) for x in group['PROYECTO'].unique() if x),
            'INMUEBLES': ' | '.join(str(x) for x in inmuebles[:5]),
            'CXPT_INGRESO_REAL': cxpt_real['TOTAL'].sum(),
            'CXPT_TRASLADOS': cxpt_traslados['TOTAL'].sum(),
            'CXC_DEVOLUCIONES': cxc['TOTAL'].sum(),
            'AP_ABONOS': ap['TOTAL'].sum(),
            'INGRESO_NETO': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
            'NUM_TRANSACCIONES': len(group),
        })
    return pd.DataFrame(results).sort_values('INGRESO_NETO', ascending=False)


def resumen_por_proyecto(df):
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []
    for proyecto, group in monetary.groupby('PROYECTO'):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        results.append({
            'PROYECTO': proyecto, 'EMPRESA': safe_mode(group['COMPRADOR']),
            'CXPT_INGRESO_REAL': cxpt_real['TOTAL'].sum(),
            'CXPT_TRASLADOS': cxpt_traslados['TOTAL'].sum(),
            'CXC_DEVOLUCIONES': cxc['TOTAL'].sum(),
            'AP_ABONOS': ap['TOTAL'].sum(),
            'INGRESO_NETO': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
            'NUM_TERCEROS': group['TERCERO'].nunique(),
            'NUM_INMUEBLES': group['INMUEBLE'].nunique(),
        })
    return pd.DataFrame(results).sort_values('INGRESO_NETO', ascending=False)


def resumen_por_empresa(df):
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []
    for empresa, group in monetary.groupby('COMPRADOR'):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        results.append({
            'EMPRESA': empresa,
            'CXPT_INGRESO_REAL': cxpt_real['TOTAL'].sum(),
            'CXPT_TRASLADOS': cxpt_traslados['TOTAL'].sum(),
            'CXC_DEVOLUCIONES': cxc['TOTAL'].sum(),
            'AP_ABONOS': ap['TOTAL'].sum(),
            'INGRESO_NETO': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
            'NUM_TERCEROS': group['TERCERO'].nunique(),
            'NUM_PROYECTOS': group['PROYECTO'].nunique(),
        })
    return pd.DataFrame(results).sort_values('INGRESO_NETO', ascending=False)


def cruce_ventas_siigo(ventas_df, siigo_df):
    """
    Cruza Ventas vs SIIGO por NIT + Empresa.
    Ventas: Ingreso Neto = CXPT real - CXC (traslados excluidos)
    SIIGO:  Neto = Crédito - Débito (cuenta 28050501)
    """
    monetary = ventas_df[ventas_df['TRANSACCION'] != 'SEG']

    ventas_grp = []
    for (nit, emp), g in monetary.groupby(['NIT TERCERO', 'COMPRADOR']):
        cxpt_real, cxpt_tras, cxc, _ = clasificar_transacciones(g)
        inmuebles = [x for x in g['INMUEBLE'].unique() if x and x != '' and x != 'SIN INMUEBLE']
        ventas_grp.append({
            'NIT': nit, 'EMPRESA': emp,
            'TERCERO': safe_mode(g['TERCERO']),
            'PROYECTOS': ' | '.join(str(x) for x in g['PROYECTO'].unique() if x),
            'INMUEBLES': ' | '.join(str(x) for x in inmuebles)[:200],
            'CXPT_REAL': float(cxpt_real['TOTAL'].sum()),
            'TRASLADOS': float(cxpt_tras['TOTAL'].sum()),
            'CXC': float(cxc['TOTAL'].sum()),
            'INGRESO_VENTAS': float(cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum()),
        })
    ventas_r = pd.DataFrame(ventas_grp)

    siigo_grp = []
    for (nit, emp), g in siigo_df.groupby(['Identificación', '_empresa']):
        siigo_grp.append({
            'NIT': nit, 'EMPRESA': emp,
            'TERCERO_SIIGO': safe_mode(g['Nombre del tercero']),
            'CENTRO_COSTO': ' | '.join(str(x) for x in g['Centro de costo'].unique() if x),
            'CREDITO_SIIGO': float(g['Crédito'].sum()),
            'DEBITO_SIIGO': float(g['Débito'].sum()),
            'NETO_SIIGO': float(g['Crédito'].sum() - g['Débito'].sum()),
        })
    siigo_r = pd.DataFrame(siigo_grp)

    cruce = ventas_r.merge(siigo_r, on=['NIT', 'EMPRESA'], how='outer', indicator=True)
    cruce['INGRESO_VENTAS'] = cruce['INGRESO_VENTAS'].fillna(0)
    cruce['NETO_SIIGO'] = cruce['NETO_SIIGO'].fillna(0)
    cruce['DIFERENCIA'] = cruce['INGRESO_VENTAS'] - cruce['NETO_SIIGO']
    cruce['ESTADO'] = cruce.apply(lambda r:
        'SOLO EN VENTAS' if r['_merge'] == 'left_only' else
        'SOLO EN SIIGO' if r['_merge'] == 'right_only' else
        'OK' if abs(r['DIFERENCIA']) <= 1 else
        'DIFERENCIA', axis=1
    )
    cruce.drop('_merge', axis=1, inplace=True)
    cruce['TERCERO'] = cruce['TERCERO'].fillna(cruce.get('TERCERO_SIIGO', ''))
    return cruce.sort_values('DIFERENCIA', key=abs, ascending=False)


def generar_reportes(ventas_path=VENTAS_FILE):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')

    print("Cargando reporte de ventas (AYA, SAL, MLG)...")
    df = cargar_ventas(ventas_path)
    print(f"  {len(df)} registros cargados")

    print("Cargando auxiliares SIIGO...")
    siigo_df = cargar_todos_siigo()
    print(f"  Total SIIGO: {len(siigo_df)} registros")

    print("Generando resúmenes...")
    tercero_df = resumen_por_tercero(df)
    proyecto_df = resumen_por_proyecto(df)
    empresa_df = resumen_por_empresa(df)

    output_file = os.path.join(OUTPUT_DIR, f'auditoria_ingresos_{timestamp}.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        tercero_df.to_excel(writer, sheet_name='Por_Tercero', index=False)
        proyecto_df.to_excel(writer, sheet_name='Por_Proyecto', index=False)
        empresa_df.to_excel(writer, sheet_name='Por_Empresa', index=False)

        if len(siigo_df) > 0:
            print("Generando cruce Ventas vs SIIGO...")
            cruce = cruce_ventas_siigo(df, siigo_df)
            cruce.to_excel(writer, sheet_name='Cruce_SIIGO', index=False)

            ok = len(cruce[cruce['ESTADO'] == 'OK'])
            dif = len(cruce[cruce['ESTADO'] == 'DIFERENCIA'])
            sv = len(cruce[cruce['ESTADO'] == 'SOLO EN VENTAS'])
            ss = len(cruce[cruce['ESTADO'] == 'SOLO EN SIIGO'])
            print(f"  OK: {ok} | Diferencia: {dif} | Solo Ventas: {sv} | Solo SIIGO: {ss}")

            diferencias = cruce[cruce['ESTADO'] != 'OK']
            diferencias.to_excel(writer, sheet_name='Diferencias', index=False)

    print(f"\nReporte generado: {output_file}")
    return output_file


if __name__ == '__main__':
    generar_reportes()
