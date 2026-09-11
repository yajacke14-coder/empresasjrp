#!/usr/bin/env python3
"""
Procesador de Auditoría de Ingresos JRP
========================================
Cruza el Reporte de Ventas JRP contra auxiliares contables de SIIGO.
Genera reportes agrupados por tercero, proyecto, empresa e inmueble.

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


def cargar_ventas(path=VENTAS_FILE):
    """Carga y limpia el reporte de ventas JRP."""
    df = pd.read_excel(path)
    df['NIT TERCERO'] = df['NIT TERCERO'].fillna(0).astype(int).astype(str).replace('0', 'SIN NIT')
    df['TOTAL'] = df['TOTAL'].fillna(0)
    df['COD FACTURA'] = df['COD FACTURA'].fillna('SIN CODIGO')
    df['INMUEBLE'] = df['INMUEBLE'].fillna('SIN INMUEBLE')
    df['INSUMO'] = df['INSUMO'].fillna('SIN INSUMO')
    return df


def cargar_auxiliar_siigo(path):
    """Carga un auxiliar contable de SIIGO (xlsx o csv)."""
    if path.endswith('.csv'):
        return pd.read_csv(path)
    return pd.read_excel(path)


def clasificar_transacciones(df):
    """
    Clasifica las transacciones separando:
    - CXPT: ingresos reales (excluyendo traslados)
    - CXPT Traslados: movimientos entre clientes (no son ingreso nuevo)
    - CXC: devoluciones y desistimientos
    - AP: abonos a proyecto
    """
    cxpt_real = df[(df['TRANSACCION'] == 'CXPT') & (df['INSUMO'] != 'Traslado')]
    cxpt_traslados = df[(df['TRANSACCION'] == 'CXPT') & (df['INSUMO'] == 'Traslado')]
    cxc = df[df['TRANSACCION'] == 'CXC']
    ap = df[df['TRANSACCION'] == 'AP']
    return cxpt_real, cxpt_traslados, cxc, ap


def resumen_por_tercero(df):
    """Genera resumen agrupado por tercero."""
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []

    for (tercero, nit), group in monetary.groupby(['TERCERO', 'NIT TERCERO']):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        results.append({
            'TERCERO': tercero,
            'NIT': nit,
            'EMPRESAS': ' | '.join(group['COMPRADOR'].unique()),
            'PROYECTOS': ' | '.join(group['PROYECTO'].unique()),
            'INMUEBLES': ' | '.join(group['INMUEBLE'].dropna().unique()[:5]),
            'CXPT_INGRESO_REAL': cxpt_real['TOTAL'].sum(),
            'CXPT_TRASLADOS': cxpt_traslados['TOTAL'].sum(),
            'CXC_DEVOLUCIONES': cxc['TOTAL'].sum(),
            'AP_ABONOS': ap['TOTAL'].sum(),
            'INGRESO_NETO': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
            'NUM_TRANSACCIONES': len(group),
            'ALERTA': ''
        })

    result_df = pd.DataFrame(results)

    # Alertas
    for idx, row in result_df.iterrows():
        alertas = []
        if row['CXC_DEVOLUCIONES'] > row['CXPT_INGRESO_REAL'] * 0.5 and row['CXPT_INGRESO_REAL'] > 0:
            alertas.append('ALTO DESISTIMIENTO')
        if row['INGRESO_NETO'] < 0:
            alertas.append('INGRESO NEGATIVO')
        if row['CXPT_TRASLADOS'] > row['CXPT_INGRESO_REAL'] * 0.3 and row['CXPT_TRASLADOS'] > 0:
            alertas.append('ALTO TRASLADO')
        if '|' in row['EMPRESAS']:
            alertas.append('MULTI-EMPRESA')
        result_df.at[idx, 'ALERTA'] = ' | '.join(alertas)

    return result_df.sort_values('INGRESO_NETO', ascending=False)


def resumen_por_proyecto(df):
    """Genera resumen agrupado por proyecto."""
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []

    for proyecto, group in monetary.groupby('PROYECTO'):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        results.append({
            'PROYECTO': proyecto,
            'EMPRESA': group['COMPRADOR'].mode().iloc[0] if len(group) > 0 else '',
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
    """Genera resumen agrupado por empresa."""
    monetary = df[df['TRANSACCION'] != 'SEG']
    results = []

    for empresa, group in monetary.groupby('COMPRADOR'):
        cxpt_real, cxpt_traslados, cxc, ap = clasificar_transacciones(group)
        results.append({
            'EMPRESA': empresa,
            'CXPT_TOTAL': (cxpt_real['TOTAL'].sum() + cxpt_traslados['TOTAL'].sum()),
            'CXPT_INGRESO_REAL': cxpt_real['TOTAL'].sum(),
            'CXPT_TRASLADOS': cxpt_traslados['TOTAL'].sum(),
            'CXC_DEVOLUCIONES': cxc['TOTAL'].sum(),
            'AP_ABONOS': ap['TOTAL'].sum(),
            'INGRESO_NETO': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
            'NUM_TERCEROS': group['TERCERO'].nunique(),
            'NUM_PROYECTOS': group['PROYECTO'].nunique(),
        })

    return pd.DataFrame(results).sort_values('INGRESO_NETO', ascending=False)


def detalle_tercero(df, tercero_nombre):
    """Obtiene el detalle completo de un tercero específico."""
    mask = df['TERCERO'].str.contains(tercero_nombre, case=False, na=False)
    registros = df[mask & (df['TRANSACCION'].isin(['CXPT', 'CXC']))].copy()
    registros['ES_TRASLADO'] = registros['INSUMO'] == 'Traslado'
    return registros.sort_values(['TRANSACCION', 'COD FACTURA'])


def cruzar_con_siigo(ventas_df, siigo_df, columna_nit_siigo='NIT', columna_valor_siigo='DEBITO'):
    """
    Cruza el reporte de ventas contra el auxiliar de SIIGO.
    Identifica diferencias por tercero.
    """
    # Resumen ventas por NIT
    ventas_monetary = ventas_df[ventas_df['TRANSACCION'] != 'SEG']
    ventas_resumen = []
    for nit, group in ventas_monetary.groupby('NIT TERCERO'):
        cxpt_real, _, cxc, _ = clasificar_transacciones(group)
        ventas_resumen.append({
            'NIT': nit,
            'TERCERO_VENTAS': group['TERCERO'].mode().iloc[0] if len(group) > 0 else '',
            'INGRESO_VENTAS': cxpt_real['TOTAL'].sum() - cxc['TOTAL'].sum(),
        })
    ventas_r = pd.DataFrame(ventas_resumen)

    # Resumen SIIGO por NIT
    siigo_df[columna_nit_siigo] = siigo_df[columna_nit_siigo].astype(str)
    siigo_r = siigo_df.groupby(columna_nit_siigo).agg(
        INGRESO_SIIGO=(columna_valor_siigo, 'sum')
    ).reset_index()
    siigo_r.rename(columns={columna_nit_siigo: 'NIT'}, inplace=True)

    # Cruce
    cruce = ventas_r.merge(siigo_r, on='NIT', how='outer', indicator=True)
    cruce['DIFERENCIA'] = cruce['INGRESO_VENTAS'].fillna(0) - cruce['INGRESO_SIIGO'].fillna(0)
    cruce['ALERTA_CRUCE'] = cruce.apply(lambda r:
        'SOLO EN VENTAS' if r['_merge'] == 'left_only' else
        'SOLO EN SIIGO' if r['_merge'] == 'right_only' else
        'DIFERENCIA' if abs(r['DIFERENCIA']) > 1 else
        'OK', axis=1
    )
    cruce.drop('_merge', axis=1, inplace=True)
    return cruce.sort_values('DIFERENCIA', key=abs, ascending=False)


def generar_reportes(ventas_path=VENTAS_FILE, siigo_path=None):
    """Genera todos los reportes y los guarda en Excel."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')

    print("Cargando reporte de ventas...")
    df = cargar_ventas(ventas_path)
    print(f"  {len(df)} registros cargados")

    print("Generando resumen por tercero...")
    tercero_df = resumen_por_tercero(df)

    print("Generando resumen por proyecto...")
    proyecto_df = resumen_por_proyecto(df)

    print("Generando resumen por empresa...")
    empresa_df = resumen_por_empresa(df)

    output_file = os.path.join(OUTPUT_DIR, f'auditoria_ingresos_{timestamp}.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        tercero_df.to_excel(writer, sheet_name='Por_Tercero', index=False)
        proyecto_df.to_excel(writer, sheet_name='Por_Proyecto', index=False)
        empresa_df.to_excel(writer, sheet_name='Por_Empresa', index=False)

        alertas = tercero_df[tercero_df['ALERTA'] != '']
        alertas.to_excel(writer, sheet_name='Alertas', index=False)

        if siigo_path and os.path.exists(siigo_path):
            print("Cargando auxiliar SIIGO...")
            siigo_df = cargar_auxiliar_siigo(siigo_path)
            print(f"  {len(siigo_df)} registros SIIGO cargados")

            nit_col = next((c for c in siigo_df.columns if 'nit' in c.lower()), None)
            val_col = next((c for c in siigo_df.columns if c.lower() in ['debito', 'valor', 'monto']), None)

            if nit_col and val_col:
                print(f"  Cruzando por columnas: NIT={nit_col}, VALOR={val_col}")
                cruce = cruzar_con_siigo(df, siigo_df, nit_col, val_col)
                cruce.to_excel(writer, sheet_name='Cruce_SIIGO', index=False)
                print(f"  {len(cruce[cruce['ALERTA_CRUCE'] != 'OK'])} diferencias encontradas")
            else:
                print(f"  ADVERTENCIA: No se encontraron columnas NIT/VALOR en SIIGO")
                print(f"  Columnas disponibles: {list(siigo_df.columns)}")

    print(f"\nReporte generado: {output_file}")
    return output_file


if __name__ == '__main__':
    siigo = None
    if len(sys.argv) > 1:
        siigo = sys.argv[1]
    elif os.path.exists(os.path.join(AUX_DIR, 'auxiliar_siigo.xlsx')):
        siigo = os.path.join(AUX_DIR, 'auxiliar_siigo.xlsx')

    generar_reportes(siigo_path=siigo)
