"""
main.py - Orquestador Integral del Sistema de Ensamblaje de Especificaciones Técnicas

Permite ejecutar el flujo completo en un solo comando:
1. Fase 1: División del Maestro Word, almacenamiento en /partidas_db/ e indexación SQLite.
2. Fase 2: Lectura de la plantilla Excel (.xlsx) y extracción en memoria de partidas solicitadas.
3. Fase 3: Ensamblaje ordenado en documento final Word (.docx) preservando tablas e imágenes.
"""

import os
import sys
import argparse
import time

from divisor import ejecutar_division, DEFAULT_REGEX
from lector_excel import leer_solicitud_excel
from ensamblador import ensamblar_expediente


def pipeline_completo(
    maestro_path: str,
    excel_path: str,
    salida_path: str = "Expediente_Final.docx",
    plantilla_path: str = None,
    output_dir: str = "partidas_db",
    db_path: str = "partidas.db",
    columna_excel: str = "Código de Partida",
    regex_pattern: str = DEFAULT_REGEX,
    backend: str = "auto",
    reset_db: bool = False
):
    inicio = time.time()
    print("=" * 65)
    print(" SISTEMA AUTOMATIZADO DE ENSAMBLAJE DE ESPECIFICACIONES TÉCNICAS")
    print("=" * 65)
    
    # 1. Fase 1: División
    print("\n>>> INICIANDO FASE 1: DIVISIÓN Y ALMACENAMIENTO...")
    partidas = ejecutar_division(
        master_path=maestro_path,
        db_path=db_path,
        output_dir=output_dir,
        regex_pattern=regex_pattern,
        backend=backend,
        reset_db=reset_db
    )
    if not partidas:
        print("[ERROR FATAL] La Fase 1 no pudo extraer ninguna partida. Abortando.")
        return False
        
    # 2. Fase 2: Lectura de Solicitud Excel
    print("\n>>> INICIANDO FASE 2: INTERFAZ DE SOLICITUD (EXCEL)...")
    codigos_solicitados = leer_solicitud_excel(
        excel_path=excel_path,
        columna=columna_excel
    )
    print(f"✔ Se detectaron {len(codigos_solicitados)} códigos solicitados en Excel.")
    
    # 3. Fase 3: Ensamblaje
    print("\n>>> INICIANDO FASE 3: ENSAMBLAJE DEL EXPEDIENTE CONSOLIDADO...")
    res = ensamblar_expediente(
        solicitud=codigos_solicitados,
        db_path=db_path,
        salida_path=salida_path,
        plantilla_path=plantilla_path,
        backend=backend
    )
    
    tiempo_total = time.time() - inicio
    print("=" * 65)
    if res.get("exito"):
        print(f"✔ PIPELINE COMPLETADO CON ÉXITO EN {tiempo_total:.2f} SEGUNDOS")
        print(f"  Expediente generado: {os.path.abspath(salida_path)}")
    else:
        print(f"⚠ PIPELINE FINALIZÓ CON ADVERTENCIAS EN {tiempo_total:.2f} SEGUNDOS")
    print("=" * 65)
    return res.get("exito", False)


def main():
    parser = argparse.ArgumentParser(
        description="Sistema de Ensamblaje Automatizado de Expedientes Técnicos (Fases 1, 2 y 3)."
    )
    parser.add_argument(
        "-m", "--maestro", required=True,
        help="Ruta al documento Word Maestro con todas las partidas (.docx)."
    )
    parser.add_argument(
        "-e", "--excel", required=True,
        help="Ruta a la plantilla Excel con los códigos de partida requeridos (.xlsx)."
    )
    parser.add_argument(
        "-o", "--salida", default="Expediente_Consolidado.docx",
        help="Ruta del documento Word consolidado final (default: Expediente_Consolidado.docx)."
    )
    parser.add_argument(
        "-p", "--plantilla", default=None,
        help="Plantilla Word base opcional (para carátula, logos, márgenes y membretes)."
    )
    parser.add_argument(
        "--db", default="partidas.db",
        help="Ruta de la base de datos SQLite (default: partidas.db)."
    )
    parser.add_argument(
        "--output-dir", default="partidas_db",
        help="Directorio local para las partidas divididas (default: partidas_db)."
    )
    parser.add_argument(
        "--columna", default="Código de Partida",
        help="Nombre de la columna en el archivo Excel (default: 'Código de Partida')."
    )
    parser.add_argument(
        "--regex", default=DEFAULT_REGEX,
        help="Expresión regular para los títulos de partida."
    )
    parser.add_argument(
        "--backend", choices=["auto", "docx", "docxcompose", "win32com"], default="auto",
        help="Motor de procesamiento (default: 'auto')."
    )
    parser.add_argument(
        "--reset-db", action="store_true",
        help="Reinicia la base de datos antes de dividir."
    )

    args = parser.parse_args()
    
    pipeline_completo(
        maestro_path=args.maestro,
        excel_path=args.excel,
        salida_path=args.salida,
        plantilla_path=args.plantilla,
        output_dir=args.output_dir,
        db_path=args.db,
        columna_excel=args.columna,
        regex_pattern=args.regex,
        backend=args.backend,
        reset_db=args.reset_db
    )


if __name__ == "__main__":
    main()
