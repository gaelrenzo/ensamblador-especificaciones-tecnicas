"""
lector_excel.py - Fase 2: Interfaz de Solicitud (El Lector de Excel)

Utiliza pandas para leer una plantilla de solicitud en formato Excel (.xlsx),
identifica la columna "Código de Partida" (o variantes comunes de nomenclatura),
extrae los valores en el orden exacto de aparición y genera una lista en memoria
limpiando celdas vacías, nulas o con espacios sobrantes.
"""

import os
import sys
import argparse
import unicodedata
from typing import List, Optional, Union

import pandas as pd


def normalizar_texto_columna(texto: str) -> str:
    """Elimina tildes y convierte a minúsculas para comparaciones flexibles."""
    if not isinstance(texto, str):
        texto = str(texto)
    texto = texto.strip().lower()
    # Remover tildes
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])


def encontrar_columna_codigo(df: pd.DataFrame, nombre_preferido: str = "Código de Partida") -> Optional[str]:
    """
    Busca la columna adecuada dentro del DataFrame, permitiendo coincidencia exacta
    o aproximada (sin tildes, mayúsculas/minúsculas o nombres similares).
    """
    # 1. Coincidencia directa
    if nombre_preferido in df.columns:
        return nombre_preferido
        
    # 2. Coincidencia insensible a mayúsculas y tildes
    norm_preferido = normalizar_texto_columna(nombre_preferido)
    for col in df.columns:
        if normalizar_texto_columna(str(col)) == norm_preferido:
            return col
            
    # 3. Variantes comunes en expedientes técnicos
    variantes_comunes = [
        "codigo de partida",
        "codigo partida",
        "cod partida",
        "codigo",
        "partida",
        "item",
        "id partida",
        "id_partida"
    ]
    
    for col in df.columns:
        col_norm = normalizar_texto_columna(str(col))
        for var in variantes_comunes:
            if col_norm == var:
                return col
                
    return None


def leer_solicitud_excel(
    excel_path: str,
    columna: str = "Código de Partida",
    hoja: Optional[Union[str, int]] = 0
) -> List[str]:
    """
    Lee un archivo Excel y extrae la lista de códigos de partida en orden.
    
    Parámetros:
        excel_path: Ruta al archivo .xlsx
        columna: Nombre de la columna que contiene los códigos de partida
        hoja: Nombre o índice de la hoja a leer (por defecto la primera)
        
    Retorna:
        Lista ordenada de strings con los códigos de partida limpios.
    """
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"El archivo Excel no existe: '{excel_path}'")
        
    # Leer el Excel con pandas
    try:
        df = pd.read_excel(excel_path, sheet_name=hoja, dtype=str)
    except Exception as e:
        raise ValueError(f"Error al leer el archivo Excel '{excel_path}': {e}")
        
    if df.empty:
        print(f"[AVISO] La hoja de Excel '{excel_path}' está vacía.")
        return []
        
    col_encontrada = encontrar_columna_codigo(df, columna)
    if col_encontrada is None:
        columnas_disponibles = ", ".join([f"'{c}'" for c in df.columns])
        raise KeyError(
            f"No se encontró la columna '{columna}' en el archivo Excel. "
            f"Columnas disponibles en la hoja: [{columnas_disponibles}]"
        )
        
    # Extraer la serie, descartar valores nulos/vacíos y limpiar espacios
    serie = df[col_encontrada].dropna()
    
    lista_codigos = []
    for val in serie:
        val_str = str(val).strip()
        # Ignorar celdas vacías o con texto "nan" o "none"
        if val_str and val_str.lower() not in ("nan", "none", ""):
            lista_codigos.append(val_str)
            
    return lista_codigos


def main():
    parser = argparse.ArgumentParser(
        description="Fase 2: Lee una plantilla Excel y extrae la lista de códigos de partida."
    )
    parser.add_argument(
        "-e", "--excel", required=True,
        help="Ruta al archivo Excel (.xlsx) con la solicitud de partidas."
    )
    parser.add_argument(
        "-c", "--columna", default="Código de Partida",
        help="Nombre de la columna que contiene los códigos (default: 'Código de Partida')."
    )
    parser.add_argument(
        "-s", "--hoja", default=0,
        help="Nombre o índice numérico de la hoja de cálculo (default: 0)."
    )
    
    args = parser.parse_args()
    
    # Manejar si el argumento de hoja es un número o texto
    hoja_param = int(args.hoja) if str(args.hoja).isdigit() else args.hoja
    
    try:
        print(f"==================================================")
        print(f" FASE 2: INTERFAZ DE SOLICITUD (LECTOR DE EXCEL)")
        print(f" Archivo Excel : {args.excel}")
        print(f" Columna       : {args.columna}")
        print(f"==================================================")
        
        codigos = leer_solicitud_excel(args.excel, columna=args.columna, hoja=hoja_param)
        
        print(f"✔ Se extrajeron {len(codigos)} partidas solicitadas en memoria:")
        for idx, cod in enumerate(codigos, 1):
            print(f"  {idx:2d}. {cod}")
            
    except Exception as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
