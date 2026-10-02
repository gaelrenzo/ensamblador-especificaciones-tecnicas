"""
ensamblador.py - Fase 3: Ensamblaje (El Merger)

Recibe la lista de partidas solicitadas (desde el archivo Excel o por argumentos),
consulta la base de datos SQLite (partidas.db) para obtener las rutas de los archivos .docx,
respetando estrictamente el orden solicitado en el Excel.
Luego, abre un documento en blanco o una plantilla base (para portadas y encabezados)
y fusiona el contenido de cada archivo de partida sin perder tablas, imágenes, estilos ni márgenes.

Soporta dos motores:
1. 'docxcompose' (Multiplataforma - Linux / Windows / Mac):
   Fusiona documentos Word combinando árboles XML, estilos, catálogos de medios y relaciones rId.
2. 'win32com' (Windows nativo con Word COM):
   Utiliza Range.InsertFile de Microsoft Word.
"""

import os
import sys
import argparse
from typing import List, Dict, Optional, Any

import docx
from db_manager import obtener_rutas_partidas_ordenadas
from lector_excel import leer_solicitud_excel


def detectar_backend_fusion(preferencia: str = "auto") -> str:
    """Detecta el motor de ensamblaje a utilizar."""
    if preferencia == "win32com":
        try:
            import win32com.client
            return "win32com"
        except ImportError:
            print("[ADVERTENCIA] win32com no disponible. Se utilizará 'docxcompose'.")
            return "docxcompose"
            
    if preferencia == "docxcompose":
        return "docxcompose"
        
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            return "win32com"
        except ImportError:
            pass
            
    return "docxcompose"


def fusionar_con_docxcompose(
    archivos_docx: List[str],
    salida_path: str,
    plantilla_path: Optional[str] = None
) -> str:
    """
    Fusiona una lista de archivos .docx utilizando la librería docxcompose.
    Mantiene 100% los formatos, tablas, imágenes, encabezados y estilos.
    """
    from docxcompose.composer import Composer
    
    os.makedirs(os.path.dirname(os.path.abspath(salida_path)), exist_ok=True) if os.path.dirname(salida_path) else None
    
    if plantilla_path and os.path.isfile(plantilla_path):
        print(f"[INFO] Utilizando plantilla base institucional: '{plantilla_path}'")
        doc_base = docx.Document(plantilla_path)
        composer = Composer(doc_base)
        archivos_a_agregar = archivos_docx
    else:
        print("[INFO] Iniciando consolidación a partir de la primera partida como documento base.")
        doc_base = docx.Document(archivos_docx[0])
        composer = Composer(doc_base)
        archivos_a_agregar = archivos_docx[1:]
        
    total_archivos = len(archivos_a_agregar)
    for idx, arch in enumerate(archivos_a_agregar, 1):
        nombre_arch = os.path.basename(arch)
        print(f"  + Anexando [{idx}/{total_archivos}]: {nombre_arch}")
        doc_item = docx.Document(arch)
        composer.append(doc_item)
        
    composer.save(salida_path)
    return salida_path


def fusionar_con_win32com(
    archivos_docx: List[str],
    salida_path: str,
    plantilla_path: Optional[str] = None
) -> str:
    """
    Fusiona una lista de archivos .docx utilizando Microsoft Word vía COM (Windows).
    """
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError("win32com no está disponible. Instale pywin32 o use el motor docxcompose.")
        
    os.makedirs(os.path.dirname(os.path.abspath(salida_path)), exist_ok=True) if os.path.dirname(salida_path) else None
    abs_salida = os.path.abspath(salida_path)
    
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    try:
        if plantilla_path and os.path.isfile(plantilla_path):
            print(f"[INFO] [COM] Cargando plantilla base: '{plantilla_path}'")
            doc = word.Documents.Open(os.path.abspath(plantilla_path))
        else:
            print("[INFO] [COM] Creando nuevo documento Word en blanco...")
            doc = word.Documents.Add()
            
        total = len(archivos_docx)
        for idx, arch in enumerate(archivos_docx, 1):
            abs_arch = os.path.abspath(arch)
            print(f"  + [COM] Insertando [{idx}/{total}]: {os.path.basename(arch)}")
            
            # Posicionar el cursor al final exacto del contenido
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            rng.InsertFile(abs_arch)
            
        doc.SaveAs(abs_salida, FileFormat=16) # 16 = wdFormatDocumentDefault
        doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
        
    return salida_path


def ensamblar_expediente(
    solicitud: Union[List[str], str],
    columna_excel: str = "Código de Partida",
    hoja_excel: Union[str, int] = 0,
    db_path: str = "partidas.db",
    salida_path: str = "Expediente_Consolidado.docx",
    plantilla_path: Optional[str] = None,
    backend: str = "auto"
) -> Dict[str, Any]:
    """
    Orquestador de la Fase 3:
    1. Obtiene la lista de códigos de partida (desde archivo Excel o lista directa).
    2. Consulta la base de datos SQLite para encontrar las rutas de los archivos .docx.
    3. Respeta el orden del Excel e informa sobre partidas no encontradas.
    4. Fusiona los documentos mediante el motor seleccionado.
    5. Retorna informe estadístico del proceso.
    """
    print(f"==================================================")
    print(f" FASE 3: ENSAMBLAJE DEL EXPEDIENTE (EL MERGER)")
    print(f" Base de Datos     : {db_path}")
    print(f" Archivo de Salida : {salida_path}")
    if plantilla_path:
        print(f" Plantilla Base    : {plantilla_path}")
    print(f"==================================================")
    
    # 1. Obtener la lista de partidas solicitadas
    if isinstance(solicitud, str) and (solicitud.endswith(".xlsx") or solicitud.endswith(".xls")):
        print(f"[INFO] Leyendo solicitud desde archivo Excel: '{solicitud}'...")
        ids_solicitados = leer_solicitud_excel(solicitud, columna=columna_excel, hoja=hoja_excel)
    elif isinstance(solicitud, list):
        ids_solicitados = [str(x).strip() for x in solicitud if str(x).strip()]
    else:
        raise ValueError("El parámetro de solicitud debe ser una lista de códigos o la ruta a un archivo .xlsx")
        
    if not ids_solicitados:
        print("[ERROR] La lista de partidas solicitadas está vacía.")
        return {"exito": False, "mensaje": "Lista de partidas vacía"}
        
    print(f"[INFO] Total de códigos solicitados en Excel: {len(ids_solicitados)}")
    
    # 2. Consultar la base de datos SQLite respetando el orden
    if not os.path.isfile(db_path):
        raise FileNotFoundError(f"No se encontró la base de datos SQLite: '{db_path}'. Ejecute primero divisor.py.")
        
    partidas_encontradas, no_encontradas = obtener_rutas_partidas_ordenadas(db_path, ids_solicitados)
    
    # Reportar partidas no encontradas
    if no_encontradas:
        print("\n" + "!" * 50)
        print(f"[ADVERTENCIA] Se solicitaron {len(no_encontradas)} partida(s) NO indexadas en la BD:")
        for nid in no_encontradas:
            print(f"  ⚠ Partida no encontrada: '{nid}'")
        print("!" * 50 + "\n")
        
    if not partidas_encontradas:
        print("[ERROR] Ninguna de las partidas solicitadas fue encontrada en la base de datos.")
        return {
            "exito": False,
            "solicitadas": len(ids_solicitados),
            "encontradas": 0,
            "no_encontradas": no_encontradas
        }
        
    print(f"[INFO] Partidas listas para ensamblar ({len(partidas_encontradas)}/{len(ids_solicitados)}):")
    archivos_a_fusionar = []
    for idx, p in enumerate(partidas_encontradas, 1):
        ruta = p["ruta_archivo"]
        if not os.path.isfile(ruta):
            print(f"  [ERROR] El archivo '{ruta}' está registrado en BD pero no existe físicamente en disco.")
            continue
        archivos_a_fusionar.append(ruta)
        print(f"  {idx:2d}. [{p['id_partida']}] {p['nombre_partida'][:40]} -> {ruta}")
        
    if not archivos_a_fusionar:
        raise FileNotFoundError("No se encontró físicamente ninguno de los archivos .docx de las partidas.")
        
    # 3. Determinar motor de fusión y ensamblar
    motor = detectar_backend_fusion(backend)
    print(f"[INFO] Motor de consolidación seleccionado: {motor.upper()}")
    
    if motor == "win32com":
        resultado_path = fusionar_con_win32com(archivos_a_fusionar, salida_path, plantilla_path)
    else:
        resultado_path = fusionar_con_docxcompose(archivos_a_fusionar, salida_path, plantilla_path)
        
    # Estadísticas del documento final
    tamano_kb = os.path.getsize(resultado_path) / 1024
    doc_final = docx.Document(resultado_path)
    total_parrafos = len(doc_final.paragraphs)
    total_tablas = len(doc_final.tables)
    total_imagenes = len(doc_final.inline_shapes)
    
    print("\n--------------------------------------------------")
    print("✔ Expediente final generado con éxito:")
    print(f"  Ruta final           : {os.path.abspath(resultado_path)}")
    print(f"  Tamaño de archivo    : {tamano_kb:.1f} KB")
    print(f"  Partidas solicitadas : {len(ids_solicitados)}")
    print(f"  Partidas ensambladas : {len(archivos_a_fusionar)}")
    print(f"  Partidas faltantes   : {len(no_encontradas)}")
    print(f"  Total párrafos       : {total_parrafos}")
    print(f"  Total tablas         : {total_tablas}")
    print(f"  Total imágenes       : {total_imagenes}")
    print("--------------------------------------------------")
    
    return {
        "exito": True,
        "salida": resultado_path,
        "solicitadas": len(ids_solicitados),
        "ensambladas": len(archivos_a_fusionar),
        "no_encontradas": no_encontradas,
        "tamano_kb": tamano_kb,
        "tablas": total_tablas,
        "imagenes": total_imagenes
    }


def main():
    parser = argparse.ArgumentParser(
        description="Fase 3: Ensambla el documento Word final a partir de la lista de partidas solicitadas."
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "-e", "--excel",
        help="Ruta al archivo Excel (.xlsx) con la lista de partidas solicitadas."
    )
    group.add_argument(
        "-l", "--lista", nargs="+",
        help="Lista de códigos de partida pasada directamente por línea de comando (ej. 01.01.01 01.02.01)."
    )
    parser.add_argument(
        "-c", "--columna", default="Código de Partida",
        help="Nombre de la columna en Excel con los códigos de partida (default: 'Código de Partida')."
    )
    parser.add_argument(
        "-d", "--db", default="partidas.db",
        help="Ruta a la base de datos SQLite (default: partidas.db)."
    )
    parser.add_argument(
        "-o", "--salida", default="Expediente_Consolidado.docx",
        help="Ruta donde se guardará el documento Word final (default: Expediente_Consolidado.docx)."
    )
    parser.add_argument(
        "-p", "--plantilla", default=None,
        help="Ruta opcional a un archivo .docx base (portada institucional, encabezados/pies de página)."
    )
    parser.add_argument(
        "-b", "--backend", choices=["auto", "docxcompose", "win32com"], default="auto",
        help="Motor de fusión: 'auto', 'docxcompose' o 'win32com'."
    )
    
    args = parser.parse_args()
    
    solicitud = args.excel if args.excel else args.lista
    try:
        ensamblar_expediente(
            solicitud=solicitud,
            columna_excel=args.columna,
            db_path=args.db,
            salida_path=args.salida,
            plantilla_path=args.plantilla,
            backend=args.backend
        )
    except Exception as e:
        print(f"\n[ERROR CRÍTICO] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
