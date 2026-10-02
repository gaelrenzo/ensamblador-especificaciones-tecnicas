"""
db_manager.py - Gestor de Base de Datos SQLite para Partidas de Especificaciones Técnicas

Maneja la creación de tablas, inserción, consulta e indexación de partidas
en la base de datos SQLite (partidas.db).
"""

import os
import sqlite3
from typing import List, Dict, Optional, Tuple, Any


def conectar_db(db_path: str = "partidas.db") -> sqlite3.Connection:
    """Establece conexión con la base de datos SQLite."""
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True) if os.path.dirname(db_path) else None
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: str = "partidas.db", reset: bool = False) -> None:
    """
    Inicializa la base de datos y crea la tabla 'partidas' según las especificaciones.
    Guarda: ID_Partida, Nombre_Partida, Ruta_Archivo y metadatos complementarios.
    """
    conn = conectar_db(db_path)
    cursor = conn.cursor()
    
    if reset:
        cursor.execute("DROP TABLE IF EXISTS partidas")
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS partidas (
            id_partida TEXT PRIMARY KEY,
            nombre_partida TEXT NOT NULL,
            ruta_archivo TEXT NOT NULL,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            num_parrafos INTEGER DEFAULT 0,
            num_tablas INTEGER DEFAULT 0,
            num_imagenes INTEGER DEFAULT 0
        )
    """)
    
    # Índice para búsquedas rápidas por ID y por ruta
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_partidas_id ON partidas(id_partida)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_partidas_ruta ON partidas(ruta_archivo)")
    
    conn.commit()
    conn.close()


def normalizar_id(id_partida: str) -> str:
    """Limpia y normaliza el ID de una partida eliminando espacios extremos."""
    return str(id_partida).strip()


def guardar_partida(
    db_path: str,
    id_partida: str,
    nombre_partida: str,
    ruta_archivo: str,
    stats: Optional[Dict[str, int]] = None
) -> None:
    """Inserta o actualiza una partida en la base de datos."""
    stats = stats or {}
    conn = conectar_db(db_path)
    cursor = conn.cursor()
    
    cursor.execute("""
        INSERT INTO partidas (id_partida, nombre_partida, ruta_archivo, num_parrafos, num_tablas, num_imagenes)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id_partida) DO UPDATE SET
            nombre_partida = excluded.nombre_partida,
            ruta_archivo = excluded.ruta_archivo,
            num_parrafos = excluded.num_parrafos,
            num_tablas = excluded.num_tablas,
            num_imagenes = excluded.num_imagenes,
            fecha_registro = CURRENT_TIMESTAMP
    """, (
        normalizar_id(id_partida),
        nombre_partida.strip(),
        ruta_archivo,
        stats.get("parrafos", 0),
        stats.get("tablas", 0),
        stats.get("imagenes", 0)
    ))
    
    conn.commit()
    conn.close()


def guardar_partidas_batch(db_path: str, partidas: List[Dict[str, Any]]) -> int:
    """Inserta una lista de partidas en una sola transacción."""
    init_db(db_path)
    conn = conectar_db(db_path)
    cursor = conn.cursor()
    
    registros = []
    for p in partidas:
        stats = p.get("stats", {})
        registros.append((
            normalizar_id(p["id_partida"]),
            p["nombre_partida"].strip(),
            p["ruta_archivo"],
            stats.get("parrafos", 0),
            stats.get("tablas", 0),
            stats.get("imagenes", 0)
        ))
    
    cursor.executemany("""
        INSERT INTO partidas (id_partida, nombre_partida, ruta_archivo, num_parrafos, num_tablas, num_imagenes)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id_partida) DO UPDATE SET
            nombre_partida = excluded.nombre_partida,
            ruta_archivo = excluded.ruta_archivo,
            num_parrafos = excluded.num_parrafos,
            num_tablas = excluded.num_tablas,
            num_imagenes = excluded.num_imagenes,
            fecha_registro = CURRENT_TIMESTAMP
    """, registros)
    
    conn.commit()
    total = cursor.rowcount
    conn.close()
    return total


def buscar_partida(db_path: str, id_buscado: str) -> Optional[Dict[str, Any]]:
    """
    Busca una partida por ID exacto o por su variante normalizada
    (soporta formatos '01.01.03' y '01_01_03').
    """
    conn = conectar_db(db_path)
    cursor = conn.cursor()
    
    id_clean = normalizar_id(id_buscado)
    id_con_puntos = id_clean.replace("_", ".")
    id_con_guiones = id_clean.replace(".", "_")
    
    cursor.execute("""
        SELECT id_partida, nombre_partida, ruta_archivo, num_parrafos, num_tablas, num_imagenes, fecha_registro
        FROM partidas
        WHERE id_partida = ? OR id_partida = ? OR id_partida = ?
        LIMIT 1
    """, (id_clean, id_con_puntos, id_con_guiones))
    
    row = cursor.fetchone()
    conn.close()
    
    if row:
        return dict(row)
    return None


def obtener_rutas_partidas_ordenadas(
    db_path: str,
    ids_solicitados: List[str]
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """
    Recibe una lista de IDs de partidas y retorna:
    - Lista de partidas encontradas en el orden exacto solicitado (con ID, nombre y ruta).
    - Lista de IDs solicitados que no fueron encontrados en la base de datos.
    """
    encontradas = []
    no_encontradas = []
    
    for id_req in ids_solicitados:
        id_req_clean = normalizar_id(id_req)
        partida = buscar_partida(db_path, id_req_clean)
        if partida:
            encontradas.append(partida)
        else:
            no_encontradas.append(id_req_clean)
            
    return encontradas, no_encontradas


def listar_todas_las_partidas(db_path: str = "partidas.db") -> List[Dict[str, Any]]:
    """Devuelve todas las partidas indexadas en la base de datos."""
    if not os.path.exists(db_path):
        return []
    conn = conectar_db(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT id_partida, nombre_partida, ruta_archivo, num_parrafos, num_tablas, num_imagenes FROM partidas ORDER BY id_partida ASC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows
