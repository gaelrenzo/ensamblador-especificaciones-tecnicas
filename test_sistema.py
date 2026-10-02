"""
test_sistema.py - Batería de Pruebas Automatizadas Unitarias y de Integración

Verifica:
1. Gestión de base de datos SQLite (db_manager.py).
2. División del documento maestro sin pérdida de tablas ni figuras (divisor.py).
3. Lectura robusta de plantillas Excel con variantes de nombres y celdas nulas (lector_excel.py).
4. Ensamblaje en orden estricto de solicitud y reporte de discrepancias (ensamblador.py).
"""

import os
import unittest
import tempfile
import shutil
import sqlite3
import pandas as pd
import docx

from db_manager import init_db, guardar_partida, buscar_partida, obtener_rutas_partidas_ordenadas
from divisor import ejecutar_division
from lector_excel import leer_solicitud_excel, encontrar_columna_codigo
from ensamblador import ensamblar_expediente


class TestSistemaEspecificaciones(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="test_esp_")
        self.db_path = os.path.join(self.temp_dir, "test_partidas.db")
        self.partidas_dir = os.path.join(self.temp_dir, "test_partidas_db")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_01_db_manager(self):
        """Verifica la persistencia y búsqueda flexible en SQLite."""
        init_db(self.db_path, reset=True)
        
        guardar_partida(
            self.db_path,
            id_partida="01.01.03",
            nombre_partida="DEMOLICION DE VEREDAS",
            ruta_archivo="partidas_db/01_01_03.docx",
            stats={"parrafos": 5, "tablas": 1, "imagenes": 1}
        )
        
        # Búsqueda por formato exacto
        p1 = buscar_partida(self.db_path, "01.01.03")
        self.assertIsNotNone(p1)
        self.assertEqual(p1["nombre_partida"], "DEMOLICION DE VEREDAS")
        self.assertEqual(p1["num_tablas"], 1)
        self.assertEqual(p1["num_imagenes"], 1)
        
        # Búsqueda flexible con guiones bajos
        p2 = buscar_partida(self.db_path, "01_01_03")
        self.assertIsNotNone(p2)
        self.assertEqual(p2["id_partida"], "01.01.03")
        
        # Verificación de orden y elementos faltantes
        solicitados = ["01.01.03", "99.99.99"]
        encontradas, faltantes = obtener_rutas_partidas_ordenadas(self.db_path, solicitados)
        self.assertEqual(len(encontradas), 1)
        self.assertEqual(encontradas[0]["id_partida"], "01.01.03")
        self.assertEqual(faltantes, ["99.99.99"])

    def test_02_lector_excel_flexibilidad(self):
        """Verifica que el lector de Excel admita variaciones de nombres y filtre vacíos."""
        excel_path = os.path.join(self.temp_dir, "test_solicitud.xlsx")
        
        # Probar con encabezado sin tilde y mayúsculas
        df = pd.DataFrame({
            "CODIGO DE PARTIDA": [" 01.01.01 ", None, "01.02.03", "", "02.01.01", "nan"],
            "Metrado": [10, None, 20, None, 30, None]
        })
        df.to_excel(excel_path, index=False)
        
        codigos = leer_solicitud_excel(excel_path)
        self.assertEqual(codigos, ["01.01.01", "01.02.03", "02.01.01"])

    def test_03_division_y_conservacion(self):
        """Verifica que divisor.py aísle correctamente las partidas con tablas e imágenes."""
        # Usar el maestro de prueba generado previamente si existe, o crearlo
        maestro_path = "Maestro_Especificaciones.docx"
        if not os.path.isfile(maestro_path):
            from generar_ejemplo import crear_documento_maestro
            crear_documento_maestro(maestro_path)
            
        partidas = ejecutar_division(
            master_path=maestro_path,
            db_path=self.db_path,
            output_dir=self.partidas_dir,
            backend="docx",
            reset_db=True
        )
        self.assertGreaterEqual(len(partidas), 6)
        
        # Verificar que 01.01.01 tenga al menos una tabla
        p_tabla = [p for p in partidas if p["id_partida"] == "01.01.01"][0]
        self.assertGreaterEqual(p_tabla["stats"]["tablas"], 1)
        self.assertTrue(os.path.isfile(p_tabla["ruta_archivo"]))

    def test_04_ensamblaje_completo(self):
        """Verifica el ensamblaje en orden y reporte del expediente final."""
        # 1. Dividir primero
        maestro_path = "Maestro_Especificaciones.docx"
        partidas = ejecutar_division(
            master_path=maestro_path,
            db_path=self.db_path,
            output_dir=self.partidas_dir,
            backend="docx",
            reset_db=True
        )
        
        # 2. Ensamblar subconjunto en orden inverso deliberado
        orden_inverso = ["02.01.02", "01.01.01"]
        salida_path = os.path.join(self.temp_dir, "expediente_test.docx")
        
        res = ensamblar_expediente(
            solicitud=orden_inverso,
            db_path=self.db_path,
            salida_path=salida_path,
            backend="docxcompose"
        )
        
        self.assertTrue(res["exito"])
        self.assertEqual(res["ensambladas"], 2)
        self.assertTrue(os.path.isfile(salida_path))
        
        # Verificar orden en el documento generado
        doc_res = docx.Document(salida_path)
        titulos = [p.text for p in doc_res.paragraphs if p.text.startswith("0")]
        self.assertTrue(titulos[0].startswith("02.01.02"))
        self.assertTrue(titulos[1].startswith("01.01.01"))


if __name__ == "__main__":
    unittest.main()
