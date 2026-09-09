# -*- coding: utf-8 -*-
"""
residuos_processor.py
Módulo portátil para o cálculo matemático e xeración do Estudio de Xestión de Residuos (EGR).
Recrea exactamente a lóxica da plantilla Excel "PLANTILLA TABLAS.xlsx" e actualiza o Word "XESTIÓN DE RESIDUOS.docx".
"""

import os
import glob
import re
import sys
import docx

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_default_residuos_paths():
    """Busca dinamicamente o arquivo .docx de XESTIÓN DE RESIDUOS e a carpeta de saída."""
    cwd = os.getcwd()
    search_dirs = [cwd, SCRIPT_DIR]
    
    for d in search_dirs:
        try:
            candidates = [
                os.path.join(d, f) for f in os.listdir(d)
                if f.lower().endswith(".docx")
                and not f.startswith("~$")
                and ("residuos" in f.lower() or "egr" in f.lower() or "xestión" in f.lower() or "gestion" in f.lower())
                and "modificado" not in f.lower()
            ]
            if candidates:
                return candidates[0], cwd
        except Exception:
            pass

    return os.path.join(cwd, "XESTIÓN DE RESIDUOS.docx"), cwd

DEFAULT_RESIDUOS_DATA = {
    # Paso 1: Datos do Proxecto
    "tipo_obra": "Melloras de condicións de seguridade, salubridade e ornato",
    "situacion": "Rúa Baterías 34",
    "poboacion": "Ferrol",
    "promotor": "El Bosque Providencia S.L.",
    "arquitecto": "Estudio Anta Arquitectos S.L.P.",
    "orzamento_pem": "O indicado no capítulo homónimo do presente proxecto",
    "duracion_obra": "3 meses",
    "superficie_construida": "9,84 m2",
    "superficie_actuacion": 9.84,  # m2 (D4 no Excel)

    # Paso 2 e 3: Selección de Residuos (SI/NO en columna B do Excel)
    "residuos_selection": {
        "tierras_excavacion": True,   # LER 17 05 04
        "asfalto": True,              # LER 17 03 02
        "madera": True,               # LER 17 02 01
        "envases_metalicos": True,    # LER 15 01 04
        "cobre_laton": True,          # LER 17 04 01
        "hierro_acero": True,         # LER 17 04 05
        "metales_mezclados": True,    # LER 17 04 07
        "cables": True,               # LER 17 04 11
        "envases_papel": True,        # LER 15 01 01
        "plastico": True,             # LER 17 02 03
        "vidrio": True,               # LER 17 02 02
        "yeso": True,                 # LER 17 08 02
        "aislamiento": True,          # LER 17 06 04
        "mezclados_rcd": True,        # LER 17 09 04
        "grava_rocas": True,          # LER 01 04 08
        "arena_arcillas": True,       # LER 01 04 09
        "hormigon": True,             # LER 17 01 01
        "ladrillos": True,            # LER 17 01 02
        "tejas_ceramicos": False,     # LER 17 01 03
        "pintura_barniz": True,       # LER 08 01 11
        "amianto": False,             # LER 17 06 05
        "detergentes": True,          # LER 20 01 30
    },

    # Paso 5: Data e Asinantes
    "data_documento": "Ferrol, xullo de 2026",
    "asinantes": "Fdo. Sergio J. Beceiro Lodeiro     M. Rosa Vilas Romalde",
    "colexiado": "ESTUDIO ANTA ARQUITECTOS S.L.P.\tNº COAG - 20.039",
    "nome_ficheiro_saida": "XESTION_DE_RESIDUOS_Modificado.docx"
}

# Definición de ítems do Excel (clave, nome, LER, %peso, densidad, tratamento, destino)
RESIDUOS_ITEMS_DEF = [
    {"key": "tierras_excavacion", "name": "Tierra y piedras distintas de las especificadas en el código 17 05 03.", "ler": "17 05 04", "pct": 0.05, "dens": 1.66, "trat": "Depósito / Tratamiento", "dest": "Gestor autorizado RNPs", "level": "I"},
    {"key": "asfalto", "name": "Mezclas bituminosas distintas de las especificadas en el código 17 03 01.", "ler": "17 03 02", "pct": 0.04, "dens": 1.00, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "madera", "name": "Madera.", "ler": "17 02 01", "pct": 0.04, "dens": 1.10, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "envases_metalicos", "name": "Envases metálicos.", "ler": "15 01 04", "pct": 0.01, "dens": 0.60, "trat": "Depósito / Tratamiento", "dest": "Gestor autorizado RNPs", "level": "II_NP_METALES"},
    {"key": "cobre_laton", "name": "Cobre, bronce, latón.", "ler": "17 04 01", "pct": 0.01, "dens": 1.50, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP_METALES"},
    {"key": "hierro_acero", "name": "Hierro y acero.", "ler": "17 04 05", "pct": 0.01, "dens": 2.10, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP_METALES"},
    {"key": "metales_mezclados", "name": "Metales mezclados.", "ler": "17 04 07", "pct": 0.01, "dens": 1.50, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP_METALES"},
    {"key": "cables", "name": "Cables distintos de los especificados en el código 17 04 10.", "ler": "17 04 11", "pct": 0.01, "dens": 1.50, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP_METALES"},
    {"key": "envases_papel", "name": "Envases de papel y cartón.", "ler": "15 01 01", "pct": 0.03, "dens": 0.75, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "plastico", "name": "Plástico.", "ler": "17 02 03", "pct": 0.01, "dens": 0.60, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "vidrio", "name": "Vidrio.", "ler": "17 02 02", "pct": 0.005, "dens": 1.00, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "yeso", "name": "Materiales de construcción a partir de yeso distintos de los especificados en el código 17 08 01.", "ler": "17 08 02", "pct": 0.01, "dens": 1.00, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "aislamiento", "name": "Materiales de aislamiento distintos de los especificados en los códigos 17 06 01 y 17 06 03.", "ler": "17 06 04", "pct": 0.02, "dens": 0.60, "trat": "Reciclado", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "mezclados_rcd", "name": "Residuos mezclados de construcción y demolición distintos de los especificados en los códigos 17 09 01, 17 09 02 y 17 09 03.", "ler": "17 09 04", "pct": 0.035, "dens": 1.50, "trat": "Depósito / Tratamiento", "dest": "Gestor autorizado RNPs", "level": "II_NP"},
    {"key": "grava_rocas", "name": "Residuos de grava y rocas trituradas distintos de los mencionados en el código 01 04 07.", "ler": "01 04 08", "pct": 0.02, "dens": 1.50, "trat": "Reciclado", "dest": "Planta reciclaje RCD", "level": "II_PETREO_ARIDOS"},
    {"key": "arena_arcillas", "name": "Residuos de arena y arcillas.", "ler": "01 04 09", "pct": 0.02, "dens": 1.60, "trat": "Reciclado", "dest": "Planta reciclaje RCD", "level": "II_PETREO_ARIDOS"},
    {"key": "hormigon", "name": "Hormigón (hormigones, morteros y prefabricados).", "ler": "17 01 01", "pct": 0.12, "dens": 1.50, "trat": "Reciclado / Vertedero", "dest": "Planta reciclaje RCD", "level": "II_PETREO"},
    {"key": "ladrillos", "name": "Ladrillos.", "ler": "17 01 02", "pct": 0.20, "dens": 1.25, "trat": "Reciclado", "dest": "Planta reciclaje RCD", "level": "II_PETREO_CERAMICOS"},
    {"key": "tejas_ceramicos", "name": "Tejas y materiales cerámicos.", "ler": "17 01 03", "pct": 0.30, "dens": 1.25, "trat": "Reciclado", "dest": "Planta reciclaje RCD", "level": "II_PETREO_CERAMICOS"},
    {"key": "pintura_barniz", "name": "Residuos de pintura y barniz que contienen disolventes orgánicos u otras sustancias peligrosas.", "ler": "08 01 11", "pct": 0.02, "dens": 0.90, "trat": "Depósito / Tratamiento", "dest": "Gestor autorizado RPs", "level": "PELIGROSOS"},
    {"key": "amianto", "name": "Materiales de construcción que contienen amianto.", "ler": "17 06 05", "pct": 0.02, "dens": 0.24, "trat": "Depósito / Tratamiento", "dest": "Gestor autorizado RPs", "level": "PELIGROSOS"},
    {"key": "detergentes", "name": "Detergentes distintos de los especificados en el código 20 01 29.", "ler": "20 01 30", "pct": 0.01, "dens": 1.00, "trat": "Tratamiento Fco/Qco", "dest": "Gestor autorizado RNPs", "level": "PELIGROSOS"},
]

def calculate_residuos(data):
    """
    Executa os cálculos exactos do Excel PLANTILLA TABLAS.xlsx a partir de superficie_actuacion e seleccións.
    """
    try:
        sup = float(data.get("superficie_actuacion", 9.84))
    except (ValueError, TypeError):
        sup = 9.84

    v_tot = sup * 0.2          # m3 volumen total residuos (D4 * 0.2)
    t_tot = v_tot              # toneladas total residuos

    sel = data.get("residuos_selection", {})

    items_calc = {}
    for item in RESIDUOS_ITEMS_DEF:
        k = item["key"]
        is_active = bool(sel.get(k, True))
        if is_active:
            w = t_tot * item["pct"]       # Peso (t)
            v = w * item["dens"]          # Volumen (m3) = dens * peso (según fórmula Excel =G*H)
        else:
            w = 0.0
            v = 0.0
        
        items_calc[k] = {
            "key": k,
            "name": item["name"],
            "ler": item["ler"],
            "pct": item["pct"],
            "dens": item["dens"],
            "trat": item["trat"],
            "dest": item["dest"],
            "level": item["level"],
            "active": is_active,
            "peso": w,
            "volumen": v
        }

    # Sumas para Tabla 1 (Resumen por familia)
    resumen_tbl1 = {
        "tierras": {"peso": items_calc["tierras_excavacion"]["peso"], "volumen": items_calc["tierras_excavacion"]["volumen"]},
        "asfalto": {"peso": items_calc["asfalto"]["peso"], "volumen": items_calc["asfalto"]["volumen"]},
        "madera": {"peso": items_calc["madera"]["peso"], "volumen": items_calc["madera"]["volumen"]},
        "metales": {
            "peso": sum(items_calc[k]["peso"] for k in ["envases_metalicos", "cobre_laton", "hierro_acero", "metales_mezclados", "cables"]),
            "volumen": sum(items_calc[k]["volumen"] for k in ["envases_metalicos", "cobre_laton", "hierro_acero", "metales_mezclados", "cables"])
        },
        "papel": {"peso": items_calc["envases_papel"]["peso"], "volumen": items_calc["envases_papel"]["volumen"]},
        "plastico": {"peso": items_calc["plastico"]["peso"], "volumen": items_calc["plastico"]["volumen"]},
        "vidrio": {"peso": items_calc["vidrio"]["peso"], "volumen": items_calc["vidrio"]["volumen"]},
        "yeso": {"peso": items_calc["yeso"]["peso"], "volumen": items_calc["yeso"]["volumen"]},
        "aridos": {
            "peso": items_calc["grava_rocas"]["peso"] + items_calc["arena_arcillas"]["peso"],
            "volumen": items_calc["grava_rocas"]["volumen"] + items_calc["arena_arcillas"]["volumen"]
        },
        "hormigon": {"peso": items_calc["hormigon"]["peso"], "volumen": items_calc["hormigon"]["volumen"]},
        "ceramicos": {
            "peso": items_calc["ladrillos"]["peso"] + items_calc["tejas_ceramicos"]["peso"],
            "volumen": items_calc["ladrillos"]["volumen"] + items_calc["tejas_ceramicos"]["volumen"]
        },
        "peligrosos": {
            "peso": sum(items_calc[k]["peso"] for k in ["pintura_barniz", "amianto", "detergentes"]),
            "volumen": sum(items_calc[k]["volumen"] for k in ["pintura_barniz", "amianto", "detergentes"])
        }
    }

    # Tabla 3: Separación obligatoria en obra
    separacion_tbl3 = [
        {"tipo": "Hormigón LER 17 01 01", "peso": resumen_tbl1["hormigon"]["peso"], "umbral": 80.0, "obligatoria": resumen_tbl1["hormigon"]["peso"] > 80.0},
        {"tipo": "Ladrillos, tejas y materiales cerámicos LER 17 01 02, LER 17 01 03", "peso": resumen_tbl1["ceramicos"]["peso"], "umbral": 40.0, "obligatoria": resumen_tbl1["ceramicos"]["peso"] > 40.0},
        {"tipo": "Piedra LER 17 05 04", "peso": resumen_tbl1["tierras"]["peso"], "umbral": None, "obligatoria": False},
        {"tipo": "Metales (incluidas sus aleaciones) LER 17 04", "peso": resumen_tbl1["metales"]["peso"], "umbral": 2.0, "obligatoria": resumen_tbl1["metales"]["peso"] > 2.0},
        {"tipo": "Madera LER 17 02 01", "peso": resumen_tbl1["madera"]["peso"], "umbral": 1.0, "obligatoria": resumen_tbl1["madera"]["peso"] > 1.0},
        {"tipo": "Plástico LER 17 02 03", "peso": resumen_tbl1["plastico"]["peso"], "umbral": 0.50, "obligatoria": resumen_tbl1["plastico"]["peso"] > 0.50},
        {"tipo": "Vidrio LER 17 02 02", "peso": resumen_tbl1["vidrio"]["peso"], "umbral": 1.0, "obligatoria": resumen_tbl1["vidrio"]["peso"] > 1.0},
        {"tipo": "Yeso LER 17 08 02", "peso": resumen_tbl1["yeso"]["peso"], "umbral": None, "obligatoria": False},
        {"tipo": "Papel y cartón LER 15 01 01", "peso": resumen_tbl1["papel"]["peso"], "umbral": 0.50, "obligatoria": resumen_tbl1["papel"]["peso"] > 0.50},
    ]

    total_peso_general = sum(it["peso"] for it in items_calc.values())
    total_volumen_general = sum(it["volumen"] for it in items_calc.values())

    return {
        "superficie_actuacion": sup,
        "volumen_total_estimado": v_tot,
        "toneladas_total_estimadas": t_tot,
        "items": items_calc,
        "resumen_tbl1": resumen_tbl1,
        "separacion_tbl3": separacion_tbl3,
        "total_peso_general": total_peso_general,
        "total_volumen_general": total_volumen_general
    }

def fmt(num):
    """Formatea número a 2 decimais con coma decimal estilo europeo (0,00)."""
    return f"{num:.2f}".replace(".", ",")

def set_cell_text_preserve_style(cell, text):
    """Establece o texto dunha celda mantendo formato e estilo."""
    if not cell.paragraphs:
        cell.text = text
        return
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = str(text)
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = str(text)
    while len(cell.paragraphs) > 1:
        p_extra = cell.paragraphs[-1]._element
        p_extra.getparent().remove(p_extra)
def set_paragraph_text_preserve_style(p, text):
    if p.runs:
        p.runs[0].text = str(text)
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = str(text)

def generate_residuos(data, template_path=None, output_path=None, output_stream=None):
    """
    Xera o documento Word de Xestión de Residuos (EGR) actualizando táboas e mantendo maquetación.
    Se output_stream se proporciona, garda directamente na memoria (BytesIO).
    """
    default_tpl, default_out_dir = get_default_residuos_paths()

    if template_path is None:
        template_path = default_tpl

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Non se atopou o arquivo de plantilla Word: {template_path}")

    filename = data.get("nome_ficheiro_saida", "XESTION_DE_RESIDUOS_Modificado.docx")
    if not filename.lower().endswith(".docx"):
        filename += ".docx"

    if output_path is None and output_stream is None:
        out_dir = os.path.dirname(template_path) if template_path else default_out_dir
        output_path = os.path.join(out_dir, filename)

    # Executar cálculos
    calc = calculate_residuos(data)

    # Cargar documento Word (manexando bloqueo se o orixinal está aberto en Word)
    try:
        doc = docx.Document(template_path)
    except Exception:
        import tempfile, subprocess
        tmp = tempfile.mktemp(suffix=".docx")
        try:
            cmd = f'powershell -Command "Copy-Item -LiteralPath \'{template_path}\' -Destination \'{tmp}\' -Force"'
            subprocess.run(cmd, shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            doc = docx.Document(tmp)
        except Exception:
            doc = docx.Document(template_path)

    # 1. Actualizar Tabla 1 (Datos generales)
    if len(doc.tables) > 0:
        t1 = doc.tables[0]
        map_t1 = [
            (0, "tipo_obra"),
            (1, "situacion"),
            (2, "poboacion"),
            (3, "promotor"),
            (4, "arquitecto"),
            (5, "orzamento_pem"),
            (6, "duracion_obra"),
            (7, "superficie_construida")
        ]
        for row_idx, key in map_t1:
            if row_idx < len(t1.rows):
                val = str(data.get(key, ""))
                set_cell_text_preserve_style(t1.rows[row_idx].cells[1], val)

    # 2. Actualizar Tabla 2 (Estimación detallada de residuos)
    if len(doc.tables) > 1:
        t2 = doc.tables[1]
        items = calc["items"]
        map_t2 = [
            (2, items["tierras_excavacion"]), (3, items["grava_rocas"]), (4, items["arena_arcillas"]),
            (8, items["hormigon"]), (9, items["ladrillos"]), (10, items["tejas_ceramicos"]),
            (14, items["madera"]), (15, items["vidrio"]), (16, items["plastico"]),
            (18, items["cobre_laton"]), (19, items["hierro_acero"]), (20, items["metales_mezclados"]),
            (21, items["cables"]), (22, items["envases_papel"]), (23, items["envases_metalicos"]),
            (24, items["asfalto"]), (25, items["yeso"]), (26, items["aislamiento"]),
            (27, items["amianto"]), (29, items["mezclados_rcd"]),
            (31, items["pintura_barniz"]), (32, items["detergentes"])
        ]
        for row_idx, it_data in map_t2:
            if row_idx < len(t2.rows):
                r_cells = t2.rows[row_idx].cells
                if len(r_cells) >= 6:
                    set_cell_text_preserve_style(r_cells[4], fmt(it_data["peso"]))
                    set_cell_text_preserve_style(r_cells[5], fmt(it_data["volumen"]))

    # 3. Actualizar Tabla 3 (Separación obligatoria en obra)
    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        for idx, sep in enumerate(calc["separacion_tbl3"], start=1):
            if idx < len(t3.rows):
                r_cells = t3.rows[idx].cells
                if len(r_cells) >= 5:
                    set_cell_text_preserve_style(r_cells[2], fmt(sep["peso"]))
                    ob_str = "OBLIGATORIA" if sep["obligatoria"] else "NO OBLIGATORIA"
                    set_cell_text_preserve_style(r_cells[4], ob_str)

    # 4. Actualizar Tabla 4 (Firmas)
    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        if len(t4.rows) > 0:
            set_cell_text_preserve_style(t4.rows[0].cells[0], data.get("asinantes", ""))
        if len(t4.rows) > 1:
            set_cell_text_preserve_style(t4.rows[1].cells[0], data.get("colexiado", ""))

    # 5. Actualizar data e párrafos de cierre
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("Ferrol,") or txt.startswith("A Coruña,") or " 202" in txt:
            if len(txt) < 40:
                set_paragraph_text_preserve_style(p, data.get("data_documento", "Ferrol, xullo 2026"))

    if output_stream is not None:
        doc.save(output_stream)
        output_stream.seek(0)
        return {
            "status": "success",
            "filename": filename,
            "calc": {
                "superficie_actuacion": calc["superficie_actuacion"],
                "volumen_total_estimado": calc["volumen_total_estimado"],
                "toneladas_total_estimadas": calc["toneladas_total_estimadas"],
                "total_peso_general": calc["total_peso_general"],
                "total_volumen_general": calc["total_volumen_general"]
            }
        }

    # Gardar con xestión de PermissionError (se o archivo de saída está aberto en Word)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    fallback_used = False
    warning_msg = None

    try:
        doc.save(output_path)
    except PermissionError:
        import datetime
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(output_path)
        output_path = f"{base}_{ts}{ext}"
        doc.save(output_path)
        fallback_used = True
        warning_msg = "O archivo orixinal estaba aberto en Word ou noutro programa. Xerouse unha copia con marca temporal."

    return {
        "status": "success",
        "output_path": os.path.abspath(output_path),
        "filename": os.path.basename(output_path),
        "size_bytes": os.path.getsize(output_path),
        "fallback_used": fallback_used,
        "warning": warning_msg,
        "calc": {
            "superficie_actuacion": calc["superficie_actuacion"],
            "volumen_total_estimado": calc["volumen_total_estimado"],
            "toneladas_total_estimadas": calc["toneladas_total_estimadas"],
            "total_peso_general": calc["total_peso_general"],
            "total_volumen_general": calc["total_volumen_general"]
        }
    }

if __name__ == "__main__":
    print("Probando residuos_processor.py...")
    res = generate_residuos(DEFAULT_RESIDUOS_DATA)
    print("Resultado xeración Residuos:", res)
