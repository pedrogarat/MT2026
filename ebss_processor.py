# -*- coding: utf-8 -*-
"""
ebss_processor.py
Módulo portátil para a lectura, análise e xeración do Estudio Básico de Seguridade e Saúde (EBSS)
conservando exactamente os estilos, táboas, formatos de parágrafo, composición e imaxes orixinais.
Deseñado para ser 100% portable a calquera ordenador e carpeta de proxecto sen rutas absolutas fixas.
"""

import os
import re
import sys
import docx

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_default_paths():
    """
    Atopa dinamicamente o arquivo de plantilla .docx e a carpeta de saída
    de xeito totalmente portátil en calquera ordenador ou directorio.
    """
    cwd = os.getcwd()
    
    # 1. Buscar se hai un .docx no directorio actual de traballo (onde se executou o .bat)
    try:
        cwd_docx = [
            os.path.join(cwd, f) for f in os.listdir(cwd)
            if f.lower().endswith(".docx")
            and not f.startswith("~$")
            and "modificado" not in f.lower()
            and "test" not in f.lower()
        ]
        if cwd_docx:
            for doc_path in cwd_docx:
                base = os.path.basename(doc_path).lower()
                if "estudo" in base or "seguridade" in base or "ebss" in base or "salud" in base:
                    return doc_path, cwd
            return cwd_docx[0], cwd
    except Exception:
        pass

    # 2. Buscar no directorio onde reside este script python
    try:
        script_docx = [
            os.path.join(SCRIPT_DIR, f) for f in os.listdir(SCRIPT_DIR)
            if f.lower().endswith(".docx")
            and not f.startswith("~$")
            and "modificado" not in f.lower()
            and "test" not in f.lower()
        ]
        if script_docx:
            return script_docx[0], cwd
    except Exception:
        pass

    # 3. Fallback á ruta de descargas orixinal se existe
    fallback_dir = r"C:\Users\Sergio\Downloads\EBSS"
    if os.path.exists(fallback_dir):
        try:
            f_docx = [
                os.path.join(fallback_dir, f) for f in os.listdir(fallback_dir)
                if f.lower().endswith(".docx")
                and not f.startswith("~$")
                and "modificado" not in f.lower()
            ]
            if f_docx:
                return f_docx[0], fallback_dir
        except Exception:
            pass

    return os.path.join(cwd, "ESTUDO BÁSICO DE SEGURIDADE E SAÚDE.docx"), cwd

# Definir para compatibilidade con calquera importador
DEFAULT_TEMPLATE_PATH, DEFAULT_OUTPUT_DIR = get_default_paths()

# Valores predeterminados iniciais
DEFAULT_DATA = {
    # Paso 1: Datos Xerais do Proxecto (Táboa 1)
    "tipo_obra": "Melloras de condicións de seguridade, salubridade e ornato",
    "situacion": "Rúa Baterías 34",
    "poboacion": "Ferrol",
    "promotor": "El Bosque Providencia S.L.",
    "arquitecto": "Estudio Anta Arquitectos S.L.P.",
    "coordinador_ss": "A determinar",
    "orzamento_pem": "O indicado no capítulo homónimo do presente proxecto",
    "duracion_obra": "3 meses",
    "num_traballadores": "6",

    # Paso 2: Condicionantes do Emprazamento (Táboa 2)
    "accesos_obra": "Directos dende vía pública pavimentada",
    "topografia_terreo": "Lixeira pendente",
    "tipo_chan": "-",
    "edificacions_lindeiras": "Si",
    "subministracion_electrica": "Si",
    "subministracion_auga": "Si",
    "sistema_saneamento": "Si",

    # Paso 3: Fases da Obra (Táboa 3)
    "fase_demolicions": "Si (puntuais)",
    "fase_terras": "Non",
    "fase_estruturas": "Non",
    "fase_cubertas": "Non",
    "fase_albanelaria": "Si",
    "fase_acabados": "Si",
    "fase_instalacions": "Si",

    # Paso 4: Instalacións provisionais e Asistencia sanitaria (P#43 e Táboa 4)
    "instalacions_provisionais_necesarias": False,
    "instalacions_provisionais_texto": "Non se estiman necesarias instalacións provisionais para a presente obra, debido as súas características e escasa duración.",
    "tlf_urxencias": "981 33 66 33",
    "centro_saude_nome": "Centro de Saúde Fontenla Maristany",
    "centro_saude_enderezo": "Praza de España, 19, 15403 Ferrol",
    "centro_saude_distancia": "1,1 km",
    "tlf_hospital": "981 33 40 00",
    "hospital_nome": "Hospital Arquitecto Marcide",
    "hospital_enderezo": "Av. de la Residencia s/n, 15405 Ferrol",
    "hospital_distancia": "4,5 km",

    # Paso 5: Maquinaria pesada de obra (Táboa 5 e Táboas 6 a 23)
    "maquinaria": {
        "guindastres_torre": False,
        "pison_vibrante": False,
        "bomba_formigonado": False,
        "montacargas": False,
        "movemento_terras": False,
        "serra_circular": True,
        "compresor": False,
        "martelo_pneumatico": True,
        "rozadora_radial": True,
        "vibrador": False,
        "formigoneira_basculante": False,
        "camion_guindastre": True,
        "camion_formigoneira": False,
        "camions": True,
        "cabrestantes_mecanicos": True,
        "dumper_motovolquete": False,
        "grupo_electroxeno": False,
        "pistola_clavadora": False,
        "soldadura_electrica": True,
    },

    # Medios auxiliares (Táboa 24)
    "medios_auxiliares": {
        "estadas_colgadas": False,
        "estadas_tubulares": False,
        "estadas_borriquetas": False,
        "escaleiras_man": False,
        "instalacion_electrica": False,
    },

    # Paso 6: Riscos Evitables (Táboa 25)
    "riscos_evitables": {
        "rotura_instalacions": True,
        "linas_alta_tension": False,
    },

    # Paso 7: Data, Asinatura e Colexiado (Párrafos finais)
    "data_documento": "Ferrol, agosto de 2026",
    "asinantes": "Fdo. Sergio J. Beceiro Lodeiro     M. Rosa Vilas Romalde",
    "colexiado": "ESTUDIO ANTA ARQUITECTOS S.L.P.\tNº COAG - 20.039",
    "nome_ficheiro_saida": "ESTUDO_BASICO_SEGURIDADE_E_SAUDE_Modificado.docx"
}

# Mapa de maquinaria na Táboa 5
MAQUINARIA_MAP = {
    "guindastres_torre":       {"row": 0, "col_check": 0, "col_name": 1, "table_idx": 5},
    "pison_vibrante":          {"row": 1, "col_check": 0, "col_name": 1, "table_idx": 6},
    "bomba_formigonado":       {"row": 2, "col_check": 0, "col_name": 1, "table_idx": 8},
    "montacargas":             {"row": 3, "col_check": 0, "col_name": 1, "table_idx": None},
    "movemento_terras":        {"row": 4, "col_check": 0, "col_name": 1, "table_idx": 11},
    "serra_circular":          {"row": 5, "col_check": 0, "col_name": 1, "table_idx": 20},
    "compresor":               {"row": 6, "col_check": 0, "col_name": 1, "table_idx": 15},
    "martelo_pneumatico":      {"row": 7, "col_check": 0, "col_name": 1, "table_idx": 17},
    "rozadora_radial":         {"row": 8, "col_check": 0, "col_name": 1, "table_idx": 19},
    "vibrador":                {"row": 9, "col_check": 0, "col_name": 1, "table_idx": 22},
    "formigoneira_basculante": {"row": 0, "col_check": 2, "col_name": 3, "table_idx": 10},
    "camion_guindastre":       {"row": 1, "col_check": 2, "col_name": 3, "table_idx": 7},
    "camion_formigoneira":     {"row": 2, "col_check": 2, "col_name": 3, "table_idx": 9},
    "camions":                 {"row": 3, "col_check": 2, "col_name": 3, "table_idx": 13},
    "cabrestantes_mecanicos":  {"row": 4, "col_check": 2, "col_name": 3, "table_idx": None},
    "dumper_motovolquete":     {"row": 5, "col_check": 2, "col_name": 3, "table_idx": 14},
    "grupo_electroxeno":       {"row": 6, "col_check": 2, "col_name": 3, "table_idx": 16},
    "pistola_clavadora":       {"row": 7, "col_check": 2, "col_name": 3, "table_idx": 18},
    "soldadura_electrica":     {"row": 8, "col_check": 2, "col_name": 3, "table_idx": 21},
}

MEDIOS_AUXILIARES_MAP = {
    "estadas_colgadas":   {"row": 1},
    "estadas_tubulares":  {"row": 3},
    "estadas_borriquetas":{"row": 5},
    "escaleiras_man":     {"row": 7},
    "instalacion_electrica":{"row": 9},
}

def set_cell_text_preserve_style(cell, text):
    """Establece o texto dunha celda mantendo o formato, fonte e parágrafo orixinais."""
    if not cell.paragraphs:
        cell.text = text
        return
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = text
    while len(cell.paragraphs) > 1:
        p_extra = cell.paragraphs[-1]._element
        p_extra.getparent().remove(p_extra)

def set_paragraph_text_preserve_style(p, text):
    """Establece o texto dun parágrafo conservando os estilos do primeiro run."""
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = text

def update_table_1(table, data):
    """Actualiza a Táboa 1: Datos do Proxecto."""
    rows = table.rows
    set_cell_text_preserve_style(rows[0].cells[1], data.get("tipo_obra", ""))
    set_cell_text_preserve_style(rows[1].cells[1], data.get("situacion", ""))
    set_cell_text_preserve_style(rows[2].cells[1], data.get("poboacion", ""))
    set_cell_text_preserve_style(rows[3].cells[1], data.get("promotor", ""))
    set_cell_text_preserve_style(rows[4].cells[1], data.get("arquitecto", ""))
    set_cell_text_preserve_style(rows[5].cells[1], data.get("coordinador_ss", ""))
    set_cell_text_preserve_style(rows[6].cells[1], data.get("orzamento_pem", ""))
    set_cell_text_preserve_style(rows[7].cells[1], data.get("duracion_obra", ""))
    set_cell_text_preserve_style(rows[8].cells[1], str(data.get("num_traballadores", "")))

def update_table_2(table, data):
    """Actualiza a Táboa 2: Características e Condicionantes do Emprazamento."""
    rows = table.rows
    set_cell_text_preserve_style(rows[0].cells[1], data.get("accesos_obra", ""))
    set_cell_text_preserve_style(rows[1].cells[1], data.get("topografia_terreo", ""))
    set_cell_text_preserve_style(rows[2].cells[1], data.get("tipo_chan", "-"))
    set_cell_text_preserve_style(rows[3].cells[1], data.get("edificacions_lindeiras", "Si"))
    set_cell_text_preserve_style(rows[4].cells[1], data.get("subministracion_electrica", "Si"))
    set_cell_text_preserve_style(rows[5].cells[1], data.get("subministracion_auga", "Si"))
    set_cell_text_preserve_style(rows[6].cells[1], data.get("sistema_saneamento", "Si"))

def update_table_3(table, data):
    """Actualiza a Táboa 3: Características xerais da obra e fases de que consta."""
    rows = table.rows
    set_cell_text_preserve_style(rows[0].cells[1], data.get("fase_demolicions", "Non"))
    set_cell_text_preserve_style(rows[1].cells[1], data.get("fase_terras", "Non"))
    set_cell_text_preserve_style(rows[2].cells[1], data.get("fase_estruturas", "Non"))
    set_cell_text_preserve_style(rows[3].cells[1], data.get("fase_cubertas", "Non"))
    set_cell_text_preserve_style(rows[4].cells[1], data.get("fase_albanelaria", "Si"))
    set_cell_text_preserve_style(rows[5].cells[1], data.get("fase_acabados", "Si"))
    set_cell_text_preserve_style(rows[6].cells[1], data.get("fase_instalacions", "Si"))

def update_table_4(table, data):
    """Actualiza a Táboa 4: Asistencia Sanitaria de Referencia."""
    c0 = table.rows[1].cells[0]
    if len(c0.paragraphs) > 1:
        set_paragraph_text_preserve_style(c0.paragraphs[1], data.get("tlf_urxencias", "981 33 66 33"))
    c1 = table.rows[1].cells[1]
    if len(c1.paragraphs) > 1:
        set_paragraph_text_preserve_style(c1.paragraphs[0], f"{data.get('centro_saude_nome', '')},")
        set_paragraph_text_preserve_style(c1.paragraphs[1], f"{data.get('centro_saude_enderezo', '')} {data.get('centro_saude_distancia', '')}")

    c0 = table.rows[2].cells[0]
    if len(c0.paragraphs) > 1:
        set_paragraph_text_preserve_style(c0.paragraphs[1], data.get("tlf_hospital", "981 33 40 00"))
    c1 = table.rows[2].cells[1]
    if len(c1.paragraphs) > 1:
        set_paragraph_text_preserve_style(c1.paragraphs[0], f"{data.get('hospital_nome', '')},")
        set_paragraph_text_preserve_style(c1.paragraphs[1], f"{data.get('hospital_enderezo', '')}\t{data.get('hospital_distancia', '')}")

def update_table_5_and_machinery_details(doc, data):
    """Actualiza a Táboa 5 (casillas de maquinaria) e as táboas de detalle asociadas."""
    t5 = doc.tables[4]
    maquinaria_data = data.get("maquinaria", {})

    for m_key, m_info in MAQUINARIA_MAP.items():
        is_active = bool(maquinaria_data.get(m_key, False))
        mark = "X" if is_active else ""
        row = t5.rows[m_info["row"]]
        col_check = m_info["col_check"]
        set_cell_text_preserve_style(row.cells[col_check], mark)

        t_idx = m_info["table_idx"]
        if t_idx is not None and t_idx < len(doc.tables):
            det_tbl = doc.tables[t_idx]
            for r_i in range(2, len(det_tbl.rows)):
                det_row = det_tbl.rows[r_i]
                if len(det_row.cells) > 0:
                    txt = det_row.cells[0].text.strip()
                    if any(header in txt for header in ["MEDIDAS PREVENTIVAS", "GRADO DE ADOPCI", "EQUIPOS DE PROTECCI"]):
                        continue
                    set_cell_text_preserve_style(det_row.cells[0], mark)

def update_table_24_medios_auxiliares(table, data):
    """Actualiza a Táboa 24: Medios Auxiliares."""
    medios_data = data.get("medios_auxiliares", {})
    for key, info in MEDIOS_AUXILIARES_MAP.items():
        is_active = bool(medios_data.get(key, False))
        mark = "X" if is_active else ""
        row_idx = info["row"]
        if row_idx < len(table.rows):
            set_cell_text_preserve_style(table.rows[row_idx].cells[0], mark)

def update_table_25_riscos_evitables(table, data):
    """Actualiza a Táboa 25: Riscos Evitables."""
    riscos_data = data.get("riscos_evitables", {})
    rotura_activa = bool(riscos_data.get("rotura_instalacions", True))
    m1 = "X" if rotura_activa else ""
    set_cell_text_preserve_style(table.rows[1].cells[0], m1)
    if len(table.rows[1].cells) > 2:
        set_cell_text_preserve_style(table.rows[1].cells[2], m1)

    alta_tension = bool(riscos_data.get("linas_alta_tension", False))
    m2 = "X" if alta_tension else ""
    set_cell_text_preserve_style(table.rows[2].cells[0], m2)
    if len(table.rows[2].cells) > 2:
        set_cell_text_preserve_style(table.rows[2].cells[2], m2)

def update_signature_and_date(doc, data):
    """Localiza e actualiza a data, as firmas e o número de colexiado."""
    paragraphs = doc.paragraphs
    fdo_idx = None
    for i, p in enumerate(paragraphs):
        if p.text.strip().startswith("Fdo."):
            fdo_idx = i
            break

    if fdo_idx is not None:
        set_paragraph_text_preserve_style(paragraphs[fdo_idx], data.get("asinantes", ""))
        if fdo_idx + 1 < len(paragraphs):
            set_paragraph_text_preserve_style(paragraphs[fdo_idx + 1], data.get("colexiado", ""))
        for j in range(fdo_idx - 1, max(0, fdo_idx - 15), -1):
            if paragraphs[j].text.strip():
                set_paragraph_text_preserve_style(paragraphs[j], data.get("data_documento", ""))
                break

def update_provisional_installations_text(doc, data):
    """Actualiza o parágrafo explicativo sobre as instalacións provisionais."""
    necesarias = bool(data.get("instalacions_provisionais_necesarias", False))
    texto_usuario = data.get("instalacions_provisionais_texto", "").strip()

    if not texto_usuario:
        if necesarias:
            texto_usuario = "Estímanse necesarias instalacións provisionais para a presente obra, dispoñendo dos servizos hixiénicos e locais axeitados especificados na normativa."
        else:
            texto_usuario = "Non se estiman necesarias instalacións provisionais para a presente obra, debido as súas características e escasa duración."

    for p in doc.paragraphs:
        if "instalaci" in p.text.lower() and ("debido as s" in p.text.lower() or "caracter" in p.text.lower() or "necesarias instalaci" in p.text.lower()):
            set_paragraph_text_preserve_style(p, f"\t{texto_usuario}")
            break

def replace_placeholders_in_paragraph(p, data):
    """Reemplaza marcadores {{CLAVE}} no parágrafo mantendo os estilos do primeiro run."""
    full_text = p.text
    if "{{" not in full_text:
        return False
    
    matches = re.findall(r"\{\{([A-Za-z0-9_\.]+)\}\}", full_text)
    if not matches:
        return False

    new_text = full_text
    found_any = False
    for key in matches:
        val = None
        if key in data:
            val = str(data[key])
        elif "." in key:
            parent_key, child_key = key.split(".", 1)
            if parent_key in data and isinstance(data[parent_key], dict):
                sub_val = data[parent_key].get(child_key)
                if isinstance(sub_val, bool):
                    val = "X" if sub_val else ""
                elif sub_val is not None:
                    val = str(sub_val)

        if val is not None:
            new_text = new_text.replace(f"{{{{{key}}}}}", val)
            found_any = True

    if found_any:
        set_paragraph_text_preserve_style(p, new_text)
    return found_any

def replace_placeholders_in_doc(doc, data):
    """Reemplaza tódolos marcadores {{CLAVE}} en parágrafos e táboas do documento."""
    replaced_count = 0
    for p in doc.paragraphs:
        if replace_placeholders_in_paragraph(p, data):
            replaced_count += 1
            
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_placeholders_in_paragraph(p, data):
                        replaced_count += 1
    return replaced_count

def generate_ebss(data, template_path=None, output_path=None, output_stream=None):
    """
    Xera o documento final do EBSS aplicando tódalas modificacións.
    Se output_stream (io.BytesIO) se proporciona, garda directamente na memoria
    sen tocar o disco (ideal para servidores web e descargas directas).
    """
    default_tmpl, default_out_dir = get_default_paths()
    if template_path is None:
        template_path = default_tmpl

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Non se atopou o arquivo de plantilla DOCX en: {template_path}")

    filename = data.get("nome_ficheiro_saida", "ESTUDO_BASICO_SEGURIDADE_E_SAUDE_Modificado.docx")
    if not filename.lower().endswith(".docx"):
        filename += ".docx"

    if output_path is None and output_stream is None:
        out_dir = os.path.dirname(template_path) if template_path else default_out_dir
        output_path = os.path.join(out_dir, filename)

    # Abrir e actualizar documento
    doc = docx.Document(template_path)

    # 1. Intentar reemplazo por marcadores {{...}} (se existen na plantilla)
    placeholders_replaced = replace_placeholders_in_doc(doc, data)

    # 2. Reemplazo por mapa fixo de táboas (compatibilidade coa plantilla base)
    if len(doc.tables) > 0:
        update_table_1(doc.tables[0], data)
    if len(doc.tables) > 1:
        update_table_2(doc.tables[1], data)
    if len(doc.tables) > 2:
        update_table_3(doc.tables[2], data)
    update_provisional_installations_text(doc, data)
    if len(doc.tables) > 3:
        update_table_4(doc.tables[3], data)
    if len(doc.tables) > 4:
        update_table_5_and_machinery_details(doc, data)
    if len(doc.tables) > 23:
        update_table_24_medios_auxiliares(doc.tables[23], data)
    if len(doc.tables) > 24:
        update_table_25_riscos_evitables(doc.tables[24], data)
    update_signature_and_date(doc, data)

    if output_stream is not None:
        doc.save(output_stream)
        output_stream.seek(0)
        return {
            "status": "success",
            "filename": filename,
            "placeholders_replaced": placeholders_replaced
        }

    # Gardar con xestión de PermissionError (se Word ten o arquivo aberto)
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
        "placeholders_replaced": placeholders_replaced
    }

if __name__ == "__main__":
    print("Probando xerador con rutas portátiles...")
    res = generate_ebss(DEFAULT_DATA)
    print(f"Xerado con éxito: {res['output_path']} ({res['size_bytes']} bytes)")
