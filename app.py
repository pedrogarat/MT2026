# -*- coding: utf-8 -*-
"""
app.py
Aplicación Web Flask para o Asistente Técnico de Proxectos (EBSS & Xestión de Residuos).
Inclúe soporte multiusuario, base de datos relacional (SQLite local / PostgreSQL nube),
xestión de proxectos con datos comúns compartidos e xeración modular independente de .docx.
"""

import os
import sys
import io
import json
import socket
import webbrowser
from datetime import datetime
from functools import wraps
from flask import (
    Flask, render_template, request, jsonify, session,
    redirect, url_for, send_file, abort
)

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from models import init_db, SessionLocal, User, Project
from ebss_processor import DEFAULT_DATA, generate_ebss, get_default_paths
from residuos_processor import (
    DEFAULT_RESIDUOS_DATA, generate_residuos, calculate_residuos,
    get_default_residuos_paths
)
from geo_service import calculate_health_distances, find_nearby_health_facilities

app = Flask(__name__, template_folder=os.path.join(CURRENT_DIR, "templates"))
app.secret_key = os.environ.get("SECRET_KEY", "asistente-tecnico-ebss-residuos-secret-key-2026")

# Inicializar táboas da base de datos ao arrancar
init_db()


# ---------------------------------------------------------
# Utilidades e Decoradores de Autenticación
# ---------------------------------------------------------
def get_db_session():
    return SessionLocal()


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            if request.is_json or request.path.startswith("/api/"):
                return jsonify({"status": "error", "error": "Non autenticado. Inicia sesión primeiro."}), 401
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated_function


def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    db = get_db_session()
    try:
        return db.query(User).filter(User.id == user_id).first()
    finally:
        db.close()


# ---------------------------------------------------------
# Rutas de Páxinas HTML
# ---------------------------------------------------------
@app.route("/")
def index():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login_page"))


@app.route("/login")
def login_page():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/dashboard")
@login_required
def dashboard():
    return render_template("dashboard.html")


@app.route("/project/<int:project_id>")
@login_required
def project_editor(project_id):
    db = get_db_session()
    try:
        project = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not project:
            abort(404, "Proxecto non atopado ou sen permisos.")
        return render_template("project_editor.html")
    finally:
        db.close()


# Rutas de compatibilidade directa con versións anteriores
@app.route("/home")
def legacy_home():
    return redirect(url_for("dashboard"))


@app.route("/ebss")
def legacy_ebss():
    return render_template("index.html")


@app.route("/residuos")
def legacy_residuos():
    return render_template("residuos.html")


# ---------------------------------------------------------
# API de Autenticación
# ---------------------------------------------------------
@app.route("/api/auth/register", methods=["POST"])
def auth_register():
    data = request.get_json() or {}
    username = data.get("username", "").strip()
    email = data.get("email", "").strip() or None
    password = data.get("password", "")

    if not username or not password:
        return jsonify({"status": "error", "error": "Usuario e contrasinal son obrigatorios."}), 400

    if len(password) < 4:
        return jsonify({"status": "error", "error": "O contrasinal debe ter polo menos 4 caracteres."}), 400

    db = get_db_session()
    try:
        existing = db.query(User).filter(User.username == username).first()
        if existing:
            return jsonify({"status": "error", "error": "O nome de usuario xa está en uso."}), 409

        user = User(username=username, email=email)
        user.set_password(password)
        db.add(user)
        db.commit()

        # Crear automáticamente un proxecto inicial de benvida
        default_common = {
            "tipo_obra": DEFAULT_DATA.get("tipo_obra", "Melloras de condicións de seguridade, salubridade e ornato"),
            "situacion": DEFAULT_DATA.get("situacion", "Rúa Baterías 34"),
            "poboacion": DEFAULT_DATA.get("poboacion", "Ferrol"),
            "promotor": DEFAULT_DATA.get("promotor", "El Bosque Providencia S.L."),
            "arquitecto": DEFAULT_DATA.get("arquitecto", "Estudio Anta Arquitectos S.L.P."),
            "coordinador_ss": DEFAULT_DATA.get("coordinador_ss", "A determinar"),
            "orzamento_pem": DEFAULT_DATA.get("orzamento_pem", "O indicado no capítulo homónimo do presente proxecto"),
            "duracion_obra": DEFAULT_DATA.get("duracion_obra", "3 meses"),
            "num_traballadores": DEFAULT_DATA.get("num_traballadores", "6"),
            "superficie_construida": DEFAULT_RESIDUOS_DATA.get("superficie_construida", "9,84 m2"),
            "superficie_actuacion": DEFAULT_RESIDUOS_DATA.get("superficie_actuacion", 9.84),
            "data_documento": "Ferrol, xullo de 2026",
            "asinantes": "Fdo. Sergio J. Beceiro Lodeiro     M. Rosa Vilas Romalde",
            "colexiado": "ESTUDIO ANTA ARQUITECTOS S.L.P.\tNº COAG - 20.039"
        }

        default_ebss = {
            "accesos_obra": DEFAULT_DATA.get("accesos_obra"),
            "topografia_terreo": DEFAULT_DATA.get("topografia_terreo"),
            "tipo_chan": DEFAULT_DATA.get("tipo_chan"),
            "edificacions_lindeiras": DEFAULT_DATA.get("edificacions_lindeiras"),
            "fase_demolicions": DEFAULT_DATA.get("fase_demolicions"),
            "fase_terras": DEFAULT_DATA.get("fase_terras"),
            "fase_estruturas": DEFAULT_DATA.get("fase_estruturas"),
            "fase_cubertas": DEFAULT_DATA.get("fase_cubertas"),
            "fase_albanelaria": DEFAULT_DATA.get("fase_albanelaria"),
            "fase_acabados": DEFAULT_DATA.get("fase_acabados"),
            "fase_instalacions": DEFAULT_DATA.get("fase_instalacions"),
            "centro_saude_nome": DEFAULT_DATA.get("centro_saude_nome"),
            "centro_saude_enderezo": DEFAULT_DATA.get("centro_saude_enderezo"),
            "tlf_urxencias": DEFAULT_DATA.get("tlf_urxencias"),
            "centro_saude_distancia": DEFAULT_DATA.get("centro_saude_distancia"),
            "hospital_nome": DEFAULT_DATA.get("hospital_nome"),
            "hospital_enderezo": DEFAULT_DATA.get("hospital_enderezo"),
            "tlf_hospital": DEFAULT_DATA.get("tlf_hospital"),
            "hospital_distancia": DEFAULT_DATA.get("hospital_distancia"),
            "maquinaria": DEFAULT_DATA.get("maquinaria", {})
        }

        default_residuos = {
            "residuos_selection": DEFAULT_RESIDUOS_DATA.get("residuos_selection", {})
        }

        proj = Project(
            user_id=user.id,
            name="Exemplo: Proxecto Ferrol",
            description="Proxecto inicial preconfigurado con datos comúns e módulos técnicos.",
            common_data=json.dumps(default_common, ensure_ascii=False),
            ebss_data=json.dumps(default_ebss, ensure_ascii=False),
            residuos_data=json.dumps(default_residuos, ensure_ascii=False)
        )
        db.add(proj)
        db.commit()

        session["user_id"] = user.id
        session["username"] = user.username

        return jsonify({"status": "success", "user": user.to_dict()})
    finally:
        db.close()


@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    data = request.get_json() or {}
    username_or_email = data.get("username", "").strip()
    password = data.get("password", "")

    if not username_or_email or not password:
        return jsonify({"status": "error", "error": "Indica usuario e contrasinal."}), 400

    db = get_db_session()
    try:
        user = db.query(User).filter(
            (User.username == username_or_email) | (User.email == username_or_email)
        ).first()

        if not user or not user.check_password(password):
            return jsonify({"status": "error", "error": "Credenciais incorrectas."}), 401

        session["user_id"] = user.id
        session["username"] = user.username

        return jsonify({"status": "success", "user": user.to_dict()})
    finally:
        db.close()


@app.route("/api/auth/logout", methods=["POST", "GET"])
def auth_logout():
    session.clear()
    if request.is_json or request.path.startswith("/api/"):
        return jsonify({"status": "success"})
    return redirect(url_for("login_page"))


# ---------------------------------------------------------
# API de Xestión de Proxectos
# ---------------------------------------------------------
@app.route("/api/projects", methods=["GET"])
@login_required
def list_projects():
    db = get_db_session()
    try:
        user = db.query(User).filter(User.id == session["user_id"]).first()
        if not user:
            session.clear()
            return jsonify({"status": "error", "error": "Usuario non atopado"}), 401

        projects_query = db.query(Project).filter(Project.user_id == user.id).order_by(Project.updated_at.desc()).all()
        projects = [p.to_dict() for p in projects_query]
        
        resp = jsonify({
            "status": "success",
            "projects": projects,
            "user": user.to_dict()
        })
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        return resp
    finally:
        db.close()


@app.route("/api/projects", methods=["POST"])
@login_required
def create_project():
    data = request.get_json() or {}
    name = data.get("name", "").strip()
    if not name:
        return jsonify({"status": "error", "error": "O nome do proxecto é obrigatorio."}), 400

    db = get_db_session()
    try:
        # Valores por defecto para novo proxecto
        common = {
            "tipo_obra": data.get("common_data", {}).get("tipo_obra", "Melloras de condicións de seguridade, salubridade e ornato"),
            "situacion": data.get("common_data", {}).get("situacion", "Rúa Baterías 34"),
            "poboacion": data.get("common_data", {}).get("poboacion", "Ferrol"),
            "promotor": data.get("common_data", {}).get("promotor", "El Bosque Providencia S.L."),
            "arquitecto": data.get("common_data", {}).get("arquitecto", "Estudio Anta Arquitectos S.L.P."),
            "coordinador_ss": data.get("common_data", {}).get("coordinador_ss", "A determinar"),
            "orzamento_pem": data.get("common_data", {}).get("orzamento_pem", "O indicado no capítulo homónimo do presente proxecto"),
            "duracion_obra": data.get("common_data", {}).get("duracion_obra", "3 meses"),
            "num_traballadores": data.get("common_data", {}).get("num_traballadores", "6"),
            "superficie_construida": data.get("common_data", {}).get("superficie_construida", "9,84 m2"),
            "superficie_actuacion": data.get("common_data", {}).get("superficie_actuacion", 9.84),
            "data_documento": data.get("common_data", {}).get("data_documento", "Ferrol, xullo de 2026"),
            "asinantes": data.get("common_data", {}).get("asinantes", "Fdo. Sergio J. Beceiro Lodeiro     M. Rosa Vilas Romalde"),
            "colexiado": data.get("common_data", {}).get("colexiado", "ESTUDIO ANTA ARQUITECTOS S.L.P.\tNº COAG - 20.039")
        }
        if "common_data" in data and isinstance(data["common_data"], dict):
            common.update(data["common_data"])

        ebss = {
            "accesos_obra": DEFAULT_DATA.get("accesos_obra"),
            "topografia_terreo": DEFAULT_DATA.get("topografia_terreo"),
            "tipo_chan": DEFAULT_DATA.get("tipo_chan"),
            "edificacions_lindeiras": DEFAULT_DATA.get("edificacions_lindeiras"),
            "fase_demolicions": DEFAULT_DATA.get("fase_demolicions"),
            "fase_terras": DEFAULT_DATA.get("fase_terras"),
            "fase_estruturas": DEFAULT_DATA.get("fase_estruturas"),
            "fase_cubertas": DEFAULT_DATA.get("fase_cubertas"),
            "fase_albanelaria": DEFAULT_DATA.get("fase_albanelaria"),
            "fase_acabados": DEFAULT_DATA.get("fase_acabados"),
            "fase_instalacions": DEFAULT_DATA.get("fase_instalacions"),
            "centro_saude_nome": DEFAULT_DATA.get("centro_saude_nome"),
            "centro_saude_enderezo": DEFAULT_DATA.get("centro_saude_enderezo"),
            "tlf_urxencias": DEFAULT_DATA.get("tlf_urxencias"),
            "centro_saude_distancia": DEFAULT_DATA.get("centro_saude_distancia"),
            "hospital_nome": DEFAULT_DATA.get("hospital_nome"),
            "hospital_enderezo": DEFAULT_DATA.get("hospital_enderezo"),
            "tlf_hospital": DEFAULT_DATA.get("tlf_hospital"),
            "hospital_distancia": DEFAULT_DATA.get("hospital_distancia"),
            "maquinaria": DEFAULT_DATA.get("maquinaria", {})
        }
        if "ebss_data" in data and isinstance(data["ebss_data"], dict):
            ebss.update(data["ebss_data"])

        residuos = {
            "residuos_selection": DEFAULT_RESIDUOS_DATA.get("residuos_selection", {})
        }
        if "residuos_data" in data and isinstance(data["residuos_data"], dict):
            residuos.update(data["residuos_data"])

        proj = Project(
            user_id=session["user_id"],
            name=name,
            description=data.get("description", ""),
            common_data=json.dumps(common, ensure_ascii=False),
            ebss_data=json.dumps(ebss, ensure_ascii=False),
            residuos_data=json.dumps(residuos, ensure_ascii=False)
        )
        db.add(proj)
        db.commit()

        return jsonify({"status": "success", "project": proj.to_dict()})
    finally:
        db.close()


@app.route("/api/projects/<int:project_id>", methods=["GET"])
@login_required
def get_project(project_id):
    db = get_db_session()
    try:
        proj = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not proj:
            return jsonify({"status": "error", "error": "Proxecto non atopado."}), 404
        return jsonify({"status": "success", "project": proj.to_dict()})
    finally:
        db.close()


@app.route("/api/projects/<int:project_id>", methods=["PUT"])
@login_required
def update_project(project_id):
    data = request.get_json() or {}
    db = get_db_session()
    try:
        proj = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not proj:
            return jsonify({"status": "error", "error": "Proxecto non atopado."}), 404

        if "name" in data:
            proj.name = data["name"].strip() or proj.name
        if "description" in data:
            proj.description = data["description"]
        if "common_data" in data:
            proj.set_common_dict(data["common_data"])
        if "ebss_data" in data:
            proj.set_ebss_dict(data["ebss_data"])
        if "residuos_data" in data:
            proj.set_residuos_dict(data["residuos_data"])

        proj.updated_at = datetime.utcnow()
        db.commit()
        return jsonify({"status": "success", "project": proj.to_dict()})
    finally:
        db.close()


@app.route("/api/projects/<int:project_id>", methods=["DELETE"])
@login_required
def delete_project(project_id):
    db = get_db_session()
    try:
        proj = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not proj:
            return jsonify({"status": "error", "error": "Proxecto non atopado."}), 404

        db.delete(proj)
        db.commit()
        return jsonify({"status": "success"})
    finally:
        db.close()


# ---------------------------------------------------------
# XERACIÓN MODULAR E INDEPENDENTE DE DOCUMENTOS
# ---------------------------------------------------------
@app.route("/api/projects/<int:project_id>/generate-ebss", methods=["POST"])
@login_required
def project_generate_ebss(project_id):
    db = get_db_session()
    try:
        proj = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not proj:
            return jsonify({"status": "error", "error": "Proxecto non atopado."}), 404

        # Combinar datos comúns con datos específicos de EBSS
        common = proj.get_common_dict()
        ebss = proj.get_ebss_dict()

        merged_data = dict(DEFAULT_DATA)
        merged_data.update(common)
        merged_data.update(ebss)

        # Xerar directamente na memoria (BytesIO)
        stream = io.BytesIO()
        pobl = common.get("poboacion", "Proxecto").strip() or "Proxecto"
        filename = f"EBSS_{pobl.replace(' ', '_')}.docx"
        merged_data["nome_ficheiro_saida"] = filename

        generate_ebss(merged_data, output_stream=stream)

        return send_file(
            stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500
    finally:
        db.close()


@app.route("/api/projects/<int:project_id>/generate-residuos", methods=["POST"])
@login_required
def project_generate_residuos(project_id):
    db = get_db_session()
    try:
        proj = db.query(Project).filter(
            Project.id == project_id,
            Project.user_id == session["user_id"]
        ).first()
        if not proj:
            return jsonify({"status": "error", "error": "Proxecto non atopado."}), 404

        # Combinar datos comúns con datos específicos de Residuos
        common = proj.get_common_dict()
        residuos = proj.get_residuos_dict()

        merged_data = dict(DEFAULT_RESIDUOS_DATA)
        merged_data.update(common)
        merged_data.update(residuos)

        # Xerar directamente na memoria (BytesIO)
        stream = io.BytesIO()
        pobl = common.get("poboacion", "Proxecto").strip() or "Proxecto"
        filename = f"XESTION_RESIDUOS_{pobl.replace(' ', '_')}.docx"
        merged_data["nome_ficheiro_saida"] = filename

        generate_residuos(merged_data, output_stream=stream)

        return send_file(
            stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500
    finally:
        db.close()


# ---------------------------------------------------------
# APIs Auxiliares e de Cálculo
# ---------------------------------------------------------
@app.route("/api/residuos/calculate", methods=["POST"])
def api_residuos_calculate():
    data = request.get_json() or {}
    try:
        calc = calculate_residuos(data)
        return jsonify(calc)
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/api/geo/calculate-distances", methods=["POST"])
def api_geo_calculate_distances():
    data = request.get_json() or {}
    obra_situacion = data.get("obra_situacion", "").strip()
    obra_poboacion = data.get("obra_poboacion", "").strip()
    cs_enderezo = data.get("centro_saude_enderezo", "").strip()
    cs_poboacion = data.get("centro_saude_poboacion", "").strip() or obra_poboacion
    hosp_enderezo = data.get("hospital_enderezo", "").strip()
    hosp_poboacion = data.get("hospital_poboacion", "").strip() or obra_poboacion

    if not obra_situacion and not obra_poboacion:
        return jsonify({"status": "error", "error": "Indica o emprazamento ou concello da obra nos Datos Comúns."}), 400

    try:
        res = calculate_health_distances(
            obra_situacion=obra_situacion,
            obra_poboacion=obra_poboacion,
            cs_enderezo=cs_enderezo,
            cs_poboacion=cs_poboacion,
            hosp_enderezo=hosp_enderezo,
            hosp_poboacion=hosp_poboacion
        )
        return jsonify(res)
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/api/geo/nearby-facilities", methods=["POST"])
def api_geo_nearby_facilities():
    data = request.get_json() or {}
    obra_situacion = data.get("obra_situacion", "").strip()
    obra_poboacion = data.get("obra_poboacion", "").strip()

    if not obra_situacion and not obra_poboacion:
        return jsonify({"status": "error", "error": "Indica a situación ou concello da obra nos Datos Comúns."}), 400

    try:
        res = find_nearby_health_facilities(obra_situacion, obra_poboacion)
        return jsonify(res)
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


@app.route("/api/defaults", methods=["GET"])
def api_defaults():
    return jsonify(DEFAULT_DATA)


@app.route("/api/residuos/defaults", methods=["GET"])
def api_residuos_defaults():
    return jsonify(DEFAULT_RESIDUOS_DATA)


# ---------------------------------------------------------
# Servidor e Execución Local
# ---------------------------------------------------------
def find_available_port(start_port=8080, max_tries=20):
    for port in range(start_port, start_port + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start_port


def start_server(port=None, auto_open=True):
    if port is None:
        port = find_available_port(8080)

    url = f"http://127.0.0.1:{port}"
    print("=" * 65, flush=True)
    print("  ASISTENTE TÉCNICO DE PROXECTOS (EBSS + XESTIÓN DE RESIDUOS)", flush=True)
    print("  SISTEMA MULTIUSUARIO CON BASE DE DATOS E PERSISTENCIA DE PROXECTOS", flush=True)
    print("=" * 65, flush=True)
    print(f" Servidor activo en: {url}", flush=True)
    print(" Preme Ctrl+C no terminal para pechar o asistente.", flush=True)
    print("=" * 65, flush=True)

    if auto_open:
        try:
            opened = webbrowser.open(url)
            if not opened and sys.platform.startswith("win"):
                os.system(f'start "" "{url}"')
        except Exception:
            pass

    app.run(host="0.0.0.0", port=port, debug=False)


if __name__ == "__main__":
    start_server()
