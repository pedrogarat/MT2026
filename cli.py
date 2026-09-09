# -*- coding: utf-8 -*-
"""
cli.py
Modo consola / terminal interactivo para o Asistente do EBSS.
Fai preguntas paso a paso en galego, permitindo premer Intro para aceptar os valores por defecto.
"""

import sys
import copy
from ebss_processor import DEFAULT_DATA, generate_ebss

def prompt_val(question, default):
    val = input(f"{question} [{default}]: ").strip()
    return val if val else default

def prompt_bool(question, default):
    def_str = "S" if default else "N"
    val = input(f"{question} (S/N) [{def_str}]: ").strip().upper()
    if not val:
        return default
    return val in ["S", "SI", "Y", "YES", "1"]

def run_cli():
    print("=" * 65)
    print(" 🛡️  CUESTIONARIO EBSS (MODO CONSOLA)")
    print(" Preme [Intro] en calquera pregunta para manter o valor por defecto.")
    print("=" * 65)

    data = copy.deepcopy(DEFAULT_DATA)

    # 1. Datos Xerais
    print("\n--- PASO 1: DATOS XERAIS DO PROXECTO ---")
    data["tipo_obra"] = prompt_val("Tipo de obra", data["tipo_obra"])
    data["situacion"] = prompt_val("Situación / Rúa", data["situacion"])
    data["poboacion"] = prompt_val("Poboación / Concello", data["poboacion"])
    data["promotor"] = prompt_val("Promotor / Cliente", data["promotor"])
    data["arquitecto"] = prompt_val("Arquitecto / Redactor", data["arquitecto"])
    data["coordinador_ss"] = prompt_val("Coordinador Seguridade e Saúde", data["coordinador_ss"])
    data["orzamento_pem"] = prompt_val("Orzamento PEM", data["orzamento_pem"])
    data["duracion_obra"] = prompt_val("Duración da obra", data["duracion_obra"])
    data["num_traballadores"] = prompt_val("Nº máximo traballadores", data["num_traballadores"])

    # 2. Emprazamento
    print("\n--- PASO 2: CONDICIONANTES DO EMPRAZAMENTO ---")
    data["accesos_obra"] = prompt_val("Accesos á obra", data["accesos_obra"])
    data["topografia_terreo"] = prompt_val("Topografía do terreo", data["topografia_terreo"])
    data["tipo_chan"] = prompt_val("Tipo de chan", data["tipo_chan"])
    data["edificacions_lindeiras"] = "Si" if prompt_bool("¿Edificacións lindeiras?", data["edificacions_lindeiras"] == "Si") else "Non"
    data["subministracion_electrica"] = "Si" if prompt_bool("¿Subministración eléctrica?", data["subministracion_electrica"] == "Si") else "Non"
    data["subministracion_auga"] = "Si" if prompt_bool("¿Subministración de auga?", data["subministracion_auga"] == "Si") else "Non"
    data["sistema_saneamento"] = "Si" if prompt_bool("¿Sistema de saneamento?", data["sistema_saneamento"] == "Si") else "Non"

    # 3. Fases da obra
    print("\n--- PASO 3: FASES DA OBRA ---")
    data["fase_demolicions"] = prompt_val("Demolicións (Si / Non / Si (puntuais))", data["fase_demolicions"])
    data["fase_terras"] = "Si" if prompt_bool("¿Movemento de terras?", data["fase_terras"] == "Si") else "Non"
    data["fase_estruturas"] = "Si" if prompt_bool("¿Cimentación e estruturas?", data["fase_estruturas"] == "Si") else "Non"
    data["fase_cubertas"] = "Si" if prompt_bool("¿Cubertas?", data["fase_cubertas"] == "Si") else "Non"
    data["fase_albanelaria"] = "Si" if prompt_bool("¿Albanelaría e cerramentos?", data["fase_albanelaria"] == "Si") else "Non"
    data["fase_acabados"] = "Si" if prompt_bool("¿Acabados?", data["fase_acabados"] == "Si") else "Non"
    data["fase_instalacions"] = "Si" if prompt_bool("¿Instalacións?", data["fase_instalacions"] == "Si") else "Non"

    # 4. Asistencia Sanitaria
    print("\n--- PASO 4: ASISTENCIA SANITARIA ---")
    data["centro_saude_nome"] = prompt_val("Centro de Saúde (Urxencias)", data["centro_saude_nome"])
    data["centro_saude_enderezo"] = prompt_val("Enderezo Centro Saúde", data["centro_saude_enderezo"])
    data["tlf_urxencias"] = prompt_val("Teléfono Urxencias", data["tlf_urxencias"])
    data["centro_saude_distancia"] = prompt_val("Distancia Centro Saúde", data["centro_saude_distancia"])

    data["hospital_nome"] = prompt_val("Hospital de referencia", data["hospital_nome"])
    data["hospital_enderezo"] = prompt_val("Enderezo Hospital", data["hospital_enderezo"])
    data["tlf_hospital"] = prompt_val("Teléfono Hospital", data["tlf_hospital"])
    data["hospital_distancia"] = prompt_val("Distancia Hospital", data["hospital_distancia"])

    # 5. Data e Nome de Saída
    print("\n--- PASO 5: PECHE E ASINATURAS ---")
    data["data_documento"] = prompt_val("Localidade e Data", data["data_documento"])
    data["asinantes"] = prompt_val("Técnicos asinantes", data["asinantes"])
    data["colexiado"] = prompt_val("Colexiado / Entidade", data["colexiado"])
    data["nome_ficheiro_saida"] = prompt_val("Nome do ficheiro .docx resultante", data["nome_ficheiro_saida"])

    print("\n⏳ Xerando documento...")
    try:
        res = generate_ebss(data)
        print("=" * 65)
        print(f"✅ DOCUMENTO XERADO CON ÉXITO:")
        print(f"   Ruta: {res['output_path']}")
        print(f"   Tamaño: {res['size_bytes'] / 1024:.1f} KB")
        print("=" * 65)
    except Exception as e:
        print(f"❌ Erro ó xerar o documento: {e}")

if __name__ == "__main__":
    run_cli()
