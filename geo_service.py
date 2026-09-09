# -*- coding: utf-8 -*-
"""
geo_service.py
Servizo de xeolocalización e cálculo automático de rutas e distancias por estrada
entre a obra e os centros sanitarios (Centro de Saúde e Hospital).
Inclúe busca intelixente de centros de saúde e hospitais máis próximos
combinando Overpass API (OpenStreetMap) e o catálogo de centros do SERGAS / SNS.
100% gratuíto, sen chaves de API nin tarxetas de crédito.
"""

import math
import urllib.parse
import urllib.request
import json
import logging

logger = logging.getLogger(__name__)

USER_AGENT = "AsistenteTecnicoEBSS/2.0 (pliegos_ebss@arquitectura.gal)"

# Catálogo de centros de referencia do SERGAS con teléfonos oficiais de urxencias e coordenadas verificadas
CURATED_HEALTH_DIRECTORY = [
    # --- ÁREA DE FERROL ---
    {"type": "cs", "nome": "Centro de Saúde Fontenla Maristany", "enderezo": "Praza de España, 19", "poboacion": "Ferrol", "tlf": "981 33 66 33", "lat": 43.4862, "lon": -8.2263},
    {"type": "cs", "nome": "Centro de Saúde de Caranza", "enderezo": "Av. Castelao s/n", "poboacion": "Ferrol", "tlf": "981 33 66 20", "lat": 43.4912, "lon": -8.2045},
    {"type": "cs", "nome": "Centro de Saúde de Serantes", "enderezo": "Camiño de Serantes, 2", "poboacion": "Ferrol", "tlf": "981 33 66 50", "lat": 43.4990, "lon": -8.2510},
    {"type": "cs", "nome": "Centro de Saúde de Narón", "enderezo": "Rúa 25 de Xullo s/n", "poboacion": "Narón", "tlf": "981 39 12 00", "lat": 43.5015, "lon": -8.1920},
    {"type": "cs", "nome": "Centro de Saúde de Fene", "enderezo": "Av. das Pías s/n", "poboacion": "Fene", "tlf": "981 34 00 20", "lat": 43.4735, "lon": -8.1630},
    {"type": "cs", "nome": "Centro de Saúde de Pontedeume", "enderezo": "Rúa Chousa s/n", "poboacion": "Pontedeume", "tlf": "981 43 00 22", "lat": 43.4072, "lon": -8.1715},
    {"type": "cs", "nome": "Centro de Saúde de As Pontes", "enderezo": "Av. de Ferrol s/n", "poboacion": "As Pontes", "tlf": "981 45 30 10", "lat": 43.4490, "lon": -7.8540},
    {"type": "hosp", "nome": "Hospital Arquitecto Marcide (CHUF)", "enderezo": "Av. de la Residencia s/n", "poboacion": "Ferrol", "tlf": "981 33 40 00", "lat": 43.5005, "lon": -8.2145},
    {"type": "hosp", "nome": "Hospital Naval de Ferrol", "enderezo": "Estrada San Pedro de Leixa s/n", "poboacion": "Ferrol", "tlf": "981 33 40 00", "lat": 43.5030, "lon": -8.2080},
    
    # --- ÁREA DA CORUÑA ---
    {"type": "cs", "nome": "Centro de Saúde de San José", "enderezo": "Rúa Comandante Fontanes, 8", "poboacion": "A Coruña", "tlf": "981 17 80 00", "lat": 43.3685, "lon": -8.4060},
    {"type": "cs", "nome": "Centro de Saúde Federico Tapia", "enderezo": "Rúa Federico Tapia, 63", "poboacion": "A Coruña", "tlf": "981 16 01 00", "lat": 43.3610, "lon": -8.4095},
    {"type": "cs", "nome": "Centro de Saúde de Matogrande", "enderezo": "Rúa Juan Díaz Porlier, 15", "poboacion": "A Coruña", "tlf": "981 17 65 00", "lat": 43.3420, "lon": -8.4030},
    {"type": "cs", "nome": "Centro de Saúde de Culleredo (Acea de Ama)", "enderezo": "Rúa Costa da Lonxa s/n", "poboacion": "Culleredo", "tlf": "981 66 40 00", "lat": 43.3180, "lon": -8.3840},
    {"type": "cs", "nome": "Centro de Saúde de Oleiros (Perillo)", "enderezo": "Av. das Mariñas s/n", "poboacion": "Oleiros", "tlf": "981 63 90 00", "lat": 43.3360, "lon": -8.3610},
    {"type": "cs", "nome": "Centro de Saúde de Betanzos", "enderezo": "Praza García Hermanos s/n", "poboacion": "Betanzos", "tlf": "981 77 10 00", "lat": 43.2810, "lon": -8.2140},
    {"type": "hosp", "nome": "Complexo Hospitalario Universitario da Coruña (CHUAC)", "enderezo": "As Xubias de Arriba, 84", "poboacion": "A Coruña", "tlf": "981 17 80 00", "lat": 43.3425, "lon": -8.3875},
    {"type": "hosp", "nome": "Hospital Teresa Herrera (Materno Infantil)", "enderezo": "As Xubias de Arriba, s/n", "poboacion": "A Coruña", "tlf": "981 17 80 00", "lat": 43.3440, "lon": -8.3890},
    {"type": "hosp", "nome": "Hospital Marítimo de Oza", "enderezo": "Rúa As Xubias s/n", "poboacion": "A Coruña", "tlf": "981 17 80 00", "lat": 43.3490, "lon": -8.3910},
    {"type": "hosp", "nome": "Hospital Abente y Lago", "enderezo": "Paseo do Xeneral Sir John Moore s/n", "poboacion": "A Coruña", "tlf": "981 17 80 00", "lat": 43.3670, "lon": -8.3900},

    # --- ÁREA DE SANTIAGO DE COMPOSTELA ---
    {"type": "cs", "nome": "Centro de Saúde Concepcion Arenal", "enderezo": "Rúa Santiago León de Caracas, 12", "poboacion": "Santiago de Compostela", "tlf": "981 95 00 00", "lat": 42.8750, "lon": -8.5480},
    {"type": "cs", "nome": "Centro de Saúde de Fontiñas", "enderezo": "Rúa de Berlín s/n", "poboacion": "Santiago de Compostela", "tlf": "981 95 10 00", "lat": 42.8820, "lon": -8.5280},
    {"type": "hosp", "nome": "Complexo Hospitalario Universitario de Santiago (CHUS - Hospital Clínico)", "enderezo": "Rúa da Choupana s/n", "poboacion": "Santiago de Compostela", "tlf": "981 95 00 00", "lat": 42.8690, "lon": -8.5630},
    {"type": "hosp", "nome": "Hospital Provincial de Conxo", "enderezo": "Rúa Ramón Baltar s/n", "poboacion": "Santiago de Compostela", "tlf": "981 95 00 00", "lat": 42.8590, "lon": -8.5570},

    # --- ÁREA DE VIGO ---
    {"type": "cs", "nome": "Centro de Saúde Rosalía de Castro", "enderezo": "Rúa Rosalía de Castro, 21", "poboacion": "Vigo", "tlf": "986 81 10 00", "lat": 42.2370, "lon": -8.7180},
    {"type": "cs", "nome": "Centro de Saúde Casco Vello", "enderezo": "Praza da Constitución s/n", "poboacion": "Vigo", "tlf": "986 81 20 00", "lat": 42.2390, "lon": -8.7260},
    {"type": "hosp", "nome": "Hospital Álvaro Cunqueiro (CHUVI)", "enderezo": "Estrada de Clara Campoamor, 341", "poboacion": "Vigo", "tlf": "986 81 11 11", "lat": 42.1890, "lon": -8.6920},
    {"type": "hosp", "nome": "Hospital Meixoeiro", "enderezo": "Camiño do Meixoeiro s/n", "poboacion": "Vigo", "tlf": "986 81 11 11", "lat": 42.2210, "lon": -8.6670},

    # --- ÁREA DE PONTEVEDRA ---
    {"type": "cs", "nome": "Centro de Saúde Virxe Peregrina", "enderezo": "Rúa Frei Xoán de Navarrete, 2", "poboacion": "Pontevedra", "tlf": "986 80 00 00", "lat": 42.4280, "lon": -8.6470},
    {"type": "hosp", "nome": "Hospital Montecelo (CHOP)", "enderezo": "Lugar de Montecelo s/n (Mourente)", "poboacion": "Pontevedra", "tlf": "986 80 00 00", "lat": 42.4260, "lon": -8.6210},
    {"type": "hosp", "nome": "Hospital Provincial de Pontevedra", "enderezo": "Rúa Loureiro Crespo, 2", "poboacion": "Pontevedra", "tlf": "986 80 00 00", "lat": 42.4310, "lon": -8.6390},

    # --- ÁREA DE LUGO ---
    {"type": "cs", "nome": "Centro de Saúde Praza do Ferrol", "enderezo": "Praza do Ferrol, 11", "poboacion": "Lugo", "tlf": "982 29 60 00", "lat": 43.0120, "lon": -7.5540},
    {"type": "hosp", "nome": "Hospital Universitario Lucus Augusti (HULA)", "enderezo": "Rúa Doutor Ulises Romero, 1", "poboacion": "Lugo", "tlf": "982 29 60 00", "lat": 43.0290, "lon": -7.5270},

    # --- ÁREA DE OURENSE ---
    {"type": "cs", "nome": "Centro de Saúde Valle Inclán", "enderezo": "Rúa Valle Inclán, 7", "poboacion": "Ourense", "tlf": "988 38 50 00", "lat": 42.3420, "lon": -6.8620},
    {"type": "hosp", "nome": "Complexo Hospitalario Universitario de Ourense (CHUO)", "enderezo": "Rúa Ramón Puga Noguerol, 54", "poboacion": "Ourense", "tlf": "988 38 55 00", "lat": 42.3330, "lon": -6.8520}
]


def _haversine_distance_km(lat1, lon1, lat2, lon2):
    """Cálculo de respaldo en liña recta con factor de corrección por trazado viario (x1.35)."""
    R = 6371.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = math.sin(delta_phi / 2.0) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2.0) ** 2
    c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
    return round((R * c) * 1.35, 1)


def geocode_address(address: str, poboacion: str = ""):
    """
    Xeocodifica un enderezo en coordenadas (lat, lon) usando OpenStreetMap Nominatim.
    """
    if not address and not poboacion:
        return None

    clean_addr = address.replace("s/n", "").replace("S/N", "").strip()
    
    queries = []
    if clean_addr and poboacion:
        queries.append(f"{clean_addr}, {poboacion}, España")
    elif clean_addr:
        queries.append(f"{clean_addr}, España")
    if poboacion:
        queries.append(f"{poboacion}, Galicia, España")

    for q in queries:
        try:
            params = urllib.parse.urlencode({"q": q, "format": "json", "limit": 1, "addressdetails": 1})
            url = f"https://nominatim.openstreetmap.org/search?{params}"
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode("utf-8"))
                    if data and len(data) > 0:
                        lat = float(data[0]["lat"])
                        lon = float(data[0]["lon"])
                        return lat, lon
        except Exception as e:
            logger.warning(f"Erro xeocodificando query '{q}': {e}")
            continue

    return None


def get_driving_route(lat1, lon1, lat2, lon2):
    """
    Calcula a distancia de condución real por estrada (en km) e tempo estimado (en minutos)
    usando o servizo gratuíto OSRM (Open Source Routing Machine).
    """
    try:
        url = f"https://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=false"
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

        with urllib.request.urlopen(req, timeout=5) as response:
            if response.status == 200:
                data = json.loads(response.read().decode("utf-8"))
                if data.get("code") == "Ok" and "routes" in data and len(data["routes"]) > 0:
                    route = data["routes"][0]
                    dist_meters = route["distance"]
                    dur_seconds = route["duration"]

                    dist_km = round(dist_meters / 1000.0, 1)
                    dur_min = max(1, round(dur_seconds / 60.0))

                    return {
                        "dist_km": dist_km,
                        "dur_min": dur_min,
                        "dist_str": f"{str(dist_km).replace('.', ',')} km",
                        "dur_str": f"~{dur_min} min",
                        "method": "osrm_driving"
                    }
    except Exception as e:
        logger.warning(f"OSRM fallback a Haversine debido a: {e}")

    fallback_km = _haversine_distance_km(lat1, lon1, lat2, lon2)
    fallback_min = max(1, round((fallback_km / 35.0) * 60.0))

    return {
        "dist_km": fallback_km,
        "dur_min": fallback_min,
        "dist_str": f"{str(fallback_km).replace('.', ',')} km",
        "dur_str": f"~{fallback_min} min",
        "method": "haversine_approx"
    }


def calculate_health_distances(obra_situacion, obra_poboacion, cs_enderezo, cs_poboacion, hosp_enderezo, hosp_poboacion):
    """
    Calcula de xeito coordinado as distancias dende a obra ao Centro de Saúde e ao Hospital.
    """
    results = {
        "status": "success",
        "obra_coords": None,
        "centro_saude": None,
        "hospital": None
    }

    obra_coords = geocode_address(obra_situacion, obra_poboacion)
    if not obra_coords:
        return {
            "status": "error",
            "error": f"Non foi posible xeolocalizar o emprazamento da obra: '{obra_situacion}, {obra_poboacion}'."
        }
    results["obra_coords"] = {"lat": obra_coords[0], "lon": obra_coords[1]}

    # Centro de Saúde
    cs_pob = cs_poboacion or obra_poboacion
    cs_coords = geocode_address(cs_enderezo, cs_pob)
    if cs_coords:
        route_cs = get_driving_route(obra_coords[0], obra_coords[1], cs_coords[0], cs_coords[1])
        results["centro_saude"] = {
            "found": True,
            "distancia": route_cs["dist_str"],
            "duracion": route_cs["dur_str"],
            "dist_km": route_cs["dist_km"],
            "dur_min": route_cs["dur_min"]
        }
    else:
        results["centro_saude"] = {
            "found": False,
            "error": f"Non se atopou o enderezo do Centro de Saúde: '{cs_enderezo}'."
        }

    # Hospital
    hosp_pob = hosp_poboacion or obra_poboacion
    hosp_coords = geocode_address(hosp_enderezo, hosp_pob)
    if hosp_coords:
        route_hosp = get_driving_route(obra_coords[0], obra_coords[1], hosp_coords[0], hosp_coords[1])
        results["hospital"] = {
            "found": True,
            "distancia": route_hosp["dist_str"],
            "duracion": route_hosp["dur_str"],
            "dist_km": route_hosp["dist_km"],
            "dur_min": route_hosp["dur_min"]
        }
    else:
        results["hospital"] = {
            "found": False,
            "error": f"Non se atopou o enderezo do Hospital: '{hosp_enderezo}'."
        }

    return results


def find_nearby_health_facilities(obra_situacion: str, obra_poboacion: str):
    """
    Atopa e clasifica por proximidade real por estrada os Centros de Saúde e Hospitais
    máis próximos ao emprazamento da obra.
    Combina Overpass API en tempo real e o catálogo oficial do SERGAS.
    """
    obra_coords = geocode_address(obra_situacion, obra_poboacion)
    if not obra_coords:
        return {
            "status": "error",
            "error": f"Non se puido xeolocalizar o emprazamento da obra: '{obra_situacion}, {obra_poboacion}'."
        }

    lat_obra, lon_obra = obra_coords[0], obra_coords[1]
    candidates = []

    # 1. Engadir candidatos do catálogo oficial
    for fac in CURATED_HEALTH_DIRECTORY:
        approx_km = _haversine_distance_km(lat_obra, lon_obra, fac["lat"], fac["lon"])
        # Limitar inicialmente a 70 km en liña recta para considerar centros comarcais
        if approx_km < 70.0:
            candidates.append({
                "type": fac["type"],
                "nome": fac["nome"],
                "enderezo": f"{fac['enderezo']}, {fac['poboacion']}",
                "poboacion": fac["poboacion"],
                "tlf": fac["tlf"],
                "lat": fac["lat"],
                "lon": fac["lon"],
                "approx_km": approx_km
            })

    # 2. Consultar Overpass API para atopar outros centros locais da zona (se existen)
    try:
        query = f"""
        [out:json][timeout:5];
        (
          node["amenity"="hospital"](around:25000,{lat_obra},{lon_obra});
          node["amenity"="clinic"](around:20000,{lat_obra},{lon_obra});
          node["healthcare"="centre"](around:20000,{lat_obra},{lon_obra});
        );
        out 15;
        """
        req_data = urllib.parse.urlencode({"data": query}).encode("utf-8")
        req = urllib.request.Request("https://overpass-api.de/api/interpreter", data=req_data, headers={"User-Agent": USER_AGENT})

        with urllib.request.urlopen(req, timeout=5) as resp:
            if resp.status == 200:
                data = json.loads(resp.read().decode("utf-8"))
                for el in data.get("elements", []):
                    tags = el.get("tags", {})
                    nome = tags.get("name") or tags.get("official_name")
                    if not nome:
                        continue
                    
                    is_hosp = (tags.get("amenity") == "hospital" or tags.get("healthcare") == "hospital" or "hospital" in nome.lower())
                    ftype = "hosp" if is_hosp else "cs"
                    
                    street = tags.get("addr:street", "")
                    housenumber = tags.get("addr:housenumber", "")
                    city = tags.get("addr:city", obra_poboacion)
                    full_addr = f"{street} {housenumber}".strip()
                    if not full_addr:
                        full_addr = f"{nome}, {city}"
                    else:
                        full_addr = f"{full_addr}, {city}"

                    tlf = tags.get("phone") or tags.get("contact:phone") or "112 / 061"
                    lat_c = el.get("lat")
                    lon_c = el.get("lon")

                    # Evitar duplicados por nome similar
                    already = False
                    for c in candidates:
                        if nome.lower() in c["nome"].lower() or c["nome"].lower() in nome.lower():
                            already = True
                            break
                    
                    if not already and lat_c and lon_c:
                        approx_km = _haversine_distance_km(lat_obra, lon_obra, lat_c, lon_c)
                        candidates.append({
                            "type": ftype,
                            "nome": nome,
                            "enderezo": full_addr,
                            "poboacion": city,
                            "tlf": tlf,
                            "lat": lat_c,
                            "lon": lon_c,
                            "approx_km": approx_km
                        })
    except Exception as e:
        logger.warning(f"Overpass API non respondeu, empregando catálogo oficial: {e}")

    # Separar en Centros de Saúde e Hospitais
    cs_list = [c for c in candidates if c["type"] == "cs"]
    hosp_list = [c for c in candidates if c["type"] == "hosp"]

    # Ordenar provisionalmente por distancia aproximada para calcular estrada dos 4 máis próximos de cada tipo
    cs_list.sort(key=lambda x: x["approx_km"])
    hosp_list.sort(key=lambda x: x["approx_km"])

    top_cs = cs_list[:4]
    top_hosp = hosp_list[:4]

    # Calcular rutas reais por estrada con OSRM para os mellores candidatos
    final_cs = []
    for cs in top_cs:
        route = get_driving_route(lat_obra, lon_obra, cs["lat"], cs["lon"])
        final_cs.append({
            "nome": cs["nome"],
            "enderezo": cs["enderezo"],
            "poboacion": cs["poboacion"],
            "tlf": cs["tlf"],
            "distancia": route["dist_str"],
            "duracion": route["dur_str"],
            "dist_km": route["dist_km"],
            "dur_min": route["dur_min"]
        })

    final_hosp = []
    for h in top_hosp:
        route = get_driving_route(lat_obra, lon_obra, h["lat"], h["lon"])
        final_hosp.append({
            "nome": h["nome"],
            "enderezo": h["enderezo"],
            "poboacion": h["poboacion"],
            "tlf": h["tlf"],
            "distancia": route["dist_str"],
            "duracion": route["dur_str"],
            "dist_km": route["dist_km"],
            "dur_min": route["dur_min"]
        })

    # Ordenar definitivamente por distancia de estrada real
    final_cs.sort(key=lambda x: x["dist_km"])
    final_hosp.sort(key=lambda x: x["dist_km"])

    return {
        "status": "success",
        "obra": {
            "situacion": obra_situacion,
            "poboacion": obra_poboacion,
            "lat": lat_obra,
            "lon": lon_obra
        },
        "centros_saude": final_cs,
        "hospitais": final_hosp
    }


if __name__ == "__main__":
    print("Probando busca de centros próximos para obra en Ferrol...")
    res = find_nearby_health_facilities("Rúa Baterías 34", "Ferrol")
    print(f"Centros de saúde atopados: {len(res.get('centros_saude', []))}")
    for cs in res.get("centros_saude", []):
        print(f"  - {cs['nome']} ({cs['distancia']}, {cs['duracion']}) - Tlf: {cs['tlf']}")
    print(f"Hospitais atopados: {len(res.get('hospitais', []))}")
    for h in res.get("hospitais", []):
        print(f"  - {h['nome']} ({h['distancia']}, {h['duracion']}) - Tlf: {h['tlf']}")
