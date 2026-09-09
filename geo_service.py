# -*- coding: utf-8 -*-
"""
geo_service.py
Servizo de xeolocalización e cálculo automático de rutas e distancias por estrada
entre a obra e os centros sanitarios (Centro de Saúde e Hospital).
Utiliza OpenStreetMap (Nominatim) e OSRM (Open Source Routing Machine).
100% gratuíto, sen requirir chaves de API nin tarxetas de crédito.
"""

import math
import urllib.parse
import urllib.request
import json
import logging

logger = logging.getLogger(__name__)

USER_AGENT = "AsistenteTecnicoEBSS/2.0 (pliegos_ebss@arquitectura.gal)"


def _haversine_distance_km(lat1, lon1, lat2, lon2):
    """Cálculo de respaldo en liña recta con factor de corrección por trazado viario (x1.35)."""
    R = 6371.0  # Radio da terra en km
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = math.sin(delta_phi / 2.0) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2.0) ** 2
    c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
    straight_line = R * c
    return round(straight_line * 1.35, 1)


def geocode_address(address: str, poboacion: str = ""):
    """
    Xeocodifica un enderezo en coordenadas (lat, lon) usando OpenStreetMap Nominatim.
    Aplica fallback gradual se a rúa exacta non se atopa.
    """
    if not address and not poboacion:
        return None

    # Limpeza básica do enderezo
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

    # Fallback matemático en caso de microcorte ou tempo de espera
    fallback_km = _haversine_distance_km(lat1, lon1, lat2, lon2)
    fallback_min = max(1, round((fallback_km / 35.0) * 60.0))  # Asumindo 35 km/h velocidade media urbana

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

    # 1. Obter coordenadas da obra
    obra_coords = geocode_address(obra_situacion, obra_poboacion)
    if not obra_coords:
        return {
            "status": "error",
            "error": f"Non foi posible xeolocalizar o emprazamento da obra: '{obra_situacion}, {obra_poboacion}'."
        }
    results["obra_coords"] = {"lat": obra_coords[0], "lon": obra_coords[1]}

    # 2. Distancia ao Centro de Saúde
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

    # 3. Distancia ao Hospital
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


if __name__ == "__main__":
    print("Probando cálculo de distancias dende a obra en Ferrol...")
    res = calculate_health_distances(
        obra_situacion="Rúa Baterías 34",
        obra_poboacion="Ferrol",
        cs_enderezo="Praza de España 19",
        cs_poboacion="Ferrol",
        hosp_enderezo="Av. de la Residencia s/n",
        hosp_poboacion="Ferrol"
    )
    print("Resultado:")
    print(json.dumps(res, indent=2, ensure_ascii=False))
