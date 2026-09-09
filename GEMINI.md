# MEMORIA TÉCNICA E CONTEXTO DO PROXECTO: ASISTENTE EBSS + XESTIÓN DE RESIDUOS (MT2026)

Este documento contén o contexto persistente, arquitectura, credenciais de servizos e regras de desenvolvemento para calquera axente de intelixencia artificial que traballe neste repositorio.

---

## 1. Descrición do Proxecto
Aplicación web multiusuario orientada a técnicos, arquitectos e enxeñeiros en Galicia para a redacción técnica automatizada de:
1. **Estudo Básico de Seguridade e Saúde (EBSS)** segundo RD 1627/1997.
2. **Estudo de Xestión de Residuos de Construcción e Demolición (EGR)** segundo RD 105/2008 e normativa galega.

- **Idioma de traballo**: Galego en interface e documentos xerados en Word (`.docx`).
- **Formato de texto**: Caracteres e unidades estándar UTF-8 (`m²`, `m³`, `t`, `× 0,2`). **Nunca empregar sintaxe LaTeX** como `$m^2$` nin similares.

---

## 2. Arquitectura do Sistema

### Backend e Lóxica (Python Flask)
- **`app.py`**: Servidor Flask con API REST, autenticación por sesión (`/api/auth/register`, `/login`, `/logout`), xestión de proxectos CRUD e descarga directa de documentos.
- **`models.py`**: Modelos SQLAlchemy (`User`, `Project`). Xestión de persistencia híbrida (PostgreSQL en produción / SQLite en local).
- **`ebss_processor.py`**: Procesamento e xeración en memoria (`io.BytesIO`) da memoria e prego de seguridade e saúde.
- **`residuos_processor.py`**: Estimación de volumes, pesos por codificación LER e xeración en memoria da memoria de xestión de residuos.
- **`geo_service.py`**: Xeolocalización aberta (OpenStreetMap / OSRM) e directorio de centros de saúde e complexos hospitalarios do SERGAS en Galicia. Cálculo de rutas e tempos en vehículo en tempo real.

### Interface Web (Frontend)
- **`templates/login.html`**: Interface moderna de autenticación e rexistro.
- **`templates/dashboard.html`**: Panel xeral con listado de proxectos, busca, creación, duplicación e exportación/importación JSON.
- **`templates/project_editor.html`**: Editor unificado con 3 pestanas:
  1. *Datos Comúns do Proxecto* (sincronizados entre ambos documentos).
  2. *Estudo de Seguridade (EBSS)* con buscador a 1 clic de centros sanitarios e descarga independente do `.docx`.
  3. *Xestión de Residuos (EGR)* con matriz de segregación, cálculo en vivo e descarga independente do `.docx`.

---

## 3. Infraestrutura na Nube e Despregamento

### Base de Datos: Supabase (PostgreSQL)
- **Plataforma**: [Supabase.com](https://supabase.com)
- **Proxecto**: `MT2026` (`gen-lang-client-0779098551` / `fsbhartztdlydahrmuwa`)
- **Configuración**: Conéctase mediante a variable de contorna `DATABASE_URL`.
- **Táboas principais**: `users` (usuarios e hash de contrasinal) e `projects` (almacenamento relacional con JSON estructurado).

### Servizo Web: Google Cloud Run
- **Nome do servizo**: `mt2026`
- **Rexión**: `europe-southwest1` (Madrid, España)
- **Contedor**: Construído automaticamente con [Dockerfile](file:///g:/OneDrive/IA-MEMORIAS/PLIEGOS/Asistente_EBSS/Dockerfile) (`python:3.11-slim` + Gunicorn)
- **Repositorio conectado**: GitHub `https://github.com/pedrogarat/MT2026.git` (rama `main`) con integración continua (CI/CD).

### Dominio e Aloxamento Web: Firebase Hosting
- **URL Pública Oficial**: **[https://mtecnicas2026.web.app/](https://mtecnicas2026.web.app/)**
- **Configuración**: [firebase.json](file:///g:/OneDrive/IA-MEMORIAS/PLIEGOS/Asistente_EBSS/firebase.json) e [.firebaserc](file:///g:/OneDrive/IA-MEMORIAS/PLIEGOS/Asistente_EBSS/.firebaserc). Redirixe o 100% do tráfico (`**`) ao servizo Cloud Run `mt2026`.
- **REQUISITO CRÍTICO DE SEGURIDADE DE FIREBASE HOSTING**:
  Firebase Hosting elimina todas as cookies das peticións agás aquela que se chame exactamente `__session`.
  En `app.py` débese manter sempre:
  ```python
  app.config["SESSION_COOKIE_NAME"] = "__session"
  ```
  Ademais das cabeceiras `Cache-Control: no-cache, private` para evitar que a CDN cachee sesións de usuario.

---

## 4. Execución Local (Modo Escritorio / Sen Conexión)
- O script [iniciar_asistente.bat](file:///g:/OneDrive/IA-MEMORIAS/PLIEGOS/Asistente_EBSS/iniciar_asistente.bat) permite arrancar a ferramenta localmente en calquera equipo con Windows:
  - Crea a contorna virtual `.venv` e instala dependencias automaticamente se non existen.
  - Se non existe `DATABASE_URL`, utiliza como respaldo a base de datos SQLite local `asistente_ebss.db`.
  - Abre o navegador automaticamente en `http://127.0.0.1:8080`.
