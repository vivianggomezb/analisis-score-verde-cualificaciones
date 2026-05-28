"""
IKIJET — Scoring Verde de Cualificaciones MNC v3
Proyecto: Transición justa en Cesar y La Guajira / OIT

Arquitectura:
- Unidad de análisis: cualificación completa (competencia general), no UC
- Score actual (0–10): catastro Alianza del Pacífico 4 países — rúbrica binaria
- Potencial de orientación (Alto/Medio/Bajo): O*NET + ESCO + SOLAS Green Skills 2030
- Modelo: claude-sonnet-4-20250514
- Python calcula los puntajes; el modelo solo responde sí/no a preguntas concretas

Uso:
    python scoring_verde.py

Requiere:
    - .env con ANTHROPIC_API_KEY y EXCEL_PATH
    - pip install anthropic pandas openpyxl python-dotenv
"""

import os, json, re, time
import pandas as pd
from dotenv import load_dotenv
import anthropic

# ──────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────

load_dotenv()
API_KEY    = os.getenv("ANTHROPIC_API_KEY")
EXCEL_PATH = os.getenv("EXCEL_PATH")
MODEL      = "claude-sonnet-4-20250514"

if not API_KEY:   raise ValueError("ANTHROPIC_API_KEY no encontrada en .env")
if not EXCEL_PATH: raise ValueError("EXCEL_PATH no encontrada en .env")

client = anthropic.Anthropic(api_key=API_KEY)

# ──────────────────────────────────────────────────────
# REFERENCIA FIJA — CATASTRO ALIANZA DEL PACÍFICO (2024)
# Chile (ChileValora), Colombia (SENA), México (CONOCER), Perú (MinTrabajo)
# ──────────────────────────────────────────────────────

CATASTRO_AP = """
CATASTRO DE PERFILES Y COMPETENCIAS VERDES — ALIANZA DEL PACÍFICO 2024
Fuente: Arredondo Pérez, G. (2024). Eurosocial Puente / Alianza del Pacífico.
Países: Chile (CHL), Colombia (CO), México (MX), Perú (PE).

ENERGÍAS RENOVABLES Y LIMPIAS
- Hidrógeno verde: operar/mantener planta H2, electrolizador, almacenamiento, agua (CHL)
- Energía termosolar: instalar helióstatos, operar/mantener campo solar y sala de control (CHL)
- Energía fotovoltaica: instalar/mantener sistemas FV residencial e industrial
  (CHL: P-3510-7411-001; CO: 291901055/291901057; MX: EC0586.01/EC1181)
- Diseñar redes de energía sostenible solar y eólica (CO: 291901057)
- Energía eólica: mantener aerogeneradores; montar sistemas eólicos (MX: EC0530; CO: 291901055)
- Biogás: operar plantas de biogás; gestión documental de proyectos (MX: EC1077/EC1128)

ELECTROMOVILIDAD
- Instalar/mantener infraestructura de recarga para vehículos eléctricos (CHL)
- Diagnosticar/reparar propulsión eléctrica e híbrida vehicular
  (CO: 280601082-085/122; CHL; MX: EC1528/EC1547)

CONSTRUCCIÓN SOSTENIBLE
- Instalar/mantener sistemas solares térmicos de circulación forzada (CHL; MX: EC0473/EC0325)
- Estructurar sistemas de energías renovables según normativa ambiental (CO: 220201093)
- Supervisar instalaciones fotovoltaicas (MX: EC1181)
- Instalar sistemas de iluminación eficiente (MX: EC0414)

RESIDUOS Y RECICLAJE
- Recolectar, clasificar, comercializar materiales reciclables (CHL: P-3830-9611-001/002)
- Recolectar/reciclar residuos sólidos y orgánicos (CO: 220201080/092; 280201236/245)
- Gestionar residuos sólidos y peligrosos (CO: 220201078/111/113)
- Desensamblar residuos de aparatos eléctricos y electrónicos (CO: 291901056)
- Manejo de residuos peligrosos (MX: EC0674)
- Recuperar, segregar y acondicionar residuos sólidos aprovechables (PE: D1938001-1/2/3)

GESTIÓN ENERGÉTICA Y AMBIENTAL EN ORGANIZACIONES
- Caracterizar uso energético; elaborar e implementar plan de mejora energética (CHL)
- Gestión ambiental estratégica y operación de sistema de gestión ambiental (MX: EC0490/EC0517/EC1543)
- Evaluar impactos/requisitos ambientales; logística inversa; políticas ambientales
  (CO: 220201088/115/090/099/102/100)

EFICIENCIA ENERGÉTICA
- Gestión de eficiencia energética organizacional (MX: EC0412)
- Mantenimiento sistema energético de inmuebles (MX: EC0413/EC0416)
- Control eficiencia energética en estaciones de bombeo (MX: EC0317)
- Diagnóstico eficiencia energética y seguridad eléctrica en vivienda (MX: EC1125)

EDUCACIÓN Y CAMPAÑAS AMBIENTALES
- Implementar campañas ambientales (CO: 220201116)
- Promoción de sensibilización ambiental; formación para desarrollo rural sustentable (MX)

SANEAMIENTO, AGUA Y SALUD AMBIENTAL
- Inspeccionar sistemas de agua / monitorear emisiones fuentes fijas (CO: 280201234/220201107)
- Tratamiento de aguas residuales (MX: EC0210/EC0214/EC0216; PE: D1937002-1/2)

BIODIVERSIDAD Y PROTECCIÓN AMBIENTAL
- Valorar ecosistemas según normativa ambiental y biodiversidad (CO: 220201095)
- Monitoreo comunitario de biodiversidad en áreas naturales protegidas (MX: EC1401)

AGRICULTURA, FORESTAL Y ZONAS VERDES
- Diseñar zonas verdes, techos verdes, jardines verticales (CO: 220201094)
- Inventariar especies forestales (CO: 270301043)
- Promover conservación de recursos forestales (PE: A0102006-1)
- Certificación producción orgánica (PE: M2974004-1/2/3)

TURISMO Y NEGOCIOS VERDES
- Categorizar negocios sostenibles según criterios verdes (CO: 220201082)
- Prestación de servicio hotelero con orientación a la sostenibilidad (MX: EC1020)
- Aplicación de prácticas verdes en área de trabajo (MX: EC0612)
- Monitoreo actividad pesquera ribereña / buenas prácticas pesqueras (MX: EC0578/EC0820)
"""

# ──────────────────────────────────────────────────────
# REFERENCIA FIJA — MARCOS GLOBALES PARA POTENCIAL
# ──────────────────────────────────────────────────────

MARCOS_POTENCIAL = """
MARCOS DE REFERENCIA PARA POTENCIAL DE ORIENTACIÓN VERDE:

1. O*NET Green Economy (DOL/ETA, EE.UU.)
Clasifica ocupaciones en tres categorías:
- Green New & Emerging: nuevas ocupaciones que surgen directamente de la transición verde
  (ej. técnico en energía solar, auditor de carbono)
- Green Enhanced Skills: ocupaciones existentes cuyas tareas y competencias cambian
  significativamente con la transición (ej. electricista con nuevas competencias en FV,
  ingeniero ambiental, operador de planta con eficiencia energética)
- Green Increased Demand: ocupaciones existentes con mayor demanda por la transición,
  sin cambio sustancial en sus tareas (ej. conductores, operadores de maquinaria)
Sectores cubiertos: energía renovable, eficiencia energética, construcción verde,
transporte limpio, conservación de recursos, gestión ambiental, agricultura sostenible,
manufactura verde, gestión de residuos.

2. ESCO (European Skills/Competences, Qualifications and Occupations — UE)
Taxonomía de competencias verdes transferibles entre ocupaciones:
- Gestión energética y eficiencia
- Economía circular y gestión de residuos
- Protección ambiental y biodiversidad
- Adaptación y mitigación climática
- Producción sostenible y agricultura ecológica
- Movilidad limpia y electromovilidad
- Construcción sostenible y renovación energética
ESCO identifica cuáles competencias de ocupaciones existentes son transferibles a
empleos verdes con formación complementaria.

3. SOLAS Green Skills 2030 (Irlanda, octubre 2024)
Estrategia nacional de formación profesional para la transición verde.
Identifica brechas de competencias verdes por sector y ocupación, y qué upskilling
es viable en el corto plazo (cursos cortos, módulos flexibles):
- Construcción y entorno construido: métodos modernos de construcción, retrofitting,
  edificación de energía casi nula, gestión ambiental, soluciones basadas en naturaleza
- Ingeniería, energía y manufactura: instalación/mantenimiento energías renovables,
  análisis de datos energéticos, sistemas energéticos nuevos, gestión de residuos
- Transporte y logística: vehículos eléctricos, electrificación de infraestructura,
  combustibles alternativos, gestión de cadena de suministro sostenible
- Agricultura, forestal y marino: agricultura sostenible, restauración de ecosistemas,
  bioeconomía, habilidades digitales aplicadas al agro
- Biodiversidad y medioambiente: ecología, ciencias ambientales, habilidades prácticas
  de gestión territorial, conocimiento de sostenibilidad
- Turismo y hostelería: implementación de sostenibilidad, integración de criterios
  verdes en operación hotelera, artesanía sostenible
- Contabilidad y negocio: reporte ESG, contabilidad de carbono, cumplimiento normativo
  ambiental, comprensión de legislación climática
Principio clave de SOLAS: la mayoría de ocupaciones existentes pueden incorporar
competencias verdes mediante formación corta y flexible, sin cambio de perfil.
"""

# ──────────────────────────────────────────────────────
# PROMPTS DE SISTEMA
# ──────────────────────────────────────────────────────

SYSTEM_SCORE = f"""Eres un analista experto en competencias laborales verdes del proyecto IKIJET
(OIT / Colombia). Evalúas cualificaciones del Marco Nacional de Cualificaciones (MNC)
colombiano para el estudio de transición justa en Cesar y La Guajira.

Tu tarea es evaluar la competencia general de una cualificación contra el Catastro de
Perfiles y Competencias Verdes de la Alianza del Pacífico (2024) y los criterios OIT.

CATASTRO DE REFERENCIA (Chile, Colombia, México, Perú):
{CATASTRO_AP}

CRITERIOS OIT DE EMPLEOS VERDES:
1. Eficiencia energética y uso eficiente de materias primas
2. Reducción de emisiones de GEI
3. Minimización de residuos y contaminación
4. Protección y restauración de ecosistemas
5. Adaptación al cambio climático

REGLAS:
- Evalúa SOLO lo que la competencia general describe explícitamente.
- "Cumplir normativa ambiental" como requisito de contexto NO cuenta como competencia verde central.
- El propósito ambiental debe ser la función central, no una externalidad.
- Ante la duda, aplica el criterio más bajo (conservadurismo).

Responde ÚNICAMENTE con JSON válido, sin backticks ni texto adicional:
{{
  "codigo_cualificacion": "",
  "p1_catastro_directa": true/false,
  "p2_catastro_parcial": true/false,
  "p3_eficiencia_energetica": true/false,
  "p4_reduccion_gei": true/false,
  "p5_residuos_contaminacion": true/false,
  "p6_ecosistemas": true/false,
  "p7_adaptacion_clima": true/false,
  "justificacion": "Máximo 3 oraciones."
}}

p1: ¿La competencia general o una función idéntica aparece en el catastro AP?
p2: ¿El catastro AP tiene una competencia en el MISMO sector productivo con funciones sustantivamente similares? (solo si p1=false). CRITERIO ESTRICTO: no basta con que ambas involucren trabajo en campo, con recursos naturales, o con maquinaria en general. La similitud debe ser de funciones específicas dentro del mismo sector. Ejemplos de lo que NO cuenta como p2=true: monitoreo de biodiversidad en áreas protegidas NO es análogo a producción agrícola operativa; instalación de sistemas eléctricos convencionales NO es análogo a instalación fotovoltaica; operación de maquinaria minera NO es análogo a operación de maquinaria en construcción sostenible.
p3-p7: ¿La función CENTRAL de esta competencia contribuye explícitamente a este criterio OIT? No cuenta si el criterio aparece solo como cumplimiento normativo de contexto ("respetando normas ambientales") o como externalidad del proceso. Debe ser el propósito principal de la competencia.
"""

SYSTEM_POTENCIAL = f"""Eres un analista experto en transición justa y empleos verdes del proyecto IKIJET (OIT).
Evalúas el potencial de orientación verde de cualificaciones del MNC colombiano.

MARCOS DE REFERENCIA:
{MARCOS_POTENCIAL}

El potencial de orientación mide si las funciones de la cualificación pueden redirigirse
o complementarse con prácticas sostenibles SIN requerir un cambio estructural del perfil.

Responde ÚNICAMENTE con JSON válido, sin backticks ni texto adicional:
{{
  "codigo_cualificacion": "",
  "q1_analogia_marcos": true/false,
  "q2_analogia_directa": true/false,
  "q3_upskilling_corto_plazo": true/false,
  "justificacion_potencial": "Máximo 2 oraciones."
}}

q1: ¿Existe en O*NET, ESCO o SOLAS una ocupación/competencia verde con funciones análogas?
q2: ¿La analogía es directa (mismas funciones en contexto sostenible)? (solo si q1=true)
q3: ¿Puede el trabajador desarrollar competencias verdes con formación corta (≤6 meses)
    sin cambiar de perfil ocupacional, según los marcos de referencia?
"""

# ──────────────────────────────────────────────────────
# CÁLCULO DE SCORES (Python, no el modelo)
# ──────────────────────────────────────────────────────

def calcular_score_actual(r: dict) -> float:
    """
    Presencia en catastro: directa=5, parcial=3, ninguna=0
    Criterios OIT: 1 punto cada uno (máx 5)
    Total máximo: 10
    """
    if r.get("p1_catastro_directa"):
        pts = 5.0
    elif r.get("p2_catastro_parcial"):
        pts = 3.0
    else:
        pts = 0.0
    pts += sum([
        r.get("p3_eficiencia_energetica", False),
        r.get("p4_reduccion_gei", False),
        r.get("p5_residuos_contaminacion", False),
        r.get("p6_ecosistemas", False),
        r.get("p7_adaptacion_clima", False),
    ])
    return round(pts, 2)


def nivel_score(score: float) -> str:
    if score >= 7.0: return "Verde"
    if score >= 4.0: return "Potencialmente verde"
    return "Convencional"


def calcular_potencial(r: dict) -> str:
    q1 = r.get("q1_analogia_marcos", False)
    q2 = r.get("q2_analogia_directa", False)
    q3 = r.get("q3_upskilling_corto_plazo", False)
    if q1 and q2 and q3: return "Alto"
    if q1 and q3:         return "Medio"
    if q1:                return "Medio"
    return "Bajo"


# ──────────────────────────────────────────────────────
# UTILIDADES
# ──────────────────────────────────────────────────────

def parse_json(texto: str) -> dict | None:
    texto = texto.strip()
    if texto.startswith("```"):
        partes = texto.split("```")
        for p in partes:
            p = p.strip().lstrip("json").strip()
            if p.startswith("{"): texto = p; break
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        return None


def llamar_api(system: str, contenido: str, reintentos=3) -> dict | None:
    for i in range(reintentos):
        try:
            r = client.messages.create(
                model=MODEL, max_tokens=500, system=system,
                messages=[{"role": "user", "content": contenido}],
            )
            resultado = parse_json(r.content[0].text)
            if resultado: return resultado
            print(f"    ⚠️  JSON inválido (intento {i+1})")
        except Exception as e:
            print(f"    ⚠️  Error API (intento {i+1}): {e}")
        time.sleep(2 ** i)
    return None


# ──────────────────────────────────────────────────────
# PROCESO PRINCIPAL
# ──────────────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"IKIJET — Scoring Verde v3 | {MODEL}")
    print(f"{'='*60}\n")

    print(f"📂 {EXCEL_PATH}")
    hojas = pd.read_excel(EXCEL_PATH, sheet_name=None)
    print(f"   Hojas: {list(hojas.keys())} | Total cualificaciones: {sum(len(d) for d in hojas.values())}\n")

    resultados = []

    for sector, df in hojas.items():
        print(f"\n{'─'*50}")
        print(f"📋 Sector: {sector} ({len(df)} cualificaciones)")
        print(f"{'─'*50}")

        for _, fila in df.iterrows():
            codigo = str(fila["Codigo cualificacion"]).strip()
            nombre = str(fila["Nombre cualificacion"]).strip()
            comp_general = str(fila["Competencia general"]).strip()
            nivel_mnc = str(fila.get("Nivel MNC", "")).strip()

            print(f"\n  🔍 {codigo} — {nombre[:55]}...")

            input_cual = json.dumps({
                "sector": sector,
                "codigo_cualificacion": codigo,
                "nombre_cualificacion": nombre,
                "nivel_mnc": nivel_mnc,
                "competencia_general": comp_general,
            }, ensure_ascii=False)

            # Llamada 1: Score actual
            resp_score = llamar_api(SYSTEM_SCORE, input_cual)
            if resp_score:
                score_num  = calcular_score_actual(resp_score)
                nivel_s    = nivel_score(score_num)
                just_score = resp_score.get("justificacion", "")
            else:
                score_num, nivel_s, just_score = None, "Error", "Error API"
                resp_score = {}

            time.sleep(0.5)

            # Llamada 2: Potencial de orientación
            resp_pot = llamar_api(SYSTEM_POTENCIAL, input_cual)
            if resp_pot:
                potencial = calcular_potencial(resp_pot)
                just_pot  = resp_pot.get("justificacion_potencial", "")
            else:
                potencial, just_pot = "Error", "Error API"
                resp_pot = {}

            print(f"    Score: {score_num} → {nivel_s} | Potencial: {potencial}")

            resultados.append({
                "Sector":                    sector,
                "Código cualificación":      codigo,
                "Nombre cualificación":      nombre,
                "Nivel MNC":                 nivel_mnc,
                "Competencia general":       comp_general,
                # Score actual — respuestas binarias
                "p1_catastro_directa":       resp_score.get("p1_catastro_directa"),
                "p2_catastro_parcial":       resp_score.get("p2_catastro_parcial"),
                "p3_eficiencia_energetica":  resp_score.get("p3_eficiencia_energetica"),
                "p4_reduccion_gei":          resp_score.get("p4_reduccion_gei"),
                "p5_residuos_contaminacion": resp_score.get("p5_residuos_contaminacion"),
                "p6_ecosistemas":            resp_score.get("p6_ecosistemas"),
                "p7_adaptacion_clima":       resp_score.get("p7_adaptacion_clima"),
                "Score actual (0-10)":       score_num,
                "Nivel score actual":        nivel_s,
                "Justificación score":       just_score,
                # Potencial — respuestas binarias
                "q1_analogia_marcos":        resp_pot.get("q1_analogia_marcos"),
                "q2_analogia_directa":       resp_pot.get("q2_analogia_directa"),
                "q3_upskilling_corto_plazo": resp_pot.get("q3_upskilling_corto_plazo"),
                "Potencial de orientación":  potencial,
                "Justificación potencial":   just_pot,
            })

            time.sleep(0.5)

    # Exportar
    print(f"\n\n{'='*60}")
    print("💾 Exportando resultados...")

    df_out = pd.DataFrame(resultados)
    output_path = EXCEL_PATH.replace(".xlsx", "_scoring_verde.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Resultados", index=False)

    print(f"✅ Guardado en: {output_path}")
    print(f"   Cualificaciones procesadas: {len(df_out)}")
    print(f"\n{'='*60}\n")


if __name__ == "__main__":
    main()
