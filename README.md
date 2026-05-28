# Scoring de Potencial Verde por Cualificación

Script de clasificación de cualificaciones del Marco Nacional de Cualificaciones (MNC) de Colombia según su potencial verde. Desarrollado para el estudio de transición justa en Cesar y La Guajira, Colombia.

## Qué hace

Para cada cualificación calcula dos componentes:

- **Score actual (0–10):** qué tan verde es la cualificación hoy, cruzando su competencia general con el Catastro de Perfiles y Competencias Verdes de la Alianza del Pacífico (SENA, 2024) y los cinco criterios OIT de empleos verdes. Clasifica en Verde (≥7), Potencialmente verde (≥4) o Convencional (<4).

- **Potencial de orientación (Alto/Medio/Bajo):** si la cualificación puede moverse hacia funciones verdes con formación complementaria de corta duración, usando como referencia O*NET Green Economy, ESCO y SOLAS Green Skills 2030.

El modelo de lenguaje responde preguntas binarias (sí/no) predefinidas. Los puntajes los calcula Python, lo que garantiza consistencia entre corridas.

## Requisitos

```
python >= 3.10
anthropic
pandas
openpyxl
python-dotenv
```

```bash
pip install -r requirements.txt
```

## Configuración

Copia `.env.template` como `.env` y completa los valores:

```
ANTHROPIC_API_KEY=sk-ant-...
EXCEL_PATH=/ruta/a/Cualificaciones.xlsx
```

El Excel debe tener una hoja por sector con columnas: `Codigo cualificacion`, `Nombre cualificacion`, `Competencia general`. La columna `Nivel MNC` es opcional.

## Uso

```bash
python scoring_verde_v2.py
```

El script procesa cada cualificación y genera `Cualificaciones_scoring_verde.xlsx` en la misma carpeta del Excel de entrada, con una fila por cualificación y todas las respuestas binarias individuales auditables.

## Modelo

`claude-sonnet-4-20250514` (Anthropic, 2025)
