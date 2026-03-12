"""
S4 Graficador — Data models
Dataclasses que representan la estructura parseada del Excel de encuestas.
"""

from dataclasses import dataclass, field
from typing import List, Optional


# ──────────────────────────────────────────────────────────────────────────────
# COLORES PREDEFINIDOS ÁGORA
# ──────────────────────────────────────────────────────────────────────────────

# Mapa de colores por label de respuesta conocida (case-insensitive lookup)
# Cuando un label coincide, se usa este color en vez del de meta_respuestas o auto-asignado
PREDEFINED_COLORS = {
    # Escala imagen / gestión
    "Muy buena": "#04967A",
    "Buena": "#03BD85",
    "Mala": "#EA3F28",
    "Muy mala": "#C1273A",
    "No sabe": "#BFBFBF",
    # Cercanía partidaria
    "Un candidato del peronismo cercano al kirchnerismo": "#0070C0",
    "Un candidato de La Libertad Avanza de Javier Milei": "#7030A0",
    "Un candidato del peronismo No kirchnerista": "#007F48",
    "Un candidato de la izquierda": "#FF008E",
    "Un candidato del PRO de Mauricio Macri": "#FFC000",
    "Un candidato del radicalismo": "#C1273A",
    "Otro": "#A6A6A6",
    # Escala intensidad
    "Mucho": "#023D9B",
    "Bastante": "#1D73FC",
    "Poco": "#FF008E",
    "Nada": "#C1273A",
    # Escala probabilidad
    "Muy probable": "#023D9B",
    "Bastante probable": "#1D73FC",
    "Poco probable": "#FF008E",
    "Nunca lo votaría": "#C1273A",
    # Escala continuidad
    "Continuar como hasta ahora": "#023D9B",
    "Continuar con algunos cambios": "#1D73FC",
    "Cambiar manteniendo solo algunas cosas": "#FF008E",
    "Cambiar totalmente": "#C1273A",
    # Sí / No
    "Si": "#023D9B",
    "Sí": "#023D9B",
    "No": "#E85833",
    # Clase social
    "Clase baja": "#000000",
    "Clase media baja": "#023D9B",
    "Clase media": "#1D73FC",
    "Clase media alta": "#E85833",
    "Clase alta": "#C1273A",
}

# Paleta para variables categóricas nominales (cuando no hay match en PREDEFINED_COLORS)
DEFAULT_PALETTE = [
    "#000000", "#023D9B", "#1D73FC", "#FF008E", "#7030A0",
    "#AF1956", "#E85833", "#007F48", "#02BFA3",
    # Extensión por si hay más de 9 categorías
    "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
]

# ──────────────────────────────────────────────────────────────────────────────
# FORMATO VISUAL ÁGORA — Especificación completa de fuentes por tipo de chart
# Tamaños en centipuntos (pt × 100). Ej: 24pt = 2400
# ──────────────────────────────────────────────────────────────────────────────

AGORA_STYLE = {
    # ── GENERAL (aplica a todas las diapositivas) ──
    "slide_title_font": "HelveticaNeue Std (Cuerpo)",
    "slide_title_size": 3600,       # 36pt
    "slide_title_bold": True,       # solo el nombre del apartado (ej "02. Gestiones")
    "footer_font": "Helvetica Neue",
    "footer_size": 1800,            # 18pt
    "footer_bold": True,

    # ── FREC_SIMPLE ──
    "frec_simple": {
        "cat_axis_font": "HelveticaNeue Std (Cuerpo)",
        "cat_axis_size": 2200,      # 22pt — eje horizontal
        "cat_axis_bold": False,
        "data_label_font": "HelveticaNeue Std (Cuerpo)",
        "data_label_size": 2400,    # 24pt — % dentro del gráfico
        "data_label_bold": True,
        "data_label_color": "bg1",  # blanco sobre barra
    },

    # ── FREC_MULTIPLE ──
    "frec_multiple": {
        "chart_title_font": "HelveticaNeue Std",
        "chart_title_size": 2400,   # 24pt
        "chart_title_bold": True,
        "legend_font": "HelveticaNeue Std",
        "legend_size": 2400,        # 24pt
        "legend_bold": False,
        "val_axis_font": "HelveticaNeue Std",
        "val_axis_size": 2200,      # 22pt — eje vertical
        "val_axis_bold": True,
        "data_label_font": "HelveticaNeue Std (Cuerpo)",
        "data_label_size": 2000,    # 20pt — % dentro del gráfico
        "data_label_bold": True,
        "data_label_color": "bg1",
    },

    # ── FREC_MULTI_GRAFICOS (usa mismos specs que FREC_SIMPLE) ──
    "frec_multi_graficos": {
        "chart_title_font": "HelveticaNeue Std",
        "chart_title_size": 2400,   # 24pt
        "chart_title_bold": True,
        "cat_axis_font": "HelveticaNeue Std (Cuerpo)",
        "cat_axis_size": 2200,      # 22pt
        "cat_axis_bold": False,
        "data_label_font": "HelveticaNeue Std (Cuerpo)",
        "data_label_size": 2400,    # 24pt
        "data_label_bold": True,
        "data_label_color": "bg1",
    },

    # ── APERTURA_SIMPLE ──
    "apertura_simple": {
        "legend_font": "HelveticaNeue Std",
        "legend_size": 2400,        # 24pt
        "legend_bold": False,
        "bracket_font": "HelveticaNeue Std",
        "bracket_size": 1800,       # 18pt — corchetes (burbujas futuras)
        "bracket_bold": True,
        "data_label_font": "Roboto Condensed",
        "data_label_size": 1400,    # 14pt — % dentro del gráfico
        "data_label_bold": True,
        "data_label_color": "bg1",
        "cat_axis_font": "HelveticaNeue Std (Títulos)",
        "cat_axis_size": 1800,      # 18pt — eje horizontal
        "cat_axis_bold": False,
    },

    # ── EVOLUTIVO1 (fase futura) ──
    "evolutivo1": {
        "legend_font": "HelveticaNeue Std (Cuerpo)",
        "legend_size": 2400,        # 24pt
        "legend_bold": False,
        "data_label_font": "HelveticaNeue Std",
        "data_label_size": 2000,    # 20pt
        "data_label_bold": True,
        "cat_axis_font": "HelveticaNeue Std",
        "cat_axis_size": 1800,      # 18pt
        "cat_axis_bold": False,
    },

    # ── EVOLUTIVO2 (fase futura) ──
    "evolutivo2": {
        "legend_font": "HelveticaNeueLT Std (Cuerpo)",
        "legend_size": 2000,        # 20pt
        "legend_bold": False,
        "chart_title_font": "HelveticaNeue Std",
        "chart_title_size": 2400,   # 24pt
        "chart_title_bold": True,
        "data_label_font": "HelveticaNeue Std (Cuerpo)",
        "data_label_size": 2000,    # 20pt
        "data_label_bold": True,
        "cat_axis_font": "HelveticaNeue Std (Cuerpo)",
        "cat_axis_size": 2000,      # 20pt
        "cat_axis_bold": False,
    },

    # ── EVOLUTIVO3 (fase futura) ──
    "evolutivo3": {
        "data_label_font": "Roboto",
        "data_label_size": 2400,    # 24pt
        "data_label_bold": True,
        "legend_font": "HelveticaNeue Std (Títulos)",
        "legend_size": 2000,        # 20pt
        "legend_bold": False,
    },

    # ── LAYOUT ──
    "frec_gap_width": 30,
    "apertura_gap_width": 50,
    "number_format": "0%",
}


def get_chart_style(chart_type: str) -> dict:
    """Retorna el sub-dict de estilo para un tipo de chart específico."""
    key_map = {
        "FREC_SIMPLE": "frec_simple",
        "FREC_MULTIPLE": "frec_multiple",
        "FREC_MULTI_GRAFICOS": "frec_multi_graficos",
        "APERTURA_SIMPLE": "apertura_simple",
        "EVOLUTIVO1": "evolutivo1",
        "EVOLUTIVO2": "evolutivo2",
        "EVOLUTIVO3": "evolutivo3",
    }
    return AGORA_STYLE.get(key_map.get(chart_type, ""), {})


@dataclass
class Answer:
    """Una respuesta dentro de una pregunta."""
    label: str              # Label de la respuesta (ej: "Muy buena")
    color: str              # Hex color (ej: "#04967A")
    color_pending: bool     # True si el color fue asignado automáticamente
    total: float            # Valor col D (proporción 0-1, para frecuencia)
    segment_values: List[float]  # Valores cols F+ (para apertura), en orden


@dataclass
class Question:
    """Una pregunta parseada con toda su metadata y datos."""
    id_pregunta: str
    tipo_slide: str             # FREC_SIMPLE, FREC_MULTIPLE, FREC_MULTI_GRAFICOS
    orientacion: str            # "V" (vertical) o "H" (horizontal) — para frecuencias
    capitulo: str               # Capítulo de la frecuencia
    cap_app: Optional[str]      # Capítulo de la apertura (None = va después de frec)
    titulo_frec: str            # Título de la slide de frecuencia (parte izq del |)
    titulo_app: str             # Título de la slide de apertura / pregunta (parte der del |)
    grupo_frec: int             # Agrupa preguntas en una misma slide de frecuencia
    multiple: bool              # True si respuesta múltiple
    pregunta_texto: str         # Texto completo de la pregunta
    respuestas: List[Answer]
    segment_labels: List[str]   # Labels de segmentos en orden
    # Estructura para burbujas de suma (fase futura)
    suma_positiva_indices: Optional[List[int]] = None   # índices de respuestas "positivas" (ej: Muy buena, Buena)
    suma_negativa_indices: Optional[List[int]] = None   # índices de respuestas "negativas" (ej: Mala, Muy Mala)


@dataclass
class SegmentGroup:
    """Un grupo de segmentos (ej: Sexo con columnas Varón, Mujer)."""
    name: str
    labels: List[str]
    col_indices: List[int] = field(default_factory=list)  # índices de columna en el Excel


@dataclass
class ParseResult:
    """Resultado completo del parsing del Excel."""
    questions: List[Question]
    segment_labels: List[str]
    segment_groups: List[SegmentGroup]
    warnings: List[str]
    color_auto_assigned: int
