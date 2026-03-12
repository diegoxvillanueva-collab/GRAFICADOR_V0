"""
S4 Graficador — Data models
Dataclasses que representan la estructura parseada del Excel de encuestas.
"""

from dataclasses import dataclass, field
from typing import List, Optional


# Paleta por defecto para asignar colores cuando faltan en meta_respuestas
DEFAULT_PALETTE = [
    "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
    "#86BCB6", "#8CD17D", "#B6992D", "#499894", "#E15759",
    "#F1CE63", "#D37295", "#A0CBE8", "#FFBE7D", "#8CD17D",
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
