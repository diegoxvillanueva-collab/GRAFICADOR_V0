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

# Formato visual Ágora — extraído del template de referencia
AGORA_STYLE = {
    "chart_title_font": "HelveticaNeueLT Std (Cuerpo)",
    "chart_title_size": 2400,  # centipuntos (24pt)
    "chart_title_bold": True,
    "data_label_font": "Roboto Condensed",
    "data_label_size": 1400,  # 14pt
    "data_label_bold": True,
    "data_label_color_scheme": "bg1",  # blanco sobre barra
    "axis_label_font": "HelveticaNeue Std (Cuerpo)",
    "axis_label_size": 1400,  # 14pt
    "axis_label_bold": True,
    "slide_title_font": "HelveticaNeueLT Std (Cuerpo)",
    "slide_title_size": 3600,  # 36pt
    "slide_title_bold": False,
    "frec_gap_width": 30,
    "apertura_gap_width": 50,
    "number_format": "0%",
}


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
