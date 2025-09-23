# config.py

# =========================
# General chart settings
# =========================
DPI = 120  # consistent sharpness for all charts

# Whether to show average line across charts
SHOW_AVERAGE_LINE = True
AVERAGE_LINE_STYLE = {
    "color": "blue",
    "linestyle": "--",
    "linewidth": 2,
    "label": "Durchschnitt",
}
AVERAGE_LABEL_FMT = "Durchschnitt: {0:.0f}"

# =========================
# Legend settings
# =========================
LEGEND_LOC = "upper center"
LEGEND_NCOL = 3  # number of columns in legends

# =========================
# Shifts & Colors
# =========================
SHIFT_ORDER = ["Früh", "Spät", "Nacht"]

SHIFT_COLORS = {
    "Früh": "#1f77b4",   # blue
    "Spät": "#ff7f0e",   # orange
    "Nacht": "#2ca02c",  # green
}

# =========================
# Targets (e.g., Sägen per day)
# =========================
SAEGEN_TARGET = 70  # daily target for Sägen (used in weekly plots)

# =========================
# Annotation policy
# =========================
ANNOTATION_FONT_SIZE = 10
ANNOTATION_FONT_WEIGHT = "bold"

TOTAL_LABEL_FONT_SIZE = 12
TOTAL_LABEL_FONT_WEIGHT = "bold"
