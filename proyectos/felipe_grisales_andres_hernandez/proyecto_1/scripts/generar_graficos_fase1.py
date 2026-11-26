#!/usr/bin/env python3
"""
Generador reproducible de figuras técnicas para la Fase 1:
Conceptualización en Diagrama Causal (CLD), Modos de Referencia y Arquitectura Micro-Macro.

Produce versiones vectoriales (.svg) y rasterizadas de alta resolución (.png)
usando únicamente la biblioteca estándar de Python y Pillow (PIL).
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Sequence, Tuple
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont

# Rutas de salida
BASE_DIR = Path(__file__).resolve().parent
FIG_DIR = BASE_DIR / "figuras"

# Fuentes del sistema (DejaVu)
FONT_REG_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_MONO_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

# Paleta técnica unificada (Swiss Technical Manual)
PALETTE = {
    "bg": "#F8FAFC",
    "surface": "#FFFFFF",
    "surface_alt": "#F1F5F9",
    "ink": "#0F172A",
    "ink_muted": "#475569",
    "ink_light": "#64748B",
    "border": "#CBD5E1",
    "border_strong": "#94A3B8",
    # Subsistemas
    "div_fill": "#EFF6FF",
    "div_stroke": "#1D4ED8",
    "div_text": "#1E3A8A",
    "qual_fill": "#ECFDF5",
    "qual_stroke": "#047857",
    "qual_text": "#064E3B",
    "param_fill": "#FFFBEB",
    "param_stroke": "#B45309",
    "param_text": "#78350F",
    "aux_fill": "#F0FDFA",
    "aux_stroke": "#0F766E",
    "aux_text": "#134E4A",
    # Polaridades y bucles
    "pos": "#0F766E",
    "neg": "#BE123C",
    "loop_stroke": "#334155",
    # Series Regímenes
    "reg1": "#DC2626",   # Pmut=0.00, E=1 (Agotamiento y estancamiento)
    "reg2a": "#059669",  # Pmut=0.05, E=1 (Equilibrio óptimo exploración/carga)
    "reg2b": "#1D4ED8",  # Pmut=0.15, E=1 (Caso Base Mochila.xlsx)
    "reg3": "#D97706",   # Pmut=0.15, E=0 (Degradación sin elitismo)
    "obs": "#0F172A",    # Datos empíricos Mochila.xlsx
    "grid": "#E2E8F0",
}


def load_font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


class DualCanvas:
    """Lienzo dual que dibuja simultáneamente en SVG y en imagen PIL supersampled 2x."""

    def __init__(self, width: int, height: int, bg: str = PALETTE["bg"], scale: int = 2) -> None:
        self.w = width
        self.h = height
        self.s = scale
        self.bg = bg
        self.img = Image.new("RGB", (width * scale, height * scale), bg)
        self.draw = ImageDraw.Draw(self.img, "RGBA")
        self.svg_elems: list[str] = [
            f'<rect x="0" y="0" width="{width}" height="{height}" fill="{bg}"/>'
        ]
        self._font_cache: dict[tuple[str, int], ImageFont.FreeTypeFont] = {}

    def font(self, bold: bool = False, mono: bool = False, size: int = 16) -> ImageFont.FreeTypeFont:
        path = FONT_MONO_PATH if mono else (FONT_BOLD_PATH if bold else FONT_REG_PATH)
        key = (path, size * self.s)
        if key not in self._font_cache:
            self._font_cache[key] = load_font(path, size * self.s)
        return self._font_cache[key]

    def _sc(self, val: float) -> float:
        return val * self.s

    def rect(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        fill: str,
        stroke: str | None = None,
        stroke_w: float = 1.5,
        rx: float = 4.0,
        dash: str | None = None,
    ) -> None:
        x0, y0, x1, y1 = self._sc(x), self._sc(y), self._sc(x + w), self._sc(y + h)
        r = self._sc(rx)
        if rx > 0:
            self.draw.rounded_rectangle(
                [x0, y0, x1, y1],
                radius=r,
                fill=fill,
                outline=stroke if stroke and not dash else None,
                width=max(1, round(self._sc(stroke_w))) if stroke else 0,
            )
        else:
            self.draw.rectangle(
                [x0, y0, x1, y1],
                fill=fill,
                outline=stroke if stroke and not dash else None,
                width=max(1, round(self._sc(stroke_w))) if stroke else 0,
            )

        if stroke and dash:
            pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h), (x, y)]
            self.polyline(pts, stroke=stroke, width=stroke_w, dash=True)

        attrs = [
            f'x="{x:.1f}"',
            f'y="{y:.1f}"',
            f'width="{w:.1f}"',
            f'height="{h:.1f}"',
            f'fill="{fill}"',
        ]
        if rx > 0:
            attrs.append(f'rx="{rx:.1f}"')
        if stroke:
            attrs.append(f'stroke="{stroke}" stroke-width="{stroke_w:.1f}"')
        if dash:
            attrs.append(f'stroke-dasharray="{dash}"')
        self.svg_elems.append(f"<rect {' '.join(attrs)}/>")

    def circle(
        self,
        cx: float,
        cy: float,
        r: float,
        fill: str,
        stroke: str | None = None,
        stroke_w: float = 1.5,
    ) -> None:
        bbox = [self._sc(cx - r), self._sc(cy - r), self._sc(cx + r), self._sc(cy + r)]
        self.draw.ellipse(
            bbox,
            fill=fill,
            outline=stroke,
            width=max(1, round(self._sc(stroke_w))) if stroke else 0,
        )
        s_attr = f' stroke="{stroke}" stroke-width="{stroke_w:.1f}"' if stroke else ""
        self.svg_elems.append(
            f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="{fill}"{s_attr}/>'
        )

    def line(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        stroke: str,
        width: float = 1.5,
        dash: bool = False,
    ) -> None:
        if not dash:
            self.draw.line(
                [self._sc(x1), self._sc(y1), self._sc(x2), self._sc(y2)],
                fill=stroke,
                width=max(1, round(self._sc(width))),
            )
            self.svg_elems.append(
                f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
                f'stroke="{stroke}" stroke-width="{width:.1f}"/>'
            )
        else:
            dx, dy = x2 - x1, y2 - y1
            dist = math.hypot(dx, dy)
            if dist < 1e-6:
                return
            dash_len, gap_len = 8.0, 6.0
            step = dash_len + gap_len
            n = max(1, int(dist // step))
            for i in range(n + 1):
                t0 = (i * step) / dist
                t1 = min(1.0, (i * step + dash_len) / dist)
                if t0 >= 1.0:
                    break
                sx0, sy0 = x1 + dx * t0, y1 + dy * t0
                sx1, sy1 = x1 + dx * t1, y1 + dy * t1
                self.draw.line(
                    [self._sc(sx0), self._sc(sy0), self._sc(sx1), self._sc(sy1)],
                    fill=stroke,
                    width=max(1, round(self._sc(width))),
                )
            self.svg_elems.append(
                f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
                f'stroke="{stroke}" stroke-width="{width:.1f}" stroke-dasharray="8 6"/>'
            )

    def polyline(
        self,
        pts: Sequence[Tuple[float, float]],
        stroke: str,
        width: float = 2.0,
        dash: bool = False,
    ) -> None:
        if len(pts) < 2:
            return
        if not dash:
            scaled = [(self._sc(px), self._sc(py)) for px, py in pts]
            self.draw.line(scaled, fill=stroke, width=max(1, round(self._sc(width))), joint="curve")
            pts_str = " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)
            self.svg_elems.append(
                f'<polyline points="{pts_str}" fill="none" stroke="{stroke}" '
                f'stroke-width="{width:.1f}" stroke-linejoin="round" stroke-linecap="round"/>'
            )
        else:
            for (x1, y1), (x2, y2) in zip(pts[:-1], pts[1:]):
                self.line(x1, y1, x2, y2, stroke=stroke, width=width, dash=True)

    def polygon(self, pts: Sequence[Tuple[float, float]], fill: str, stroke: str | None = None) -> None:
        scaled = [(self._sc(px), self._sc(py)) for px, py in pts]
        self.draw.polygon(scaled, fill=fill, outline=stroke)
        pts_str = " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)
        s_attr = f' stroke="{stroke}" stroke-width="1"' if stroke else ""
        self.svg_elems.append(f'<polygon points="{pts_str}" fill="{fill}"{s_attr}/>')

    def text(
        self,
        x: float,
        y: float,
        content: str,
        size: int = 16,
        color: str = PALETTE["ink"],
        bold: bool = False,
        mono: bool = False,
        anchor: str = "lt",
    ) -> None:
        lines = content.split("\n")
        if len(lines) > 1:
            line_h = size * 1.28
            total_h = (len(lines) - 1) * line_h
            start_y = y - total_h / 2 if anchor[1] == "m" else y
            for idx, ln in enumerate(lines):
                self.text(x, start_y + idx * line_h, ln, size=size, color=color, bold=bold, mono=mono, anchor=anchor)
            return

        f = self.font(bold=bold, mono=mono, size=size)
        bbox = self.draw.textbbox((0, 0), content, font=f)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        sx, sy = self._sc(x), self._sc(y)

        if anchor[0] == "m":
            px = sx - tw / 2 - bbox[0]
            svg_anchor = "middle"
        elif anchor[0] == "r":
            px = sx - tw - bbox[0]
            svg_anchor = "end"
        else:
            px = sx - bbox[0]
            svg_anchor = "start"

        if anchor[1] == "m":
            py = sy - th / 2 - bbox[1]
            svg_dy = "0.35em"
        elif anchor[1] == "b":
            py = sy - th - bbox[1]
            svg_dy = "0em"
        else:
            py = sy - bbox[1]
            svg_dy = "0.85em"

        self.draw.text((px, py), content, font=f, fill=color)
        fam = "DejaVu Sans Mono, monospace" if mono else "DejaVu Sans, sans-serif"
        weight = "bold" if bold else "normal"
        self.svg_elems.append(
            f'<text x="{x:.1f}" y="{y:.1f}" fill="{color}" font-family="{fam}" '
            f'font-size="{size}" font-weight="{weight}" text-anchor="{svg_anchor}" dy="{svg_dy}">'
            f'{escape(content)}</text>'
        )

    def arrowhead(self, tip: Tuple[float, float], prev: Tuple[float, float], color: str, size: float = 11.0) -> None:
        dx = tip[0] - prev[0]
        dy = tip[1] - prev[1]
        ang = math.atan2(dy, dx)
        a1 = ang + math.pi * 0.84
        a2 = ang - math.pi * 0.84
        p1 = (tip[0] + size * math.cos(a1), tip[1] + size * math.sin(a1))
        p2 = (tip[0] + size * math.cos(a2), tip[1] + size * math.sin(a2))
        self.polygon([tip, p1, p2], fill=color)

    def bezier_arrow(
        self,
        pts_ctrl: Sequence[Tuple[float, float]],
        color: str,
        width: float = 2.4,
        dash: bool = False,
        polarity: str | None = None,
        pol_t: float = 0.5,
        pol_offset: Tuple[float, float] = (0.0, 0.0),
        label: str | None = None,
        label_pos: Tuple[float, float] | None = None,
    ) -> None:
        n_steps = 56
        samples: list[Tuple[float, float]] = []
        if len(pts_ctrl) == 2:
            p0, p1 = pts_ctrl
            for i in range(n_steps + 1):
                t = i / n_steps
                samples.append((p0[0] + (p1[0] - p0[0]) * t, p0[1] + (p1[1] - p0[1]) * t))
        elif len(pts_ctrl) == 3:
            p0, p1, p2 = pts_ctrl
            for i in range(n_steps + 1):
                t = i / n_steps
                mt = 1 - t
                x = mt * mt * p0[0] + 2 * mt * t * p1[0] + t * t * p2[0]
                y = mt * mt * p0[1] + 2 * mt * t * p1[1] + t * t * p2[1]
                samples.append((x, y))
        elif len(pts_ctrl) == 4:
            p0, p1, p2, p3 = pts_ctrl
            for i in range(n_steps + 1):
                t = i / n_steps
                mt = 1 - t
                x = mt**3 * p0[0] + 3 * mt**2 * t * p1[0] + 3 * mt * t**2 * p2[0] + t**3 * p3[0]
                y = mt**3 * p0[1] + 3 * mt**2 * t * p1[1] + 3 * mt * t**2 * p2[1] + t**3 * p3[1]
                samples.append((x, y))
        else:
            raise ValueError("pts_ctrl debe tener entre 2 y 4 puntos.")

        tip = samples[-1]
        prev = samples[-4]
        ang = math.atan2(tip[1] - prev[1], tip[0] - prev[0])
        line_end = (tip[0] - 8.5 * math.cos(ang), tip[1] - 8.5 * math.sin(ang))
        draw_samples = samples[:-2] + [line_end]

        if dash:
            for i in range(0, len(draw_samples) - 1, 2):
                self.line(
                    draw_samples[i][0],
                    draw_samples[i][1],
                    draw_samples[i + 1][0],
                    draw_samples[i + 1][1],
                    stroke=color,
                    width=width,
                )
        else:
            scaled = [(self._sc(px), self._sc(py)) for px, py in draw_samples]
            self.draw.line(scaled, fill=color, width=max(1, round(self._sc(width))), joint="curve")

        if len(pts_ctrl) == 2:
            d_str = f"M {pts_ctrl[0][0]:.1f} {pts_ctrl[0][1]:.1f} L {line_end[0]:.1f} {line_end[1]:.1f}"
        elif len(pts_ctrl) == 3:
            d_str = (
                f"M {pts_ctrl[0][0]:.1f} {pts_ctrl[0][1]:.1f} "
                f"Q {pts_ctrl[1][0]:.1f} {pts_ctrl[1][1]:.1f} {line_end[0]:.1f} {line_end[1]:.1f}"
            )
        else:
            d_str = (
                f"M {pts_ctrl[0][0]:.1f} {pts_ctrl[0][1]:.1f} "
                f"C {pts_ctrl[1][0]:.1f} {pts_ctrl[1][1]:.1f}, "
                f"{pts_ctrl[2][0]:.1f} {pts_ctrl[2][1]:.1f}, {line_end[0]:.1f} {line_end[1]:.1f}"
            )
        dash_attr = ' stroke-dasharray="7 5"' if dash else ""
        self.svg_elems.append(
            f'<path d="{d_str}" fill="none" stroke="{color}" stroke-width="{width:.1f}"{dash_attr} stroke-linecap="round"/>'
        )

        self.arrowhead(tip, prev, color=color, size=11.5)

        if polarity:
            idx = max(1, min(len(samples) - 2, int(pol_t * len(samples))))
            bx = samples[idx][0] + pol_offset[0]
            by = samples[idx][1] + pol_offset[1]
            self.circle(bx, by, r=12.5, fill="#FFFFFF", stroke=color, stroke_w=2.0)
            sym = "+" if polarity == "+" else "−"
            self.text(bx, by - 0.5, sym, size=16, color=color, bold=True, anchor="mm")

        if label and label_pos:
            self.text(
                label_pos[0],
                label_pos[1],
                label,
                size=12,
                color=PALETTE["ink_muted"],
                bold=False,
                anchor="mm",
            )

    def save(self, stem: Path) -> tuple[Path, Path]:
        svg_path = stem.with_suffix(".svg")
        png_path = stem.with_suffix(".png")
        svg_body = "\n  ".join(self.svg_elems)
        svg_full = (
            f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {self.w} {self.h}" '
            f'width="{self.w}" height="{self.h}">\n  {svg_body}\n</svg>\n'
        )
        svg_path.write_text(svg_full, encoding="utf-8")
        out_img = self.img.resize((self.w, self.h), Image.Resampling.LANCZOS)
        out_img.save(png_path, format="PNG", optimize=True)
        return svg_path, png_path


def draw_node_box(
    c: DualCanvas,
    cx: float,
    cy: float,
    w: float,
    h: float,
    title: str,
    subtitle: str,
    fill: str,
    stroke: str,
    title_color: str,
    is_stock: bool = False,
) -> None:
    x, y = cx - w / 2, cy - h / 2
    if is_stock:
        c.rect(x - 4, y - 4, w + 8, h + 8, fill="#FFFFFF", stroke=stroke, stroke_w=1.8, rx=4)
        c.rect(x, y, w, h, fill=fill, stroke=stroke, stroke_w=2.6, rx=3)
    else:
        c.rect(x, y, w, h, fill=fill, stroke=stroke, stroke_w=2.0, rx=4)
    c.text(cx, cy - 11, title, size=14.5, color=title_color, bold=True, anchor="mm")
    c.text(cx, cy + 13, subtitle, size=12, color=PALETTE["ink_muted"], anchor="mm")


def draw_loop_badge(
    c: DualCanvas,
    cx: float,
    cy: float,
    code: str,
    name: str,
    clockwise: bool = True,
    accent: str = PALETTE["loop_stroke"],
) -> None:
    r = 24.0
    c.circle(cx, cy, r=r, fill="#FFFFFF", stroke=accent, stroke_w=2.0)
    arc_r = r + 6.5
    pts: list[Tuple[float, float]] = []
    angles = (
        [math.radians(a) for a in range(-140, 145, 8)]
        if clockwise
        else [math.radians(a) for a in range(140, -145, -8)]
    )
    for ang in angles:
        pts.append((cx + arc_r * math.cos(ang), cy + arc_r * math.sin(ang)))
    c.polyline(pts, stroke=accent, width=2.0)
    c.arrowhead(pts[-1], pts[-3], color=accent, size=8.5)
    c.text(cx, cy, code, size=15.5, color=accent, bold=True, anchor="mm")
    c.text(cx, cy + r + 19, name, size=11, color=PALETTE["ink_muted"], bold=True, anchor="mm")


# ---------------------------------------------------------------------------
# FIGURA 1: Diagrama Causal Completo (CLD)
# ---------------------------------------------------------------------------
def build_cld() -> tuple[Path, Path]:
    c = DualCanvas(1920, 1140)

    # Encabezado
    c.rect(24, 20, 1872, 74, fill="#FFFFFF", stroke=PALETTE["border"], stroke_w=1.5, rx=4)
    c.text(
        48,
        44,
        "FASE 1 · DIAGRAMA CAUSAL (CLD): MODELO AGREGADO DE 2 NIVELES DEL ALGORITMO GENÉTICO",
        size=21,
        color=PALETTE["ink"],
        bold=True,
        anchor="lm",
    )
    c.text(
        48,
        72,
        "Mochila entera acotada (N = 8, L = 8 bits, 4 objetos) · 17 enlaces causales dirigidos · 4 bucles de balance (B1–B4) y acoplamiento inter-nivel",
        size=14,
        color=PALETTE["ink_muted"],
        anchor="lm",
    )

    # Leyenda
    c.circle(1525, 57, r=10, fill="#FFFFFF", stroke=PALETTE["pos"], stroke_w=2.0)
    c.text(1525, 56, "+", size=14, color=PALETTE["pos"], bold=True, anchor="mm")
    c.text(1545, 57, "Polaridad Positiva", size=13, color=PALETTE["ink"], bold=True, anchor="lm")

    c.circle(1715, 57, r=10, fill="#FFFFFF", stroke=PALETTE["neg"], stroke_w=2.0)
    c.text(1715, 56, "−", size=14, color=PALETTE["neg"], bold=True, anchor="mm")
    c.text(1735, 57, "Polaridad Negativa", size=13, color=PALETTE["ink"], bold=True, anchor="lm")

    # Contenedores de Subsistemas
    c.rect(
        44, 212, 650, 742, fill="#FFFFFF", stroke="#93C5FD", stroke_w=1.6, rx=4, dash="8 6"
    )
    c.text(
        66, 232, "SUBSISTEMA 1 · DIVERSIDAD GENÉTICA (VARIACIÓN ALÉLICA)",
        size=13, color=PALETTE["div_text"], bold=True, anchor="lm",
    )

    c.rect(
        935, 212, 740, 742, fill="#FFFFFF", stroke="#6EE7B7", stroke_w=1.6, rx=4, dash="8 6"
    )
    c.text(
        955, 232, "SUBSISTEMA 2 · APTITUD Y BRECHA AL ÓPTIMO",
        size=13, color=PALETTE["qual_text"], bold=True, anchor="lm",
    )

    # Coordenadas cuidadosamente espaciadas (cero traslapes)
    POS_ID = (420, 305)
    POS_D = (420, 580)
    POS_OD = (420, 855)

    POS_B = (1285, 305)
    POS_IQ = (1070, 580)
    POS_Q = (1490, 580)
    POS_OQ = (1490, 855)

    POS_PM = (920, 145)
    POS_PC = (805, 385)
    POS_CAP = (820, 765)
    POS_EL = (970, 1025)

    # ---------------- FLECHAS CAUSALES (17 ENLACES) ----------------
    # E7: D -> ID (-)
    c.bezier_arrow(
        [(295, 558), (225, 442), (295, 325)],
        color=PALETTE["neg"],
        width=2.6,
        polarity="-",
        pol_t=0.50,
        pol_offset=(-18, 0),
        label="Menor espacio\n(1 − D)",
        label_pos=(145, 442),
    )
    # E8: ID -> D (+)
    c.bezier_arrow(
        [(545, 325), (615, 442), (545, 558)],
        color=PALETTE["pos"],
        width=2.6,
        polarity="+",
        pol_t=0.50,
        pol_offset=(18, 0),
        label="Repone\nvariación",
        label_pos=(660, 442),
    )

    # E5: D -> OD (+)
    c.bezier_arrow(
        [(295, 602), (225, 718), (295, 835)],
        color=PALETTE["pos"],
        width=2.6,
        polarity="+",
        pol_t=0.50,
        pol_offset=(-18, 0),
        label="Pérdida\nproporcional",
        label_pos=(145, 718),
    )
    # E6: OD -> D (-)
    c.bezier_arrow(
        [(545, 835), (615, 718), (545, 602)],
        color=PALETTE["neg"],
        width=2.6,
        polarity="-",
        pol_t=0.50,
        pol_offset=(18, 0),
        label="Drena\nstock",
        label_pos=(658, 718),
    )

    # E4: D -> IQ (+) [Acoplamiento de Fisher]
    c.bezier_arrow(
        [(552, 580), (940, 580)],
        color=PALETTE["pos"],
        width=3.2,
        polarity="+",
        pol_t=0.50,
        pol_offset=(0, -18),
        label="Habilita búsqueda (Fisher)",
        label_pos=(755, 606),
    )

    # E3: IQ -> Q (+)
    c.bezier_arrow(
        [(1195, 580), (1358, 580)],
        color=PALETTE["pos"],
        width=3.0,
        polarity="+",
        pol_t=0.50,
        pol_offset=(0, -18),
        label="Eleva calidad",
        label_pos=(1276, 606),
    )

    # E1: Q -> B (-)
    c.bezier_arrow(
        [(1490, 534), (1490, 345), (1422, 305)],
        color=PALETTE["neg"],
        width=2.6,
        polarity="-",
        pol_t=0.50,
        pol_offset=(18, -8),
        label="Reduce brecha",
        label_pos=(1542, 405),
    )

    # E2: B -> IQ (+)
    c.bezier_arrow(
        [(1148, 305), (1070, 345), (1070, 540)],
        color=PALETTE["pos"],
        width=2.6,
        polarity="+",
        pol_t=0.50,
        pol_offset=(18, -8),
        label="Impulsa mejora",
        label_pos=(1155, 400),
    )

    # E9: Q -> OQ (+)
    c.bezier_arrow(
        [(1390, 624), (1325, 720), (1390, 818)],
        color=PALETTE["pos"],
        width=2.6,
        polarity="+",
        pol_t=0.50,
        pol_offset=(-18, 0),
        label="Mayor exposición",
        label_pos=(1250, 720),
    )

    # E10: OQ -> Q (-)
    c.bezier_arrow(
        [(1590, 818), (1650, 720), (1590, 624)],
        color=PALETTE["neg"],
        width=2.6,
        polarity="-",
        pol_t=0.50,
        pol_offset=(18, 0),
        label="Erosiona\nQ(t)",
        label_pos=(1688, 720),
    )

    # E11: Pm -> ID (+)
    c.bezier_arrow(
        [(800, 152), (650, 175), (535, 270)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.52,
        pol_offset=(-8, -15),
        label="Explorador",
        label_pos=(655, 162),
    )

    # E12: Pm -> OQ (+)
    c.bezier_arrow(
        [(1040, 145), (1835, 130), (1845, 845), (1620, 855)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.45,
        pol_offset=(18, 0),
        label="Disruptivo\n(carga mutacional)",
        label_pos=(1848, 535),
    )

    # E13: Pc -> IQ (+)
    c.bezier_arrow(
        [(805, 420), (895, 488), (1010, 540)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.48,
        pol_offset=(-14, 12),
    )

    # E17: Cap / Phi_fact -> IQ (+)
    c.bezier_arrow(
        [(820, 730), (910, 668), (1000, 618)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.48,
        pol_offset=(-14, -12),
    )

    # E14: El -> IQ (+)
    c.bezier_arrow(
        [(985, 990), (1045, 800), (1068, 620)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.55,
        pol_offset=(-16, 0),
    )

    # E15: El -> OD (+)
    c.bezier_arrow(
        [(854, 1005), (710, 955), (574, 875)],
        color=PALETTE["pos"],
        width=2.2,
        dash=True,
        polarity="+",
        pol_t=0.50,
        pol_offset=(-10, 16),
    )

    # E16: El -> OQ (-)
    c.bezier_arrow(
        [(1086, 1005), (1235, 955), (1360, 875)],
        color=PALETTE["neg"],
        width=2.2,
        dash=True,
        polarity="-",
        pol_t=0.50,
        pol_offset=(12, 16),
    )

    # ---------------- BUCLES DE RETROALIMENTACIÓN ----------------
    draw_loop_badge(c, 420, 442, "B3", "Reposición Mutacional", clockwise=True, accent=PALETTE["div_stroke"])
    draw_loop_badge(c, 420, 718, "B2", "Agotamiento Selección", clockwise=False, accent=PALETTE["neg"])
    draw_loop_badge(c, 1285, 442, "B1", "Cierre de Brecha", clockwise=False, accent=PALETTE["qual_stroke"])
    draw_loop_badge(c, 1490, 720, "B4", "Carga Mutacional", clockwise=False, accent=PALETTE["neg"])

    # ---------------- CAJAS DE NODOS ----------------
    draw_node_box(
        c, *POS_D, 256, 84,
        "Diversidad Genética D(t)",
        "Nivel 1 · Varianza alélica [0, 1]",
        PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"], is_stock=True,
    )
    draw_node_box(
        c, *POS_Q, 258, 84,
        "Calidad Media Q(t)",
        "Nivel 2 · Aptitud norm. F̄/Z* [0, 1]",
        PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"], is_stock=True,
    )

    draw_node_box(
        c, *POS_ID, 256, 70,
        "Inyección por Mutación I_D",
        "Flujo de entrada a D [1/gen]",
        PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"],
    )
    draw_node_box(
        c, *POS_OD, 304, 70,
        "Consumo Selección y Deriva O_D",
        "Flujo de salida de D [1/gen]",
        PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"],
    )
    draw_node_box(
        c, *POS_B, 272, 70,
        "Brecha al Óptimo B(t)",
        "Backlog: B(t) = 1 − Q(t)",
        PALETTE["aux_fill"], PALETTE["aux_stroke"], PALETTE["aux_text"],
    )
    draw_node_box(
        c, *POS_IQ, 252, 76,
        "Ganancia de Aptitud I_Q",
        "Flujo de entrada a Q [1/gen]",
        PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"],
    )
    draw_node_box(
        c, *POS_OQ, 258, 70,
        "Carga Mutacional O_Q",
        "Flujo de salida de Q [1/gen]",
        PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"],
    )

    draw_node_box(
        c, *POS_PM, 240, 68,
        "Tasa de Mutación P_mut",
        "Parámetro AG (base = 0.15)",
        PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"],
    )
    draw_node_box(
        c, *POS_PC, 224, 66,
        "Prob. de Cruce P_cr",
        "Parámetro AG (base = 0.89)",
        PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"],
    )
    draw_node_box(
        c, *POS_CAP, 246, 68,
        "Capacidad C y Fact. Φ_fact",
        "Restricción mochila (140 / 55)",
        PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"],
    )
    draw_node_box(
        c, *POS_EL, 232, 68,
        "Elitismo E ∈ {0, 1}",
        "Protege mejor individuo (base = 1)",
        PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"],
    )

    c.text(
        48,
        1112,
        "Nota: Bucle B = signo negativo neto (balance/estabilización). Todos los bucles B1–B4 son estabilizadores; la convergencia prematura ocurre cuando B2 vacía D(t) y apaga el enlace D → I_Q.",
        size=12.5,
        color=PALETTE["ink_muted"],
        anchor="lm",
    )

    return c.save(FIG_DIR / "diagrama_causal_cld")


# ---------------------------------------------------------------------------
# FIGURA 2: Modos de Referencia y Plano de Fase
# ---------------------------------------------------------------------------
def simulate_reference_modes() -> dict[str, tuple[list[float], list[float], list[float]]]:
    """
    Trayectorias suaves características de los Regímenes de Referencia del AG (T = 30 gen.)
    partiendo del estado inicial observado en Mochila.xlsx (D0 = 0.9375, Q0 = 0.5071).
    """
    k_sel, k_e, k_q, k_eq, k_load = 0.0419, 0.7613, 0.1279, 0.2149, 1.0518
    N, pc, phi, q_rand = 8.0, 0.89, 1.0, 0.50
    T = 30

    configs = {
        "Reg1": (0.00, 1.0),   # Agotamiento genético / Estancamiento prematuro
        "Reg2a": (0.05, 1.0),  # Equilibrio óptimo exploración - carga
        "Reg2b": (0.15, 1.0),  # Caso base Mochila.xlsx (alta exploración + carga en Q̄)
        "Reg3": (0.15, 0.0),   # Degradación por carga mutacional sin elitismo
    }
    out: dict[str, tuple[list[float], list[float], list[float]]] = {}
    for key, (pm, E) in configs.items():
        d_list = [0.9375]
        q_list = [0.5071]
        b_list = [1.0 - 0.5071]
        for _ in range(T):
            d, q = d_list[-1], q_list[-1]
            i_d = 4.0 * pm * (1.0 - pm) * (1.0 - d)
            o_d = ((1.0 - E / N) / N + k_sel + k_e * (E / N)) * d
            d_next = max(0.0, min(1.0, d + i_d - o_d))

            i_q = (k_q * (1.0 + 0.25 * pc) + k_eq * E) * d * (1.0 - q) * phi
            o_q = k_load * (1.0 - (1.0 - pm) ** 2) * (1.0 - E / N) * max(0.0, q - q_rand)
            q_next = max(0.0, min(1.0, q + i_q - o_q))

            d_list.append(d_next)
            q_list.append(q_next)
            b_list.append(1.0 - q_next)
        out[key] = (d_list, q_list, b_list)
    return out


def draw_axes_panel(
    c: DualCanvas,
    px: float,
    py: float,
    pw: float,
    ph: float,
    title: str,
    ylabel: str,
    xlabel: str,
    x_ticks: Sequence[float],
    y_ticks: Sequence[float],
    x_min: float,
    x_max: float,
    y_min: float,
    y_max: float,
) -> tuple[float, float, float, float]:
    c.rect(px, py, pw, ph, fill="#FFFFFF", stroke=PALETTE["border"], stroke_w=1.5, rx=4)
    c.text(px + 20, py + 24, title, size=15.5, color=PALETTE["ink"], bold=True, anchor="lm")

    gx, gy = px + 92, py + 56
    gw, gh = pw - 122, ph - 114

    for yt in y_ticks:
        if yt < y_min - 1e-6 or yt > y_max + 1e-6:
            continue
        frac = (yt - y_min) / (y_max - y_min)
        yy = gy + gh - frac * gh
        c.line(gx, yy, gx + gw, yy, stroke=PALETTE["grid"], width=1.1)
        c.text(gx - 10, yy, f"{yt:.2f}", size=11.5, color=PALETTE["ink_muted"], mono=True, anchor="rm")

    for xt in x_ticks:
        if xt < x_min - 1e-6 or xt > x_max + 1e-6:
            continue
        frac = (xt - x_min) / (x_max - x_min)
        xx = gx + frac * gw
        c.line(xx, gy, xx, gy + gh, stroke=PALETTE["grid"], width=1.1)
        lbl = f"{int(xt)}" if float(xt).is_integer() and x_max > 2 else f"{xt:.2f}"
        c.text(xx, gy + gh + 16, lbl, size=11.5, color=PALETTE["ink_muted"], mono=True, anchor="mm")

    c.line(gx, gy, gx, gy + gh, stroke=PALETTE["ink"], width=1.8)
    c.line(gx, gy + gh, gx + gw, gy + gh, stroke=PALETTE["ink"], width=1.8)

    c.text(px + 18, gy + gh / 2, ylabel, size=12.5, color=PALETTE["ink"], bold=True, anchor="lm")
    c.text(gx + gw / 2, gy + gh + 38, xlabel, size=12.5, color=PALETTE["ink"], bold=True, anchor="mm")
    return gx, gy, gw, gh


def map_pt(
    x: float,
    y: float,
    gx: float,
    gy: float,
    gw: float,
    gh: float,
    x_min: float,
    x_max: float,
    y_min: float,
    y_max: float,
) -> Tuple[float, float]:
    mx = gx + ((x - x_min) / (x_max - x_min)) * gw
    my = gy + gh - ((y - y_min) / (y_max - y_min)) * gh
    return (mx, my)


def build_reference_modes() -> tuple[Path, Path]:
    c = DualCanvas(1820, 1110)
    data = simulate_reference_modes()
    t_vals = list(range(31))

    # Encabezado
    c.rect(24, 20, 1772, 102, fill="#FFFFFF", stroke=PALETTE["border"], stroke_w=1.5, rx=4)
    c.text(
        48,
        44,
        "FASE 1 · MODOS DE REFERENCIA DINÁMICOS Y RETRATO DE FASE DEL ALGORITMO GENÉTICO",
        size=20,
        color=PALETTE["ink"],
        bold=True,
        anchor="lm",
    )
    c.text(
        48,
        71,
        "Comportamiento típico sobre t = 0..30 generaciones (C = 140) y anclaje empírico con Mochila.xlsx (Gen 0, Gen 1, Gen 2)",
        size=13.5,
        color=PALETTE["ink_muted"],
        anchor="lm",
    )

    legend_items = [
        (48, 102, PALETTE["reg2b"], "Régimen II (Base): P_mut=0.15, E=1"),
        (420, 102, PALETTE["reg2a"], "Régimen II-opt: P_mut=0.05, E=1"),
        (785, 102, PALETTE["reg1"], "Régimen I (Colapso): P_mut=0.00, E=1"),
        (1175, 102, PALETTE["reg3"], "Régimen III (Sin elitismo): P_mut=0.15, E=0"),
    ]
    for lx, ly, col, lbl in legend_items:
        c.line(lx, ly, lx + 26, ly, stroke=col, width=3.5)
        c.text(lx + 34, ly, lbl, size=12, color=PALETTE["ink"], bold=True, anchor="lm")
    c.circle(1570, 102, r=5.0, fill="#FFFFFF", stroke=PALETTE["obs"], stroke_w=2.2)
    c.text(1583, 102, "Mochila.xlsx (Gen 0–2)", size=12, color=PALETTE["obs"], bold=True, anchor="lm")

    obs_t = [0, 1, 2]
    obs_D = [0.9375, 0.7969, 0.8750]
    obs_Q = [0.5071, 0.6643, 0.6321]
    obs_B = [1.0 - q for q in obs_Q]

    # Panel A: Diversidad D(t)
    gx, gy, gw, gh = draw_axes_panel(
        c, 24, 140, 874, 455,
        "A. Nivel 1 — Diversidad Genética D(t)",
        "D(t)", "Generación t",
        [0, 5, 10, 15, 20, 25, 30], [0.00, 0.25, 0.50, 0.75, 1.00],
        0, 30, 0.0, 1.08,
    )
    for key, col in [
        ("Reg1", PALETTE["reg1"]),
        ("Reg3", PALETTE["reg3"]),
        ("Reg2b", PALETTE["reg2b"]),
        ("Reg2a", PALETTE["reg2a"]),
    ]:
        pts = [map_pt(t, val, gx, gy, gw, gh, 0, 30, 0.0, 1.08) for t, val in zip(t_vals, data[key][0])]
        c.polyline(pts, stroke=col, width=3.0)

    for key, col, dy, txt in [
        ("Reg3", PALETTE["reg3"], -16, f"D*_eq ≈ {data['Reg3'][0][-1]:.2f} (E=0)"),
        ("Reg2b", PALETTE["reg2b"], 18, f"D*_eq ≈ {data['Reg2b'][0][-1]:.2f} (Base P_mut=0.15)"),
        ("Reg2a", PALETTE["reg2a"], -16, f"D*_eq ≈ {data['Reg2a'][0][-1]:.2f} (P_mut=0.05)"),
        ("Reg1", PALETTE["reg1"], -16, "D → 0 (colapso de diversidad)"),
    ]:
        px_e, py_e = map_pt(30, data[key][0][-1], gx, gy, gw, gh, 0, 30, 0.0, 1.08)
        c.text(px_e - 8, py_e + dy, txt, size=11.5, color=col, bold=True, anchor="rm")

    for ot, od in zip(obs_t, obs_D):
        ox, oy = map_pt(ot, od, gx, gy, gw, gh, 0, 30, 0.0, 1.08)
        c.circle(ox, oy, r=5.5, fill="#FFFFFF", stroke=PALETTE["obs"], stroke_w=2.4)

    # Panel B: Calidad Media Q(t)
    gx, gy, gw, gh = draw_axes_panel(
        c, 922, 140, 874, 455,
        "B. Nivel 2 — Calidad Media Normalizada Q(t)",
        "Q(t)", "Generación t",
        [0, 5, 10, 15, 20, 25, 30], [0.00, 0.25, 0.50, 0.75, 1.00],
        0, 30, 0.0, 1.08,
    )
    y_opt = map_pt(0, 1.0, gx, gy, gw, gh, 0, 30, 0.0, 1.08)[1]
    c.line(gx, y_opt, gx + gw, y_opt, stroke=PALETTE["ink_light"], width=1.4, dash=True)
    c.text(gx + gw - 8, y_opt - 11, "Óptimo global Q = 1.00 (Z* = 70)", size=11, color=PALETTE["ink_muted"], anchor="rm")
    y_azar = map_pt(0, 0.5, gx, gy, gw, gh, 0, 30, 0.0, 1.08)[1]
    c.line(gx, y_azar, gx + gw, y_azar, stroke=PALETTE["border_strong"], width=1.3, dash=True)
    c.text(gx + gw - 8, y_azar + 14, "Nivel aleatorio Q_azar = 0.50", size=11, color=PALETTE["ink_muted"], anchor="rm")

    for key, col in [
        ("Reg1", PALETTE["reg1"]),
        ("Reg3", PALETTE["reg3"]),
        ("Reg2b", PALETTE["reg2b"]),
        ("Reg2a", PALETTE["reg2a"]),
    ]:
        pts = [map_pt(t, val, gx, gy, gw, gh, 0, 30, 0.0, 1.08) for t, val in zip(t_vals, data[key][1])]
        c.polyline(pts, stroke=col, width=3.0)

    # Etiquetas bien separadas explicando Q_media vs Q_max
    p_r1 = map_pt(30, data["Reg1"][1][-1], gx, gy, gw, gh, 0, 30, 0.0, 1.08)
    c.text(p_r1[0] - 8, p_r1[1] - 14, "Clonación Q̄ = Q_max ≈ 0.90 (estancado < 1.00)", size=11.5, color=PALETTE["reg1"], bold=True, anchor="rm")
    p_r2a = map_pt(30, data["Reg2a"][1][-1], gx, gy, gw, gh, 0, 30, 0.0, 1.08)
    c.text(p_r2a[0] - 8, p_r2a[1] - 13, f"Equilibrio óptimo Q̄ ≈ {data['Reg2a'][1][-1]:.2f} (Q_max → 1.00)", size=11.5, color=PALETTE["reg2a"], bold=True, anchor="rm")
    p_r2b = map_pt(30, data["Reg2b"][1][-1], gx, gy, gw, gh, 0, 30, 0.0, 1.08)
    c.text(p_r2b[0] - 8, p_r2b[1] + 15, f"Base Q̄ ≈ {data['Reg2b'][1][-1]:.2f} (Q_max → 1.00 + carga mut.)", size=11.5, color=PALETTE["reg2b"], bold=True, anchor="rm")
    p_r3 = map_pt(30, data["Reg3"][1][-1], gx, gy, gw, gh, 0, 30, 0.0, 1.08)
    c.text(p_r3[0] - 8, p_r3[1] + 15, f"Degradado Q̄ ≈ {data['Reg3'][1][-1]:.2f} (sin elitismo)", size=11.5, color=PALETTE["reg3"], bold=True, anchor="rm")

    for ot, oq in zip(obs_t, obs_Q):
        ox, oy = map_pt(ot, oq, gx, gy, gw, gh, 0, 30, 0.0, 1.08)
        c.circle(ox, oy, r=5.5, fill="#FFFFFF", stroke=PALETTE["obs"], stroke_w=2.4)

    # Panel C: Brecha / Backlog B(t) = 1 - Q(t) [con y_ticks dentro de [0, 0.80]]
    gx, gy, gw, gh = draw_axes_panel(
        c, 24, 616, 874, 468,
        "C. Brecha / Backlog al Óptimo B(t) = 1 − Q(t)",
        "B(t)", "Generación t",
        [0, 5, 10, 15, 20, 25, 30], [0.00, 0.20, 0.40, 0.60, 0.80],
        0, 30, 0.0, 0.80,
    )
    for key, col in [
        ("Reg1", PALETTE["reg1"]),
        ("Reg3", PALETTE["reg3"]),
        ("Reg2b", PALETTE["reg2b"]),
        ("Reg2a", PALETTE["reg2a"]),
    ]:
        pts = [map_pt(t, val, gx, gy, gw, gh, 0, 30, 0.0, 0.80) for t, val in zip(t_vals, data[key][2])]
        c.polyline(pts, stroke=col, width=3.0)
    for ot, ob in zip(obs_t, obs_B):
        ox, oy = map_pt(ot, ob, gx, gy, gw, gh, 0, 30, 0.0, 0.80)
        c.circle(ox, oy, r=5.5, fill="#FFFFFF", stroke=PALETTE["obs"], stroke_w=2.4)
    c.text(gx + gw - 12, gy + 24, "A menor B(t), menor impulso de ganancia I_Q (Bucle B1)", size=12, color=PALETTE["ink_muted"], anchor="rm")

    # Panel D: Retrato de Fase (D vs Q)
    gx, gy, gw, gh = draw_axes_panel(
        c, 922, 616, 874, 468,
        "D. Retrato de Fase D(t) vs Q(t) — Atractores del Sistema",
        "Q(t)", "Diversidad Genética D(t)",
        [0.00, 0.25, 0.50, 0.75, 1.00], [0.25, 0.50, 0.75, 1.00],
        0.0, 1.06, 0.25, 1.06,
    )
    for key, col in [
        ("Reg1", PALETTE["reg1"]),
        ("Reg3", PALETTE["reg3"]),
        ("Reg2b", PALETTE["reg2b"]),
        ("Reg2a", PALETTE["reg2a"]),
    ]:
        ds, qs, _ = data[key]
        pts = [map_pt(d, q, gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06) for d, q in zip(ds, qs)]
        c.polyline(pts, stroke=col, width=3.0)
        c.circle(pts[-1][0], pts[-1][1], r=6.5, fill=col, stroke="#FFFFFF", stroke_w=2.0)
        mid = max(2, len(pts) // 4)
        c.arrowhead(pts[mid + 1], pts[mid], color=col, size=10.5)

    p0 = map_pt(0.9375, 0.5071, gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06)
    c.circle(p0[0], p0[1], r=7.0, fill="#FFFFFF", stroke=PALETTE["obs"], stroke_w=2.6)
    c.text(p0[0] - 8, p0[1] + 22, "Inicio t=0 (0.94, 0.51)", size=11.5, color=PALETTE["obs"], bold=True, anchor="rm")

    p_a1 = map_pt(data["Reg1"][0][-1], data["Reg1"][1][-1], gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06)
    c.text(p_a1[0] + 12, p_a1[1] - 14, f"Atractor I (D → 0, Q̄={data['Reg1'][1][-1]:.2f})", size=11.5, color=PALETTE["reg1"], bold=True, anchor="lm")
    p_a2a = map_pt(data["Reg2a"][0][-1], data["Reg2a"][1][-1], gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06)
    c.text(p_a2a[0] + 12, p_a2a[1] - 16, f"Atractor II-opt ({data['Reg2a'][0][-1]:.2f}, {data['Reg2a'][1][-1]:.2f})", size=11.5, color=PALETTE["reg2a"], bold=True, anchor="lm")
    p_a2b = map_pt(data["Reg2b"][0][-1], data["Reg2b"][1][-1], gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06)
    c.text(p_a2b[0] + 12, p_a2b[1] - 10, f"Atractor II Base ({data['Reg2b'][0][-1]:.2f}, {data['Reg2b'][1][-1]:.2f})", size=11.5, color=PALETTE["reg2b"], bold=True, anchor="lm")
    p_a3 = map_pt(data["Reg3"][0][-1], data["Reg3"][1][-1], gx, gy, gw, gh, 0.0, 1.06, 0.25, 1.06)
    c.text(p_a3[0] - 12, p_a3[1] + 18, f"Atractor III ({data['Reg3'][0][-1]:.2f}, {data['Reg3'][1][-1]:.2f}, E=0)", size=11.5, color=PALETTE["reg3"], bold=True, anchor="rm")

    return c.save(FIG_DIR / "modos_referencia_ag")


# ---------------------------------------------------------------------------
# FIGURA 3: Arquitectura Micro-AG <-> Macro-DS
# ---------------------------------------------------------------------------
def build_architecture() -> tuple[Path, Path]:
    c = DualCanvas(1800, 880)

    c.rect(24, 20, 1752, 74, fill="#FFFFFF", stroke=PALETTE["border"], stroke_w=1.5, rx=4)
    c.text(
        48,
        44,
        "FASE 1 · ARQUITECTURA DE ACOPLAMIENTO MICRO–MACRO (ALGORITMO GENÉTICO ↔ DINÁMICA DE SISTEMAS)",
        size=20,
        color=PALETTE["ink"],
        bold=True,
        anchor="lm",
    )
    c.text(
        48,
        72,
        "Traducción bidireccional entre operadores discretos por generación (Excel / Python) y niveles agregados de Dinámica de Sistemas",
        size=13.5,
        color=PALETTE["ink_muted"],
        anchor="lm",
    )

    cols = [
        (36, 120, 520, 724, "#FFFFFF", "#94A3B8", "NIVEL MICRO · ALGORITMO GENÉTICO (t → t+1)", PALETTE["ink"]),
        (640, 120, 520, 724, "#FFFFFF", "#0F766E", "INTERFAZ DE AGREGACIÓN Y PAISAJE", PALETTE["aux_text"]),
        (1244, 120, 520, 724, "#FFFFFF", "#1D4ED8", "NIVEL MACRO · DINÁMICA DE SISTEMAS (2 STOCKS)", PALETTE["div_text"]),
    ]
    for x, y, w, h, fill, stroke, hdr, hcol in cols:
        c.rect(x, y, w, h, fill=fill, stroke=stroke, stroke_w=1.8, rx=4)
        c.rect(x, y, w, 46, fill=PALETTE["surface_alt"], stroke=stroke, stroke_w=1.8, rx=4)
        c.text(x + w / 2, y + 23, hdr, size=14, color=hcol, bold=True, anchor="mm")

    micro_blocks = [
        (296, 232, 460, 86, "1. Población Binaria (N = 8, L = 8 bits)", "Cromosomas b₁..b₈ → fenotipo (x₁, x₂, x₃, x₄)", PALETTE["surface_alt"], PALETTE["border_strong"], PALETTE["ink"]),
        (296, 344, 460, 96, "2. Funciones Internas de Mochila", "Z(x) = 4x₁+5x₂+6x₃+3x₄  |  W(x) = 7x₁+6x₂+8x₃+2x₄ ≤ C\nPena de muerte: F(b) = Z si W ≤ C, si no 0", PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"]),
        (296, 464, 460, 94, "3. Selección (Ruleta) y Cruce 1-Punto", "Prob_i = F_i / ΣF  |  P_cr = 0.89 (4 parejas)\nExplotación de bloques constructivos", PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"]),
        (296, 582, 460, 94, "4. Mutación Bit-Flip y Elitismo", "Inversión bit a bit (P_mut = 0.15)\nElitismo E ∈ {0, 1}: reserva fila 1 al mejor F_max", PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"]),
        (296, 704, 460, 96, "5. Nueva Población t + 1 (Excel / Python)", "Salidas por generación: Z_max(t), Z̄_ef(t), Frec. p_j(t)\nAnclaje de verificación exacta celda a celda", PALETTE["surface_alt"], PALETTE["ink_light"], PALETTE["ink"]),
    ]
    for bx, by, bw, bh, title, sub, fill, stroke, col in micro_blocks:
        c.rect(bx - bw / 2, by - bh / 2, bw, bh, fill=fill, stroke=stroke, stroke_w=1.8, rx=4)
        lines = sub.split("\n")
        if len(lines) == 1:
            c.text(bx, by - 14, title, size=14.5, color=col, bold=True, anchor="mm")
            c.text(bx, by + 13, lines[0], size=12, color=PALETTE["ink_muted"], anchor="mm")
        else:
            c.text(bx, by - 23, title, size=14.5, color=col, bold=True, anchor="mm")
            c.text(bx, by + 4, lines[0], size=12, color=PALETTE["ink_muted"], anchor="mm")
            c.text(bx, by + 23, lines[1], size=12, color=PALETTE["ink_muted"], anchor="mm")

    for y_top, y_bot in [(275, 296), (392, 417), (511, 535), (629, 656)]:
        c.bezier_arrow([(296, y_top), (296, y_bot)], color=PALETTE["ink_muted"], width=2.0)

    mid_blocks = [
        (900, 244, 460, 102, "Métrica Observable D_AG(t)", "D_AG(t) = (1/8) Σ 4·p_j(t)·(1 − p_j(t))\nMide heterogeneidad genética media en [0, 1]", PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"]),
        (900, 384, 460, 102, "Métrica Observable Q_AG(t) y Brecha B(t)", "Q_AG(t) = F̄(t) / Z*_fact   |   B_AG(t) = 1 − Q_AG(t)\nMide progreso promedio relativo al óptimo factible", PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"]),
        (900, 536, 460, 114, "Parámetros Estructurales del Paisaje (2⁸ = 256)", "Φ_fact(C): fracción factible (1.000 en C=140; 0.688 en C=55)\nQ_azar(C): calidad esperada aleatoria (0.500 vs 0.402)\nZ*_fact(C): óptimo factible (70 vs 46)", PALETTE["param_fill"], PALETTE["param_stroke"], PALETTE["param_text"]),
        (900, 698, 460, 110, "Criterio Cuantitativo de Validación (Fase 6)", "Comparación AG promedio (Monte Carlo) vs Modelo DS:\nRMSE(Q) ≤ 0.05   y   RMSE(D) ≤ 0.05\nPruebas estadísticas α = 0.05 (Fase 7)", PALETTE["aux_fill"], PALETTE["aux_stroke"], PALETTE["aux_text"]),
    ]
    for bx, by, bw, bh, title, sub, fill, stroke, col in mid_blocks:
        c.rect(bx - bw / 2, by - bh / 2, bw, bh, fill=fill, stroke=stroke, stroke_w=1.8, rx=4)
        lines = sub.split("\n")
        start_y = by - (len(lines) * 10 + 12)
        c.text(bx, start_y, title, size=14.5, color=col, bold=True, anchor="mm")
        for idx, ln in enumerate(lines):
            c.text(bx, start_y + 24 + idx * 20, ln, size=12, color=PALETTE["ink_muted"], anchor="mm")

    macro_blocks = [
        (1504, 244, 460, 108, "Stock 1: Diversidad Genética D(t)", "D(t) = D(t−1) + (I_D − O_D)·Δt\nI_D (Mutación, B3) vs O_D (Selección + Deriva + E, B2)", PALETTE["div_fill"], PALETTE["div_stroke"], PALETTE["div_text"], True),
        (1504, 396, 460, 112, "Stock 2: Calidad Media Normalizada Q(t)", "Q(t) = Q(t−1) + (I_Q − O_Q)·Δt\nI_Q (Fisher: D·B·Φ^α, B1) vs O_Q (Carga mut., B4)", PALETTE["qual_fill"], PALETTE["qual_stroke"], PALETTE["qual_text"], True),
        (1504, 548, 460, 104, "Variables Auxiliares y Backlog B(t)", "B(t) = 1 − Q(t) actúa como motor del cierre de brecha\nEntradas exógenas: P_mut, P_cr, E, C, N", PALETTE["aux_fill"], PALETTE["aux_stroke"], PALETTE["aux_text"], False),
        (1504, 698, 460, 110, "Predicción Dinámica de Regímenes", "1. Colapso si P_mut = 0 (D → 0 ⇒ estancamiento)\n2. Equilibrio D*_eq > 0 y Q̄ + Q_max óptimos con E = 1\n3. Degradación estacionaria con P_mut alto y E = 0", PALETTE["surface_alt"], PALETTE["ink"], PALETTE["ink"], False),
    ]
    for bx, by, bw, bh, title, sub, fill, stroke, col, stock in macro_blocks:
        if stock:
            c.rect(bx - bw / 2 - 4, by - bh / 2 - 4, bw + 8, bh + 8, fill="#FFFFFF", stroke=stroke, stroke_w=1.6, rx=4)
        c.rect(bx - bw / 2, by - bh / 2, bw, bh, fill=fill, stroke=stroke, stroke_w=2.2 if stock else 1.8, rx=4)
        lines = sub.split("\n")
        start_y = by - (len(lines) * 10 + 12)
        c.text(bx, start_y, title, size=14.5, color=col, bold=True, anchor="mm")
        for idx, ln in enumerate(lines):
            c.text(bx, start_y + 24 + idx * 20, ln, size=12, color=PALETTE["ink_muted"], anchor="mm")

    # Flechas conectoras tocando exactamente los bordes de los bloques (x=526 -> x=670, x=1130 -> x=1274)
    c.bezier_arrow([(526, 688), (586, 688), (586, 244), (670, 244)], color=PALETTE["div_stroke"], width=2.4)
    c.bezier_arrow([(526, 718), (608, 718), (608, 384), (670, 384)], color=PALETTE["qual_stroke"], width=2.4)
    c.bezier_arrow([(526, 344), (566, 344), (566, 536), (670, 536)], color=PALETTE["param_stroke"], width=2.2, dash=True)

    c.bezier_arrow([(1130, 244), (1274, 244)], color=PALETTE["div_stroke"], width=2.6)
    c.bezier_arrow([(1130, 384), (1274, 384)], color=PALETTE["qual_stroke"], width=2.6)
    c.bezier_arrow([(1130, 536), (1274, 548)], color=PALETTE["param_stroke"], width=2.2, dash=True)
    c.bezier_arrow([(1274, 698), (1130, 698)], color=PALETTE["aux_stroke"], width=2.6)

    c.text(
        36,
        862,
        "Frontera del modelo: los 8 bits individuales quedan encapsulados dentro del Nivel Micro y se resumen en el Nivel Macro mediante dos stocks continuos [D(t), Q(t)].",
        size=12.5,
        color=PALETTE["ink_muted"],
        anchor="lm",
    )

    return c.save(FIG_DIR / "arquitectura_subsistemas")


def main() -> None:
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    for svg_p, png_p in [build_cld(), build_reference_modes(), build_architecture()]:
        print(f"Generado: {svg_p.relative_to(BASE_DIR)} | {png_p.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()
