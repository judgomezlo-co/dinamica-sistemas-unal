"""
Diagrama de Forrester del modelo agregado de 2 niveles (Diversidad D y Calidad Q).

Dibuja niveles (rectangulos), flujos (tuberias con valvula), fuentes/sumideros (nubes),
variables auxiliares (circulos), parametros exogenos y conectores de informacion,
siguiendo exactamente las ecuaciones del modelo alineado al curso:

    I_D = r_mut (1 - D)          r_mut = c_m P_mut
    O_D = k_D D                  k_D   = 1/N + k_sel + k_e E/N
    I_Q = k_Q D B,  B = 1 - Q    k_Q   = (k_base + k_cr P_cr + k_el E) Phi_fact
    O_Q = k_deg max(0, Q - Q_azar)   k_deg = k_carga P_mut (1 - E/N)

Uso: python src/diagrama_forrester.py [ruta_salida.png]
"""

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch, FancyBboxPatch, Polygon

plt.rcParams.update({
    "font.family": "Arial",
    "mathtext.fontset": "custom",
    "mathtext.rm": "Arial",
    "mathtext.it": "Arial:italic",
    "mathtext.bf": "Arial:bold",
})

AZUL, AZUL_F = "#1E3A8A", "#EEF4FF"
VERDE, VERDE_F = "#065F46", "#ECFDF5"
NARANJA, NARANJA_F = "#9A3412", "#FFF7ED"
GRIS = "#374151"


def nube(ax, x, y, color):
    for dx, dy, r in [(-0.32, 0.0, 0.30), (0.0, 0.14, 0.34), (0.32, 0.0, 0.30), (0.0, -0.14, 0.30)]:
        ax.add_patch(Circle((x + dx, y + dy), r, facecolor="white", edgecolor=color, lw=1.4, zorder=2))
    for dx, dy, r in [(-0.32, 0.0, 0.27), (0.0, 0.14, 0.31), (0.32, 0.0, 0.27), (0.0, -0.14, 0.27)]:
        ax.add_patch(Circle((x + dx, y + dy), r, facecolor="white", edgecolor="none", zorder=3))


def tuberia(ax, x0, x1, y, color, flecha=True):
    ax.plot([x0, x1], [y + 0.09, y + 0.09], color=color, lw=1.8, zorder=1, solid_capstyle="butt")
    ax.plot([x0, x1], [y - 0.09, y - 0.09], color=color, lw=1.8, zorder=1, solid_capstyle="butt")
    if flecha:
        ax.add_patch(Polygon([[x1, y + 0.26], [x1 + 0.34, y], [x1, y - 0.26]], closed=True,
                             facecolor=color, edgecolor=color, zorder=2))


def valvula(ax, x, y, color, etiqueta, ecuacion, abajo=True):
    ax.add_patch(Polygon([[x - 0.28, y + 0.36], [x + 0.28, y + 0.36], [x - 0.28, y - 0.36],
                          [x + 0.28, y - 0.36]], closed=True, facecolor="white",
                         edgecolor=color, lw=1.6, zorder=4))
    ax.plot([x, x], [y, y + 0.62], color=color, lw=1.4, zorder=4)
    ax.add_patch(Circle((x, y + 0.74), 0.13, facecolor=color, edgecolor=color, zorder=5))
    dy = -0.62 if abajo else 1.22
    ax.text(x, y + dy, etiqueta, ha="center", va="center", fontsize=10.5, weight="bold",
            color=color, zorder=6)
    ax.text(x, y + dy - 0.36, ecuacion, ha="center", va="center", fontsize=9, color=GRIS, zorder=6)


def nivel(ax, x, y, w, h, color, fondo, titulo, subtitulo, ecuacion):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="square,pad=0",
                                facecolor=fondo, edgecolor=color, lw=2.2, zorder=4))
    ax.text(x, y + 0.26, titulo, ha="center", va="center", fontsize=11, weight="bold", color=color, zorder=5)
    ax.text(x, y - 0.12, subtitulo, ha="center", va="center", fontsize=9, color=GRIS, zorder=5)
    ax.text(x, y - 0.42, ecuacion, ha="center", va="center", fontsize=9, color=GRIS, zorder=5)


def auxiliar(ax, x, y, etiqueta, color, r=0.42):
    ax.add_patch(Circle((x, y), r, facecolor="white", edgecolor=color, lw=1.5, zorder=4))
    ax.text(x, y, etiqueta, ha="center", va="center", fontsize=10, weight="bold", color=color, zorder=5)


def parametro(ax, x, y, texto):
    ax.text(x, y, texto, ha="center", va="center", fontsize=8.6, color=NARANJA, zorder=5,
            bbox=dict(boxstyle="round,pad=0.28", facecolor=NARANJA_F, edgecolor=NARANJA, lw=1.0))


def enlace(ax, p0, p1, color=GRIS, rad=0.0, estilo=(0, (4, 2.5)), lw=1.25):
    ax.add_patch(FancyArrowPatch(p0, p1, connectionstyle=f"arc3,rad={rad}", arrowstyle="-|>",
                                 mutation_scale=11, color=color, lw=lw, linestyle=estilo, zorder=3,
                                 shrinkA=0, shrinkB=0))


def dibujar(salida):
    fig, ax = plt.subplots(figsize=(10.8, 6.3), dpi=200)
    ax.set_xlim(0, 16)
    ax.set_ylim(1.1, 10.6)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    yD, yQ = 7.0, 2.5
    xs, xi, xn, xo, xk = 0.9, 3.9, 8.0, 12.1, 15.1
    ancho, alto = 3.9, 1.5
    izq, der = xn - ancho / 2, xn + ancho / 2

    # ---------------- Subsistema 1: Diversidad genetica ----------------
    nube(ax, xs, yD, AZUL)
    tuberia(ax, xs + 0.62, izq - 0.36, yD, AZUL)
    nivel(ax, xn, yD, ancho, alto, AZUL, AZUL_F, "Nivel 1: Diversidad D(t)",
          "varianza alélica media  [0, 1]", r"$D_{n+1}=D_n+\Delta t\,(I_D-O_D)$")
    tuberia(ax, der, xk - 0.97, yD, AZUL)
    nube(ax, xk, yD, AZUL)
    valvula(ax, xi, yD, AZUL, r"$I_D$: inyección por mutación", r"$I_D=r_{mut}\,(1-D)$")
    valvula(ax, xo, yD, AZUL, r"$O_D$: selección y deriva", r"$O_D=k_D\,D$")

    auxiliar(ax, xi, 9.35, r"$r_{mut}$", AZUL)
    auxiliar(ax, xo, 9.35, r"$k_D$", AZUL)
    parametro(ax, 1.45, 9.35, r"$P_{mut}$,  $c_m$")
    parametro(ax, 14.75, 9.35, r"$N$,  $k_{sel}$,  $k_e$,  $E$")
    enlace(ax, (2.2, 9.35), (xi - 0.44, 9.35))
    enlace(ax, (13.7, 9.35), (xo + 0.44, 9.35))
    enlace(ax, (xi, 9.35 - 0.43), (xi, yD + 0.88))
    enlace(ax, (xo, 9.35 - 0.43), (xo, yD + 0.88))
    # D -> I_D (polaridad -) y D -> O_D (polaridad +)
    enlace(ax, (izq + 0.7, yD + alto / 2), (xi + 0.16, yD + 0.80), rad=0.35)
    enlace(ax, (der - 0.7, yD + alto / 2), (xo - 0.16, yD + 0.80), rad=-0.35)

    # ---------------- Subsistema 2: Calidad media ----------------
    nube(ax, xs, yQ, VERDE)
    tuberia(ax, xs + 0.62, izq - 0.36, yQ, VERDE)
    nivel(ax, xn, yQ, ancho, alto, VERDE, VERDE_F, "Nivel 2: Calidad Q(t)",
          r"aptitud media $\bar{F}/Z^*_{fact}$  [0, 1]", r"$Q_{n+1}=Q_n+\Delta t\,(I_Q-O_Q)$")
    tuberia(ax, der, xk - 0.97, yQ, VERDE)
    nube(ax, xk, yQ, VERDE)
    valvula(ax, xi, yQ, VERDE, r"$I_Q$: ganancia de aptitud", r"$I_Q=k_Q\,D\,B$")
    valvula(ax, xo, yQ, VERDE, r"$O_Q$: carga mutacional", r"$O_Q=k_{deg}\,\max(0,\,Q-Q_{azar})$")

    auxiliar(ax, 1.75, 4.45, r"$k_Q$", VERDE)
    auxiliar(ax, 7.0, 4.45, r"$B$", VERDE)
    auxiliar(ax, 11.0, 4.45, r"$k_{deg}$", VERDE)
    auxiliar(ax, 14.35, 4.45, r"$Q_{azar}$", VERDE, r=0.52)
    parametro(ax, 1.75, 5.5, r"$k_{base}$, $k_{cr}P_{cr}$, $k_{el}E$, $\Phi_{fact}$")
    parametro(ax, 11.0, 5.5, r"$k_{carga}$,  $P_{mut}$,  $E/N$")
    ax.text(7.55, 4.45, r"$B=1-Q$ (brecha / backlog)", ha="left", va="center", fontsize=8.8, color=VERDE)

    enlace(ax, (1.75, 5.2), (1.75, 4.88))
    enlace(ax, (11.0, 5.2), (11.0, 4.88))
    enlace(ax, (1.75 + 0.30, 4.45 - 0.30), (xi - 0.13, yQ + 0.85))
    enlace(ax, (11.0 + 0.30, 4.45 - 0.30), (xo - 0.13, yQ + 0.85))
    enlace(ax, (14.35 - 0.40, 4.45 - 0.34), (xo + 0.13, yQ + 0.85))
    # Q -> B -> I_Q (bucle B1)
    enlace(ax, (7.6, yQ + alto / 2), (7.1, 4.45 - 0.42))
    enlace(ax, (7.0 - 0.42, 4.45 - 0.06), (xi + 0.16, yQ + 0.78), rad=0.12)
    # Q -> O_Q (bucle B4)
    enlace(ax, (der - 0.7, yQ + alto / 2), (xo - 0.16, yQ + 0.80), rad=-0.35)
    # Acoplamiento entre niveles: D habilita I_Q (Fisher)
    enlace(ax, (6.75, yD - alto / 2), (xi + 0.12, yQ + 0.90), color=AZUL,
           estilo=(0, (5, 2.5)), lw=1.6)
    ax.text(5.45, 5.1, r"acoplamiento $D \rightarrow I_Q$ (Fisher)", ha="right", va="center",
            fontsize=8.8, color=AZUL)

    # Leyenda
    ax.text(8.0, 10.25, "Convenciones:  rectángulo = nivel  ·  válvula = flujo  ·  círculo = variable auxiliar  ·  "
            "nube = fuente o sumidero\nrecuadro = parámetro exógeno  ·  línea punteada = conector de información",
            fontsize=8.6, color=GRIS, va="center", ha="center", linespacing=1.4)

    fig.savefig(salida, dpi=160, bbox_inches="tight", pad_inches=0.08, facecolor="white")
    plt.close(fig)
    return salida


if __name__ == "__main__":
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("informe/figuras/diagrama_forrester.png")
    destino.parent.mkdir(parents=True, exist_ok=True)
    print(dibujar(destino))
