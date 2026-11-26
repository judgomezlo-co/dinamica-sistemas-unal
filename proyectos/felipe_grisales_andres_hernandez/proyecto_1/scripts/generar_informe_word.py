"""
Generador del informe final del proyecto en Word, sobre la plantilla institucional
"Formato para presentacion de trabajos ASI - Word.docx".

Fuentes de contenido: docs/*.md, Fase_1_Conceptualizacion_Causal/, latex/*.tex,
excel/Modelo_AG_Dinamica_de_Sistemas.xlsx y docs/Bitacora_IA_Proyecto_Dinamica.md.

Uso:
    python src/generar_informe_word.py            # genera el .docx
    python src/generar_informe_word.py --pdf      # ademas actualiza el indice con Word y exporta PDF

Requiere: python-docx, lxml, latex2mathml, matplotlib y Microsoft Office (MML2OMML.XSL);
la opcion --pdf requiere pywin32 y Microsoft Word.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from diagrama_forrester import dibujar as dibujar_forrester  # noqa: E402
from informe_docx import Informe  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]
PLANTILLA = RAIZ / "Formato para presentacion de trabajos ASI - Word.docx"
DIR_INFORME = RAIZ / "informe"
SALIDA = DIR_INFORME / "Informe_Final_Proyecto_Dinamica.docx"
FIG1 = RAIZ / "Fase_1_Conceptualizacion_Causal" / "figuras"
FIG_FORRESTER = DIR_INFORME / "figuras" / "diagrama_forrester.png"

TITULO = ("Modelado de un algoritmo genético mediante dinámica de sistemas: "
          "modelo agregado de dos niveles para el problema de la mochila")
TITULO_CORTO = "Modelado de un algoritmo genético mediante dinámica de sistemas"
AUTORES = ["Felipe Grisales López", "Andrés Hernández"]
CURSO = "Informe final del proyecto del curso Dinámica de Sistemas"
ANIO = "2026"

FIGURAS = ["modos", "arquitectura", "cld", "forrester"]
TABLAS = ["sistema", "gen02", "modulos", "fases", "frontera", "variables", "enlaces", "bucles",
          "espacio", "isomorfismo", "alineacion", "hojas", "ag05", "euler", "estabilidad", "rmse",
          "extremos", "sensibilidad", "politicas", "noparam", "anova", "bitacora"]


# =============================================================================
# Secciones preliminares
# =============================================================================
def resumen(inf):
    inf.titulo1("Resumen", numerado=False)
    inf.p(r"Este trabajo representa, mediante dinámica de sistemas, la convergencia de un algoritmo "
          r"genético (AG) que resuelve el problema de la mochila entera acotada trabajado en clase "
          r"(cuatro objetos, cromosomas de 8 bits y población de 8 individuos). El objetivo fue "
          r"identificar y clasificar las variables del AG en niveles, flujos, brecha, entradas, salidas "
          r"y parámetros, y verificar si un modelo agregado de dos niveles —diversidad genética $D(t)$ y "
          r"calidad media normalizada $Q(t)$— basta para reproducir su comportamiento. La metodología, "
          r"cuantitativa, experimental y de simulación, siguió ocho fases: diagrama causal con 17 enlaces "
          r"y cuatro bucles de balance; diagrama de Forrester y ecuaciones alineadas con el curso; réplica "
          r"corregida del AG en Excel y Python; prueba piloto con Euler explícito ($\Delta t=1$ generación, "
          r"estable con $g_D=0.39$); calibración de siete parámetros contra 500 réplicas Monte Carlo; "
          r"validación dimensional, de condiciones extremas y de sensibilidad; y experimentación "
          r"estadística con $\alpha=0.05$. El modelo cumplió RMSE ≤ 0.05 en los ocho escenarios "
          r"evaluados (máximo 0.044), incluidos tres fuera de muestra. En el escenario de clase el "
          r"equilibrio es $D^*=0.67$ y $Q^*=0.75$. Las pruebas de Wilcoxon, Kruskal-Wallis y "
          r"Mann-Whitney ($p<0.001$) y el ANOVA de tres factores muestran que el elitismo (34.2 % de la "
          r"varianza) y la mutación (23.4 %) gobiernan la calidad final, mientras que el cruce no es "
          r"significativo ($p=0.72$). Se confirma que sin mutación la diversidad se agota y la búsqueda "
          r"se estanca, y se recomienda operar con $P_{mut}=0.05$, $P_{cr}=0.89$ y elitismo.")
    inf.p(r"**Palabras clave:** dinámica de sistemas, algoritmo genético, problema de la mochila, "
          r"diversidad genética, simulación de Euler, validación de modelos, análisis estadístico.",
          before=120)


def introduccion(inf):
    inf.titulo1("Introducción", numerado=False)
    inf.p(r"Los algoritmos genéticos son métodos de búsqueda inspirados en la selección natural que "
          r"operan sobre una población de soluciones codificadas mediante operadores de selección, cruce, "
          r"mutación y reemplazo (Holland, 1975; Goldberg, 1989). Aunque cada operador es sencillo de "
          r"implementar paso a paso en una hoja de cálculo o en un script, observar cromosomas individuales "
          r"no explica por qué el algoritmo converge hacia una solución, se estanca prematuramente o "
          r"mantiene un equilibrio entre exploración y explotación. Esa pregunta es de naturaleza "
          r"sistémica: el comportamiento emerge de la realimentación entre la diversidad disponible en la "
          r"población y la calidad de las soluciones encontradas.")
    inf.p(r"El proyecto surge de la integración de dos asignaturas cursadas en el mismo semestre: "
          r"Sistemas Inteligentes Computacionales, en la que se construyó en Excel un AG para el problema "
          r"de la mochila entera acotada (libro *Mochila.xlsx*), y Dinámica de Sistemas, cuyo curso aporta "
          r"el lenguaje de niveles, flujos, brechas, bucles de realimentación, equilibrio, simulación por "
          r"Euler y validación de modelos (Forrester, 1961; Sterman, 2000). La idea de los autores fue "
          r"explicar, con las herramientas de la dinámica de sistemas, cómo un AG llega a una solución por "
          r"convergencia y cómo se relacionan entre sí sus generaciones.")
    inf.p(r"El objetivo general es modelar un AG como un sistema dinámico agregado e interpretar el "
          r"comportamiento de su realimentación y su convergencia. Los objetivos específicos son: "
          r"(i) identificar las variables del AG y clasificarlas en niveles, flujos, backlog, entradas, "
          r"salidas y parámetros; (ii) formular el diagrama causal, el diagrama de Forrester y las "
          r"ecuaciones del sistema; y (iii) replicar el AG por iteración en Excel y en Python para calibrar, "
          r"validar (RMSE ≤ 0.05) y contrastar estadísticamente el efecto de la mutación, el cruce y el "
          r"elitismo.")
    inf.p(r"El alcance se limita al AG de la mochila trabajado en clase: cuatro tipos de objeto, "
          r"cromosomas de 8 bits y población de 8 individuos, con selección por ruleta, cruce de un punto, "
          r"mutación bit a bit y elitismo de un individuo. El modelo de dinámica de sistemas reproduce la "
          r"trayectoria media esperada del AG —el promedio de muchas corridas— y no el ruido de una corrida "
          r"aislada; además, los cromosomas no son un fluido continuo, por lo que la analogía con los "
          r"modelos de tanque del curso tiene límites que se declaran explícitamente. Los siete parámetros "
          r"de tasa se estiman por calibración y su validez está circunscrita a los escenarios evaluados.")
    inf.p(r"La investigación es cuantitativa, experimental y de simulación, y se desarrolló en ocho fases: "
          r"conceptualización causal, formulación en Forrester y ecuaciones, réplica en Excel y Python, "
          r"prueba piloto, calibración, validación, experimentación y análisis. Las herramientas fueron "
          r"Microsoft Excel (libro de 8 hojas con 8 856 fórmulas y 7 gráficas), Python (simulación Monte "
          r"Carlo, calibración y generación del libro) y LaTeX (compendio matemático).")
    inf.p(r"El documento se organiza así: el capítulo 1 presenta los antecedentes (sistema experimental, "
          r"auditoría de los libros de clase, fundamentos del curso y planteamiento del problema); el "
          r"capítulo 2 describe la metodología; los capítulos 3 a 7 desarrollan las fases del proyecto "
          r"—conceptualización causal, formulación del modelo, réplica y simulación, calibración y "
          r"validación, y experimentación estadística—; el capítulo 8 recoge las conclusiones; el anexo A "
          r"contiene la bitácora de uso de inteligencia artificial y al final se lista la bibliografía.")


# =============================================================================
# Capitulo 1. Antecedentes
# =============================================================================
def antecedentes(inf):
    inf.titulo1("Antecedentes")
    inf.p(r"Este capítulo resume la base experimental y conceptual del proyecto: el AG de la mochila "
          r"trabajado en clase, los hallazgos de la auditoría técnica de los libros de Excel de la "
          r"asignatura, los conceptos del curso de Dinámica de Sistemas que se aplican y el planteamiento "
          r"formal del problema.")

    inf.titulo2("Sistema experimental: el algoritmo genético de la mochila")
    inf.p(r"El sistema de referencia es el problema de la mochila entera acotada desarrollado en la hoja "
          r"*Entrega* de *Mochila.xlsx*, a partir del material de clase de Sistemas Inteligentes "
          r"Computacionales (Universidad Nacional de Colombia, 2026b). Se dispone de cuatro tipos de objeto "
          r"con utilidades $u=(4,5,6,3)$ y pesos $w=(7,6,8,2)$. Cada individuo es un cromosoma de 8 bits "
          r"particionado en segmentos de 1, 2, 3 y 2 bits que codifican las cantidades enteras de cada "
          r"objeto; la población tiene 8 individuos y los operadores son selección por ruleta "
          r"proporcional, cruce de un punto con $P_{cr}=0.89$ (cuatro parejas por generación), mutación bit "
          r"a bit con $P_{mut}=0.15$ y elitismo de un individuo ($E=1$). La Tabla {tab:sistema} resume los "
          r"parámetros.")
    inf.tabla("sistema", "Parámetros del sistema experimental de referencia.",
              ["Parámetro", "Valor o definición"],
              [["Tipos de objeto", r"4 objetos, $m\in\{1,2,3,4\}$"],
               [r"Utilidades $u$", "(4, 5, 6, 3)"],
               [r"Pesos $w$", "(7, 6, 8, 2)"],
               [r"Longitud del cromosoma $L$",
                r"8 bits (1 + 2 + 3 + 2): $x_1\in\{0,1\}$, $x_2\in\{0,\dots,3\}$, $x_3\in\{0,\dots,7\}$, "
                r"$x_4\in\{0,\dots,3\}$"],
               ["Espacio de búsqueda",
                r"$2^8=256$ genotipos; máximo en 11111111 → (1, 3, 7, 3), con $W_{max}=87$ y $Z^*=70$"],
               [r"Tamaño de población $N$", "8 individuos"],
               [r"Capacidad $C$",
                "140 en la corrida de clase; 55 como escenario vinculante; 14 como caso extremo "
                "(lectura alternativa de los apuntes manuscritos)"],
               ["Operadores",
                r"Ruleta proporcional; cruce de un punto ($P_{cr}=0.89$); mutación bit a bit "
                r"($P_{mut}=0.15$); elitismo ($E=1$)"],
               ["Paso y horizonte", r"$\Delta t=1$ generación; $T=30$ generaciones"]],
              [2600, 6159], size=18,
              fuente="elaboración propia a partir de *Mochila.xlsx* y de los apuntes del proyecto.")
    inf.p(r"Las tres generaciones registradas en el libro de clase constituyen la evidencia empírica "
          r"inicial (Tabla {tab:gen02}). Entre la generación 0 y la 1 la utilidad media sube de 35.50 a "
          r"46.50 mientras la diversidad cae de 0.9375 a 0.7969; entre la 1 y la 2 la mutación reinyecta "
          r"diversidad (0.8750) a costa de una leve caída de la media (44.25). El nuevo mejor individuo de "
          r"la generación 1 (11111110, $Z=67$, $W=85$) nació en la pareja 4 **sin cruce** "
          r"($u_{cr}=0.991>0.89$), por mutación directa del bit 7 ($u_7=0.019<0.15$), hecho que motivó la "
          r"segunda hipótesis del proyecto.")
    inf.tabla("gen02", "Métricas de las generaciones 0 a 2 registradas en *Mochila.xlsx*.",
              ["Gen.", r"$Z_{max}$", r"$\bar{Z}$", r"$Q_{AG}$", r"$B_{AG}$", r"$D_{AG}$", "Evento observado"],
              [["0", "61", "35.50", "0.5071", "0.4929", "0.9375",
                "Población inicial diversa; mejor individuo 11111100 ($W=81$)"],
               ["1", "67", "46.50", "0.6643", "0.3357", "0.7969",
                "El elitismo retiene $Z=61$; el nuevo mejor ($Z=67$) surge por mutación sin cruce"],
               ["2", "67", "44.25", "0.6321", "0.3679", "0.8750",
                "El elitismo sostiene $Z_{max}=67$; la mutación reinyecta diversidad y reduce la media"]],
              [600, 750, 800, 900, 900, 900, 3909], size=18,
              alinear=["center"] * 6 + ["left"],
              fuente=r"*Mochila.xlsx*, filas 8–15, 113–120 y 217–224; $Q_{AG}=\bar{Z}/70$, "
                     r"$B_{AG}=1-Q_{AG}$ y $D_{AG}$ es la varianza alélica media normalizada.")

    inf.titulo2("Auditoría de los libros de Excel de la asignatura")
    inf.p(r"Antes de modelar se auditaron los libros y scripts del semestre (*Mochila.xlsx*, "
          r"*Mochila-valoresDec.xlsx*, *mochilaAGVacio*, *mochilainicial26.py* y *mochilacorregida26.py*). "
          r"La auditoría confirmó los parámetros reales del sistema ($C=140$, $N=8$, $L=8$) e identificó "
          r"cinco fallas que debían corregirse en el nuevo libro por iteración:")
    inf.vinetas([
        r"**Restricción de capacidad inactiva.** El peso máximo alcanzable con 8 bits es 87, de modo que "
        r"con $C=140$ el 100 % de los 256 cromosomas es factible y la restricción nunca filtra soluciones; "
        r"con $C=14$, en cambio, 7 de los 8 individuos iniciales serían infactibles.",
        r"**Ausencia de penalización en la aptitud.** La columna de factibilidad solo imprime «SI» o «NO», "
        r"pero la ruleta sigue usando la utilidad completa; con una capacidad restrictiva, individuos "
        r"infactibles recibirían las mayores probabilidades de selección. Se adoptó la pena de muerte: "
        r"aptitud efectiva nula si $W>C$.",
        r"**Pasos manuales.** La selección, el cruce y el reemplazo elitista estaban escritos como valores "
        r"fijos o referencias manuales; se automatizaron con fórmulas (búsqueda sobre la probabilidad "
        r"acumulada, corte paramétrico y selección automática del élite).",
        r"**Fórmulas incompatibles y errores de texto.** Una fórmula matricial producía #VALUE! cuando no "
        r"había cruce y un cuadro de texto reportaba «4 objetos 1» en lugar de una unidad del objeto 1.",
        r"**Errores de arrastre en la ruleta.** En los libros auxiliares la probabilidad de selección de "
        r"una generación se dividía entre la suma de la generación anterior, de modo que la probabilidad "
        r"acumulada cerraba en 1.0612 o en 0.9268 en lugar de 1.",
    ])
    inf.p(r"Estas correcciones se incorporaron en la hoja 03_AG_Paso_a_Paso del libro final, que "
          r"reutiliza la población inicial y los números aleatorios de la corrida de clase y extiende la "
          r"simulación hasta la generación 5.")

    inf.titulo2("Fundamentos de dinámica de sistemas aplicados")
    inf.p(r"En dinámica de sistemas el comportamiento se explica por la estructura de acumulaciones "
          r"(niveles) y tasas de cambio (flujos) conectadas por bucles de realimentación (Forrester, 1961; "
          r"Sterman, 2000): un nivel solo cambia por la diferencia entre sus flujos de entrada y de salida, "
          r"$dX/dt=F_e-F_s$. La formulación del proyecto se apoya en los diez módulos del curso "
          r"(Universidad Nacional de Colombia, 2026a), cuyos conceptos se relacionan en la Tabla "
          r"{tab:modulos} con su uso en el proyecto.")
    inf.tabla("modulos", "Módulos del curso de Dinámica de Sistemas y su aplicación en el proyecto.",
              ["Módulo", "Concepto central", "Aplicación en el proyecto"],
              [["1. Pensar en sistemas", "Propósito, frontera y decisión",
                r"Frontera del modelo y pregunta de decisión sobre $P_{mut}$, $P_{cr}$, $E$ y $C$"],
               ["2. Comportamiento dinámico", "Modos de referencia",
                "Regímenes de convergencia, equilibrio y degradación del AG"],
               ["3. Niveles, flujos y parámetros", "Diccionario de variables y unidades",
                "Clasificación de variables y prueba dimensional"],
               ["4. Realimentación y causalidad", "Polaridad de enlaces y bucles",
                "Diagrama causal con 17 enlaces y bucles B1–B4"],
               ["5. Diagramas de niveles y flujos", "Diagrama de Forrester",
                "Modelo de dos niveles acoplados"],
               ["6. Tanque hidráulico", r"Balance y drenaje proporcional $q_s=kh$",
                r"Consumo de diversidad $O_D=k_D D$"],
               ["7. Equilibrio y transitorio", r"$dX/dt=0$ y constante de tiempo $\tau=1/k$",
                r"Equilibrios $D^*$, $Q^*$ y constantes $\tau_D$, $\tau_Q$"],
               ["8. Simulación con Euler", "Regla de un solo instante y estabilidad",
                r"Simulación $n\to n+1$ con $|g_D|<1$"],
               ["9. Modelos isomorfos", "Analogías y sus límites",
                "Isomorfismo tanque – mesa de ayuda TI – AG"],
               ["10. Sistemas informáticos", "Disponibilidad, integral de brecha y decisión",
                r"Factor $\Phi_{fact}$, integral $E_B$ y política recomendada"]],
              [2600, 2900, 3259], size=17,
              fuente="elaboración propia con base en el material del curso.")

    inf.titulo2("Planteamiento del problema e hipótesis")
    inf.p(r"**Problema.** ¿Cómo usar los conceptos de la dinámica de sistemas —niveles, flujos, backlog o "
          r"brecha, entradas, salidas, parámetros y bucles de realimentación— para representar las "
          r"iteraciones de un AG como un sistema dinámico, explicar cómo se relacionan sus generaciones y "
          r"analizar cómo llega a la convergencia o al equilibrio?")
    inf.p(r"**Hipótesis 1 (suficiencia estructural de dos niveles).** La convergencia de un AG se debe a "
          r"una realimentación en la que la selección consume diversidad a mayor ritmo que la mutación la "
          r"repone; sin diversidad la búsqueda se estanca, por lo que un modelo agregado de dos niveles "
          r"—diversidad y calidad— basta para reproducir el comportamiento medio del AG.")
    inf.p(r"**Hipótesis 2 (papel de la mutación intrageneracional).** El mejor individuo no siempre "
          r"proviene de la realimentación heredada de la generación anterior; la mutación aplicada en la "
          r"generación actual puede producir mejoras directas y determinantes.")


# =============================================================================
# Capitulo 2. Metodologia
# =============================================================================
def metodologia(inf):
    inf.titulo1("Metodología")
    inf.p(r"La investigación es cuantitativa, experimental y de simulación. Se siguió el proceso de "
          r"modelado de la dinámica de sistemas en ocho fases encadenadas (Tabla {tab:fases}), en las que "
          r"cada entregable alimenta al siguiente y queda registrado en el repositorio del proyecto.")
    inf.tabla("fases", "Fases metodológicas del proyecto y productos principales.",
              ["Fase", "Nombre", "Producto principal", "Estado"],
              [["1", "Conceptualización en diagrama causal",
                "Modos de referencia, hipótesis dinámicas, frontera, clasificación de variables y "
                "diagrama causal con 17 enlaces", "Completada"],
               ["2", "Formulación en Forrester y ecuaciones",
                "Modelo de dos niveles alineado con el curso, equilibrios y compendio LaTeX", "Completada"],
               ["3", "Réplica del Excel en Python y verificación",
                "Hoja 03_AG_Paso_a_Paso (generaciones 0 a 5) y generador en Python", "Completada"],
               ["4", "Prueba piloto",
                r"Horizonte $T=30$, estabilidad de Euler y sensibilidad a $\Delta t=0.5$", "Completada"],
               ["5", "Calibración",
                "Siete parámetros estimados sobre 5 escenarios con 500 réplicas Monte Carlo", "Completada"],
               ["6", "Validación",
                "Prueba dimensional, 5 condiciones extremas, sensibilidad ±10 % y RMSE en 8 escenarios",
                "Completada"],
               ["7", "Experimentación",
                "Wilcoxon, Kruskal-Wallis, Mann-Whitney y ANOVA de tres factores", "Completada"],
               ["8", "Análisis y conclusiones",
                "Contraste de hipótesis, política recomendada e informe final",
                "En curso (diapositivas de sustentación pendientes)"]],
              [700, 2500, 3859, 1700], size=18, alinear=["center", "left", "left", "left"],
              fuente="README del repositorio del proyecto.")
    inf.p(r"**Herramientas.** El libro *Modelo_AG_Dinamica_de_Sistemas.xlsx* concentra el modelo "
          r"completo en 8 hojas con 8 856 fórmulas y 7 gráficas nativas; el script "
          r"*generar_modelo_excel.py* replica en Python el AG exacto de 8 bits, ejecuta las simulaciones "
          r"Monte Carlo, calibra el modelo y genera el libro; y el documento "
          r"*Formulas_y_Desarrollo_Matematico.tex* desarrolla paso a paso todas las deducciones.")
    inf.p(r"**Criterios de validación.** Se aplicaron las pruebas del curso: consistencia estructural y "
          r"dimensional, comportamiento en condiciones extremas, plausibilidad de parámetros, sensibilidad "
          r"±10 % y estabilidad numérica de Euler. Como criterio cuantitativo, el error cuadrático medio "
          r"entre el modelo y el promedio Monte Carlo del AG debía cumplir RMSE ≤ 0.05 tanto en la calidad "
          r"como en la diversidad.")
    inf.p(r"**Diseño experimental y análisis estadístico.** Con un nivel de significancia $\alpha=0.05$ "
          r"se plantearon cuatro pruebas: Wilcoxon de rangos con signo (muestras pareadas) para el efecto "
          r"de $P_{mut}>0$ sobre la diversidad final; Kruskal-Wallis para el efecto de cuatro niveles de "
          r"$P_{mut}$ sobre la calidad final; U de Mann-Whitney para el efecto del elitismo; y ANOVA de "
          r"tres factores ($P_{mut}\times P_{cr}\times E$) para efectos principales e interacciones "
          r"(Wilcoxon, 1945; Mann y Whitney, 1947; Kruskal y Wallis, 1952; Montgomery, 2017). Las pruebas "
          r"no paramétricas usan 30 réplicas por grupo y el ANOVA un diseño factorial balanceado "
          r"$2\times2\times2$ con 20 réplicas por celda ($N=160$).")


# =============================================================================
# Capitulo 3. Conceptualizacion causal (Fase 1)
# =============================================================================
def conceptualizacion(inf):
    inf.titulo1("Conceptualización causal del algoritmo genético")
    inf.p(r"La primera fase construyó el modelo conceptual que traduce la dinámica discreta del AG en una "
          r"arquitectura agregada de realimentación: modos de referencia, hipótesis dinámicas, frontera "
          r"del modelo, clasificación de variables y diagrama causal.")

    inf.titulo2("Modos de referencia")
    inf.p(r"Antes de formular ecuaciones se identificaron los patrones temporales de las variables clave "
          r"—diversidad $D(t)$, calidad $Q(t)$ y brecha $B(t)$— bajo distintas configuraciones de los "
          r"operadores (Figura {fig:modos}). Se distinguen cuatro regímenes:")
    inf.vinetas([
        r"**Régimen I, convergencia prematura ($P_{mut}=0$, $E=1$).** Sin mutación no entra nueva "
        r"diversidad; la selección, la fijación elitista y la deriva de una población pequeña llevan "
        r"$D(t)\to0$. La calidad sube rápido al inicio, pero la población se vuelve clónica "
        r"($\bar{Q}\approx Q_{max}\approx0.90$) y queda atrapada con una brecha residual cercana a 0.10.",
        r"**Régimen II-opt, equilibrio productivo ($P_{mut}=0.05$, $E=1$).** La mutación moderada compensa "
        r"el consumo por selección y estabiliza la diversidad en $D^*\approx0.44$; la calidad media "
        r"converge a un nivel alto ($\bar{Q}\approx0.82$) y el elitismo permite que el mejor individuo "
        r"alcance el óptimo ($Z^*=70$).",
        r"**Régimen II-base, alta exploración con carga mutacional ($P_{mut}=0.15$, $E=1$).** Es la "
        r"configuración de clase: la diversidad se mantiene alta ($D^*\approx0.67$) y se garantiza hallar "
        r"el óptimo, pero como $(1-0.15)^8\approx0.272$, cerca del 72.8 % de los hijos no élite sufre al "
        r"menos una mutación por generación, lo que sostiene la media en $\bar{Q}\approx0.75$.",
        r"**Régimen III, degradación sin elitismo ($P_{mut}=0.15$, $E=0$).** Sin la protección del élite "
        r"la carga mutacional afecta a toda la población: la diversidad sube ($D^*\approx0.75$), la "
        r"calidad media cae hacia $\bar{Q}\approx0.64$ y puede perderse el mejor individuo ya descubierto.",
    ])
    inf.figura("modos", FIG1 / "modos_referencia_ag.png",
               r"Modos de referencia dinámicos y retrato de fase del algoritmo genético ($t=0$ a 30, $C=140$).",
               "elaboración propia (Fase 1, *generar_graficos_fase1.py*), con los datos de *Mochila.xlsx* "
               "para las generaciones 0 a 2.", ancho_cm=15.2)
    inf.p(r"El retrato de fase (panel D) muestra que, desde el mismo estado inicial observado en clase "
          r"$(D_0,Q_0)=(0.94,0.51)$, el sistema converge hacia tres tipos de atractores: la frontera "
          r"absorbente $D=0$ del régimen I, los atractores interiores productivos con $E=1$ y el atractor "
          r"degradado con $E=0$.")

    inf.titulo2("Hipótesis dinámicas")
    inf.p(r"**H1, suficiencia estructural de dos niveles acoplados.** La trayectoria de convergencia o "
          r"estancamiento emerge de la realimentación entre dos niveles: la diversidad genética $D(t)$ y la "
          r"calidad media normalizada $Q(t)$. La selección y la deriva consumen diversidad a un ritmo "
          r"proporcional a $D(t)$; si la mutación no la repone, $D(t)$ colapsa, bloquea la ganancia de "
          r"aptitud y estanca la búsqueda con una brecha residual $B(t)>0$.")
    inf.p(r"**H2, dualidad funcional de la mutación.** La mutación cumple dos funciones opuestas: una ruta "
          r"exploratoria positiva ($P_{mut}\to I_D\to D\to I_Q\to Q$), que inyecta variación y habilita "
          r"saltos directos a mejores genotipos incluso sin cruce, como ocurrió en la generación 1 de "
          r"clase; y una ruta disruptiva negativa ($P_{mut}\to O_Q\to Q$), que degrada la aptitud media de "
          r"los hijos no protegidos por el elitismo y explica por qué la media puede bajar aunque el mejor "
          r"individuo se mantenga.")

    inf.titulo2("Frontera del modelo y arquitectura micro–macro")
    inf.p(r"Para no replicar el AG bit a bit —lo que no explicaría su comportamiento— sin omitir la "
          r"restricción de capacidad, se definió explícitamente la frontera del sistema (Tabla "
          r"{tab:frontera}) y un puente entre el nivel microscópico (cromosomas, operadores y "
          r"restricciones) y el macroscópico (niveles y flujos), ilustrado en la Figura {fig:arquitectura}.")
    inf.tabla("frontera", "Frontera del modelo (model boundary chart).",
              ["Endógenas", "Exógenas", "Excluidas"],
              [["Nivel 1: diversidad genética $D(t)$\n"
                "Nivel 2: calidad media normalizada $Q(t)$\n"
                "Flujos de $D$: inyección por mutación $I_D$ y consumo por selección y deriva $O_D$\n"
                "Flujos de $Q$: ganancia de aptitud $I_Q$ y carga mutacional $O_Q$\n"
                "Backlog: brecha al óptimo $B(t)=1-Q(t)$",
                r"Operadores: $P_{mut}$, $P_{cr}$, $E\in\{0,1\}$, $N=8$, $L=8$" "\n"
                r"Problema: $u$, $w$ y $C\in\{140,55\}$" "\n"
                r"Paisaje combinatorio: $\Phi_{fact}(C)$, $Q_{azar}(C)$, $Z^*_{fact}(C)$" "\n"
                r"Condiciones iniciales $D(0)$ y $Q(0)$",
                "Identidad y posición de cada cromosoma\n"
                "Números aleatorios específicos por bit o por pareja\n"
                "Epistasis fina entre segmentos 1 + 2 + 3 + 2 (agregada en las tasas calibradas)\n"
                "Operadores no usados en clase (torneo, cruce multipunto, mutación adaptativa)"]],
              [2920, 2920, 2919], size=18,
              fuente="elaboración propia (Fase 1).")
    inf.figura("arquitectura", FIG1 / "arquitectura_subsistemas.png",
               "Arquitectura de acoplamiento micro–macro entre el AG y el modelo de dinámica de sistemas.",
               "elaboración propia (Fase 1). La figura conserva la notación preliminar de la Fase 1; las "
               "ecuaciones definitivas de los flujos se presentan en el capítulo 4.", ancho_cm=15.2)
    inf.p(r"El puente micro–macro se apoya en tres invariantes del espacio de búsqueda calculados sobre "
          r"los 256 genotipos: la tasa de factibilidad $\Phi_{fact}$ (fracción de cromosomas con "
          r"$W\le C$), el óptimo factible $Z^*_{fact}$, que normaliza la calidad, y la calidad esperada al "
          r"azar $Q_{azar}$, nivel hacia el cual la mutación arrastra a la población. Sus valores exactos "
          r"se presentan en la Tabla {tab:espacio} (sección 4.1).")

    inf.titulo2("Clasificación de variables")
    inf.p(r"En respuesta al primer objetivo específico, las variables del AG se clasificaron en las "
          r"categorías de la dinámica de sistemas, con unidades y rango (Tabla {tab:variables}). La "
          r"unidad de tiempo es la generación (gen).")
    inf.tabla("variables", "Clasificación de las variables del AG en categorías de dinámica de sistemas.",
              ["Categoría", "Variable", "Símbolo", "Unidades", "Definición operativa"],
              [["Nivel 1", "Diversidad genética", "$D(t)$", "adim. [0, 1]",
                r"Varianza alélica media de los 8 loci: $\frac{1}{8}\sum_j 4p_j(1-p_j)$"],
               ["Nivel 2", "Calidad media normalizada", "$Q(t)$", "adim. [0, 1]",
                r"Aptitud efectiva media sobre el óptimo factible: $\bar{F}/Z^*_{fact}$"],
               ["Flujo de entrada a $D$", "Inyección por mutación", "$I_D$", r"$\text{gen}^{-1}$",
                r"Restitución de polimorfismo por inversión de bits sobre la fracción fijada $(1-D)$"],
               ["Flujo de salida de $D$", "Consumo por selección y deriva", "$O_D$", r"$\text{gen}^{-1}$",
                r"Homogeneización por ruleta, muestreo finito ($1/N$) y copia élite"],
               ["Flujo de entrada a $Q$", "Ganancia de aptitud", "$I_Q$", r"$\text{gen}^{-1}$",
                r"Mejora por selección, cruce y elitismo, sujeta a $D$, $B$ y $\Phi_{fact}$"],
               ["Flujo de salida de $Q$", "Carga mutacional", "$O_Q$", r"$\text{gen}^{-1}$",
                r"Erosión de la calidad de los hijos no protegidos hacia $Q_{azar}$"],
               ["Backlog o brecha", "Brecha al óptimo", "$B(t)$", "adim. [0, 1]",
                r"$B=1-Q$; en utilidad, $B_Z=Z^*_{fact}-\bar{Z}$"],
               ["Auxiliares", "Utilidad, peso y aptitud efectiva", "$Z_i$, $W_i$, $F_i$", "utilidad, peso",
                r"$Z_i=\sum u_m x_{i,m}$, $W_i=\sum w_m x_{i,m}$, $F_i=Z_i$ si $W_i\le C$ (si no, 0)"],
               ["Parámetros estructurales", "Factibilidad, óptimo y línea base",
                r"$\Phi_{fact}$, $Z^*_{fact}$, $Q_{azar}$", "adim., utilidad",
                r"Propiedades fijas del paisaje determinadas por $(u,w,L,C)$"],
               ["Entradas configurables", "Operadores y mochila",
                r"$P_{mut}$, $P_{cr}$, $E$, $C$, $N$", "adim., peso, ind.",
                "Variables de decisión manipuladas en la experimentación"],
               ["Salidas", "Trayectorias, equilibrio y error", r"$Q(t)$, $D(t)$, $Z_{max}(t)$, RMSE",
                "adim., utilidad", "Métricas de validación y de comparación de políticas"]],
              [1500, 1700, 1250, 1150, 3159], size=17, unida=False,
              fuente="elaboración propia (Fase 1).")
    inf.p(r"El concepto de *backlog*, que en los apuntes iniciales generaba dudas en el contexto de un "
          r"algoritmo evolutivo, corresponde a la brecha al óptimo: la «presión de trabajo pendiente» que "
          r"impulsa la ganancia $I_Q$ y se consume asintóticamente a medida que $Q\to1$.")

    inf.titulo2("Diagrama causal y bucles de realimentación")
    inf.p(r"La Figura {fig:cld} presenta el diagrama causal con los dos niveles, los cuatro flujos, la "
          r"brecha, los bloques de parámetros exógenos y 17 enlaces dirigidos. Cada polaridad se definió "
          r"por el signo de la derivada parcial $\partial Y/\partial X$ con las demás variables "
          r"constantes; por ejemplo, el enlace $Q\to B$ es negativo porque $\partial B/\partial Q=-1$, y el "
          r"enlace $D\to I_Q$ es positivo por el teorema fundamental de la selección natural, según el "
          r"cual la tasa de mejora de la aptitud media es proporcional a la varianza genética disponible "
          r"(Fisher, 1930). La Tabla {tab:enlaces} documenta los 17 enlaces.")
    inf.figura("cld", FIG1 / "diagrama_causal_cld.png",
               "Diagrama causal (CLD) del modelo agregado de dos niveles del algoritmo genético.",
               "elaboración propia (Fase 1, *generar_graficos_fase1.py*).", ancho_cm=15.2)
    inf.tabla("enlaces", "Matriz de enlaces causales y justificación de polaridades.",
              ["ID", "Enlace", "Signo", "Mecanismo en el AG"],
              [["E1", r"$Q\to B$", "−", r"$B=1-Q$, por lo que $\partial B/\partial Q=-1$"],
               ["E2", r"$B\to I_Q$", "+", "Más brecha implica más alelos subóptimos que la selección y el "
                                          "cruce pueden sustituir"],
               ["E3", r"$I_Q\to Q$", "+", "Flujo de entrada del nivel 2"],
               ["E4", r"$D\to I_Q$", "+", r"Teorema de Fisher: sin diversidad ($D=0$) no hay mejora"],
               ["E5", r"$D\to O_D$", "+", "La pérdida de diversidad es proporcional a la diversidad existente"],
               ["E6", r"$O_D\to D$", "−", "Flujo de salida del nivel 1"],
               ["E7", r"$D\to I_D$", "−", r"La mutación solo crea variación donde hay homogeneidad $(1-D)$"],
               ["E8", r"$I_D\to D$", "+", "Flujo de entrada del nivel 1"],
               ["E9", r"$Q\to O_Q$", "+", r"Cuanto más arriba de $Q_{azar}$, más probable que una mutación "
                                          r"sea dañina"],
               ["E10", r"$O_Q\to Q$", "−", "Flujo de salida del nivel 2"],
               ["E11", r"$P_{mut}\to I_D$", "+", "Ruta exploratoria de H2"],
               ["E12", r"$P_{mut}\to O_Q$", "+", "Ruta disruptiva de H2"],
               ["E13", r"$P_{cr}\to I_Q$", "+", "Más parejas recombinan esquemas complementarios"],
               ["E14", r"$E\to I_Q$", "+", r"El élite asegura un piso creciente para $Z_{max}$"],
               ["E15", r"$E\to O_D$", "+", "La copia élite acelera la homogeneización"],
               ["E16", r"$E\to O_Q$", "−", r"El élite no se muta: protege la fracción $1/N$"],
               ["E17", r"$\Phi_{fact}\to I_Q$", "+", "Más hijos factibles y menos rechazos por pena de muerte"]],
              [700, 1700, 800, 5559], size=18, alinear=["center", "center", "center", "left"],
              unida=False, fuente="elaboración propia (Fase 1).")
    inf.p(r"Multiplicando las polaridades a lo largo de cada ciclo cerrado se verifica que los cuatro "
          r"bucles son de balance (Tabla {tab:bucles}), aunque desempeñan papeles evolutivos opuestos.")
    inf.tabla("bucles", "Bucles de realimentación del modelo.",
              ["Bucle", "Ruta causal", "Signo neto", "Función dinámica"],
              [["B1", r"$Q\to B\to I_Q\to Q$", "(−)(+)(+) = (−)",
                r"Búsqueda asintótica: motor de explotación que empuja $Q$ hacia 1 y se desacelera "
                r"cuando $B\to0$"],
               ["B2", r"$D\to O_D\to D$", "(+)(−) = (−)",
                r"Agotamiento: la ruleta, la deriva ($N=8$) y la fijación elitista vacían $D$ "
                r"exponencialmente"],
               ["B3", r"$D\to I_D\to D$", "(−)(+) = (−)",
                r"Reposición mutacional: con $P_{mut}>0$ compite con B2 y sostiene $D^*>0$"],
               ["B4", r"$Q\to O_Q\to Q$", "(+)(−) = (−)",
                r"Freno por carga mutacional: limita la calidad estacionaria por debajo de 1 según la "
                r"fracción no protegida $(1-E/N)$"]],
              [800, 1900, 1500, 4559], size=18, alinear=["center", "center", "center", "left"],
              fuente="elaboración propia (Fase 1).")
    inf.p(r"El comportamiento complejo del AG no se debe a la aparición de bucles nuevos, sino al "
          r"**cambio de dominancia** entre B1 y B4 mediado por el enlace multiplicativo $D\to I_Q$. Al "
          r"inicio, con alta diversidad y brecha grande, el producto $D\cdot B$ es alto y domina B1, lo que "
          r"produce el salto de calidad de la generación 0 a la 1. Después, si $P_{mut}=0$, B2 domina sin "
          r"oposición, lleva $D\to0$ y «apaga» a B1 antes de cerrar la brecha (convergencia prematura); si "
          r"$P_{mut}>0$ y $E=1$, B3 equilibra a B2 y mantiene la diversidad que B1 necesita para seguir "
          r"explorando, mientras B4 fija el nivel estacionario de la calidad media; y si la mutación es "
          r"alta y $E=0$, B4 domina sobre B1 y la calidad media cae hacia $Q_{azar}$.")


# =============================================================================
# Capitulo 4. Formulacion (Fase 2)
# =============================================================================
def formulacion(inf):
    inf.titulo1("Formulación del modelo de dos niveles")
    inf.p(r"La segunda fase convirtió la estructura causal en un diagrama de Forrester y en ecuaciones "
          r"dimensionalmente consistentes. Primero se formaliza el nivel microscópico del AG, que genera "
          r"los datos de referencia; luego se deducen las ecuaciones agregadas, sus equilibrios y su "
          r"solución analítica.")

    inf.titulo2("Modelo microscópico del algoritmo genético")
    inf.p(r"Cada individuo $i$ es un vector binario $\mathbf{b}_i\in\{0,1\}^8$ que se decodifica por "
          r"conversión binaria posicional en cuatro cantidades enteras:", after=60)
    inf.ecuacion(r"x_{i,1}=b_{i,1},\quad x_{i,2}=2b_{i,2}+b_{i,3},\quad x_{i,3}=4b_{i,4}+2b_{i,5}+b_{i,6},"
                 r"\quad x_{i,4}=2b_{i,7}+b_{i,8}")
    inf.p(r"La utilidad, el peso y la aptitud efectiva con pena de muerte son:", after=60)
    inf.ecuacion(r"Z_i=\sum_{m=1}^{4}u_m x_{i,m},\qquad W_i=\sum_{m=1}^{4}w_m x_{i,m},\qquad "
                 r"F_i=\begin{cases}Z_i & \text{si } W_i\le C\\ 0 & \text{si } W_i>C\end{cases}")
    inf.p(r"Por ejemplo, el individuo 1 de la generación 0, $\mathbf{b}_1=(0,1,0,0,1,0,1,1)$, se "
          r"decodifica como $\mathbf{x}_1=(0,2,2,3)$, con $Z_1=31$ y $W_1=34$: es factible con $C=140$ o "
          r"$C=55$, pero no con $C=14$. Los operadores se formulan así: la ruleta asigna "
          r"$p_i^{sel}=F_i/\sum_k F_k$ (con respaldo uniforme $1/N$ si toda la población es infactible) y "
          r"elige el primer individuo cuya probabilidad acumulada cubre un aleatorio $r\sim U(0,1)$; el "
          r"cruce de un punto, aplicado si $u_{cr}<P_{cr}$, usa el corte $c=\min(7,\lfloor 7u_{corte}"
          r"\rfloor+1)$; la mutación invierte cada bit si $u_j<P_{mut}$; y con elitismo la fila 1 de la "
          r"nueva generación se reserva para el mejor individuo anterior. Al cierre de cada generación el "
          r"estado microscópico se agrega en las métricas puente hacia el modelo macroscópico:", after=60)
    inf.ecuacion(r"\bar{Z}(t)=\frac{1}{N}\sum_{i=1}^{N}F_i(t),\qquad Q_{AG}(t)=\frac{\bar{Z}(t)}{Z^*_{fact}},"
                 r"\qquad D_{AG}(t)=\frac{1}{L}\sum_{j=1}^{L}4\,p_j(t)\left(1-p_j(t)\right)")
    inf.p(r"donde $p_j$ es la proporción de unos en la columna $j$. El factor 4 normaliza la varianza "
          r"alélica al intervalo [0, 1], pues $p(1-p)$ alcanza su máximo 0.25 en $p=0.5$. En la "
          r"generación 0 de clase, $\sum F_i=284$, de modo que $\bar{Z}(0)=35.50$, $Q_{AG}(0)=0.5071$ y "
          r"$D_{AG}(0)=7.5/8=0.9375$; para una población inicial aleatoria el valor esperado es "
          r"$1-1/N=0.875$, que se usa como condición inicial $D_0$ en las simulaciones Monte Carlo.")
    inf.p(r"La enumeración exhaustiva de los 256 genotipos permite calcular exactamente los invariantes "
          r"del espacio de búsqueda para cada capacidad (Tabla {tab:espacio}). Con $C=140$ la restricción "
          r"es inactiva; con $C=55$ el 31.25 % de los genotipos es rechazado por la pena de muerte.")
    inf.tabla("espacio", "Propiedades exactas del espacio de búsqueda de 256 estados según la capacidad.",
              ["Escenario de capacidad", "Estados factibles", r"$\Phi_{fact}$", r"$Z^*_{fact}$",
               r"$Q_{azar}$"],
              [[r"Clase, restricción inactiva ($C=140$)", "256 / 256", "1.0000", "70", "0.5000"],
               [r"Capacidad vinculante ($C=55$)", "176 / 256", "0.6875", "46", "0.4018"],
               [r"Caso extremo de los apuntes ($C=14$)", "20 / 256", "0.0781", "15", "0.0456"]],
              [3359, 1500, 1300, 1300, 1300], size=18,
              alinear=["left", "center", "center", "center", "center"],
              fuente=r"elaboración propia por enumeración exhaustiva de los $2^8$ cromosomas (hoja "
                     r"02_Parametros_y_Espacio).")

    inf.titulo2("Diagrama de Forrester")
    inf.p(r"La Figura {fig:forrester} muestra el diagrama de Forrester: dos niveles acoplados, cada uno "
          r"con un flujo de entrada y uno de salida controlados por válvulas, variables auxiliares que "
          r"agrupan los parámetros en tasas efectivas y conectores de información que cierran los bucles "
          r"B1 a B4. El conector que va de la diversidad a la válvula de ganancia de aptitud acopla los dos "
          r"subsistemas: si $D=0$ la válvula $I_Q$ se cierra aunque persista brecha.")
    inf.figura("forrester", FIG_FORRESTER, "Diagrama de Forrester del modelo agregado de dos niveles.",
               "elaboración propia a partir de las ecuaciones del modelo (*diagrama_forrester.py*).",
               ancho_cm=14.5)

    inf.titulo2("Nivel 1: diversidad genética")
    inf.p(r"El balance de diversidad sigue la plantilla de acumulación del curso, con un flujo de entrada "
          r"de ajuste hacia la meta $D_{max}=1$ y un flujo de salida de drenaje proporcional, análogo a la "
          r"salida $q_s=kh$ del tanque hidráulico:", after=60)
    inf.ecuacion(r"\frac{dD}{dt}=I_D-O_D,\qquad I_D=r_{mut}\left(1-D\right),\quad r_{mut}=c_m P_{mut},"
                 r"\qquad O_D=k_D D,\quad k_D=\frac{1}{N}+k_{sel}+k_e\frac{E}{N}")
    inf.p(r"Los tres sumandos de $k_D$ representan mecanismos aditivos en $\text{gen}^{-1}$: la deriva por "
          r"muestreo finito ($1/N$), la presión de selección de la ruleta ($k_{sel}$) y la fijación "
          r"adicional por la copia élite ($k_e E/N$). Al sustituir se obtiene una ecuación lineal de primer "
          r"orden; imponiendo el equilibrio $dD/dt=0$ e integrando por separación de variables sobre la "
          r"desviación $e_D=D-D^*$ resultan el equilibrio, la constante de tiempo y la solución "
          r"analítica:", after=60)
    inf.ecuacion(r"D^*=\frac{r_{mut}}{r_{mut}+k_D},\qquad \tau_D=\frac{1}{r_{mut}+k_D},\qquad "
                 r"D(t)=D^*+\left(D_0-D^*\right)e^{-t/\tau_D}")
    inf.p(r"Si $P_{mut}=0$, entonces $D^*=0$: la diversidad se agota por completo. En el escenario base "
          r"($P_{mut}=0.15$, $E=1$, $N=8$), con los parámetros calibrados $c_m=2.7209$, $k_{sel}=0.0250$ "
          r"y $k_e=0.3971$, se obtiene $r_{mut}=0.4081$ $\text{gen}^{-1}$, $k_D=0.1996$ $\text{gen}^{-1}$, "
          r"$D^*=0.6715$ y $\tau_D=1.6454$ generaciones, de modo que el 95 % del transitorio se cierra en "
          r"unas $3\tau_D\approx4.94$ generaciones.")

    inf.titulo2("Nivel 2: calidad media y brecha")
    inf.p(r"La calidad media se rige por un flujo de ganancia que cierra la brecha al óptimo y que solo "
          r"opera si hay diversidad, y por un flujo de carga mutacional isomorfo a la ley de enfriamiento "
          r"hacia un nivel ambiente, que aquí es la calidad al azar:", after=60)
    inf.ecuacion(r"\frac{dQ}{dt}=I_Q-O_Q,\qquad I_Q=k_Q\,D\,B=k_Q\,D\left(1-Q\right),\qquad "
                 r"O_Q=k_{deg}\max\left(0,\,Q-Q_{azar}\right)", after=60)
    inf.ecuacion(r"k_Q=\left(k_{base}+k_{cr}P_{cr}+k_{el}E\right)\Phi_{fact},\qquad "
                 r"k_{deg}=k_{carga}P_{mut}\left(1-\frac{E}{N}\right)")
    inf.p(r"El factor $\Phi_{fact}$ actúa como la disponibilidad lineal del módulo 10, el factor "
          r"$(1-E/N)$ expresa que el elitismo protege a $E$ de los $N$ individuos del deterioro y el "
          r"operador $\max(0,\cdot)$ garantiza que la mutación solo degrada cuando $Q>Q_{azar}$. Con "
          r"$P_{mut}>0$, imponiendo $dQ/dt=0$ en el equilibrio $D=D^*$:", after=60)
    inf.ecuacion(r"Q^*=\frac{k_Q D^*+k_{deg}Q_{azar}}{k_Q D^*+k_{deg}},\qquad "
                 r"\tau_Q=\frac{1}{k_Q D^*+k_{deg}}")
    inf.p(r"En el escenario base ($k_{base}=0.1018$, $k_{cr}=0.0218$, $k_{el}=0.1814$, "
          r"$k_{carga}=1.5854$, $\Phi_{fact}=1$, $Q_{azar}=0.5$) resultan $k_Q=0.3026$ $\text{gen}^{-1}$, "
          r"$k_{deg}=0.2081$ $\text{gen}^{-1}$, $Q^*=0.7470$ —una utilidad media de 52.29 sobre 70— y "
          r"$\tau_Q=2.4314$ generaciones.")
    inf.p(r"**Demostración analítica de la hipótesis 1.** Si $P_{mut}=0$, entonces $r_{mut}=0$ y "
          r"$k_{deg}=0$, de modo que $D(t)=D_0e^{-k_D t}$ y $O_Q=0$. Al integrar la ecuación de $Q$ por "
          r"separación de variables y tomar el límite $t\to\infty$:", after=60)
    inf.ecuacion(r"Q(t)=1-\left(1-Q_0\right)\exp\left[-\frac{k_Q D_0}{k_D}\left(1-e^{-k_D t}\right)\right]"
                 r"\;\Rightarrow\;Q_\infty=1-\left(1-Q_0\right)e^{-k_Q D_0/k_D}<1")
    inf.p(r"Sin mutación, la diversidad se agota exponencialmente antes de cerrar la brecha y el algoritmo "
          r"se estanca en un valor subóptimo $Q_\infty<1$ que depende de la diversidad inicial $D_0$. De "
          r"forma equivalente, al eliminar el tiempo la trayectoria en el plano de fase es "
          r"$Q(D)=1-(1-Q_0)\exp[-(k_Q/k_D)(D_0-D)]$ (documento *Ecuacion_de_Trayectoria.tex*), que al "
          r"agotarse la diversidad ($D\to0$) conduce al mismo valor $Q_\infty$.")

    inf.titulo2("Consistencia dimensional e isomorfismo")
    inf.p(r"Las unidades operan como una prueba lógica obligatoria: $D$ y $Q$ son adimensionales, los "
          r"cuatro flujos se expresan en $\text{gen}^{-1}$ y el paso $\Delta t$ en generaciones, de modo "
          r"que $[D_{n+1}]=[D_n]+[\Delta t][I_D-O_D]=1+\text{gen}\cdot\text{gen}^{-1}=1$; la utilidad media "
          r"$\bar{Z}=Z^*_{fact}Q$ conserva unidades de utilidad y las constantes de tiempo $\tau=1/k$ "
          r"quedan en generaciones. El modelo es además isomorfo a los sistemas canónicos del curso (Tabla "
          r"{tab:isomorfismo}), con la salvedad de que los cromosomas no son un fluido continuo: el modelo "
          r"reproduce la trayectoria media esperada del AG y no el ruido de una corrida aislada.")
    inf.tabla("isomorfismo", "Correspondencia isomorfa entre el tanque, la mesa de ayuda TI y el AG.",
              ["Papel en DS", "Tanque (módulo 6)", "Mesa de ayuda TI (módulo 10)", "AG, nivel 1",
               "AG, nivel 2"],
              [["Nivel", "Volumen $V$ o altura $h$", "Backlog de tickets", "Diversidad $D(t)$",
                r"Calidad $Q(t)$ o $\bar{Z}(t)$"],
               ["Flujo de entrada", "Caudal $q_e$", r"Llegadas $\lambda$", "Inyección $I_D$",
                "Ganancia $I_Q$"],
               ["Flujo de salida", "Drenaje $q_s=kh$", r"Resoluciones $\mu$", "Consumo $O_D=k_D D$",
                "Carga $O_Q$"],
               ["Brecha o meta", "$h^*-h$", r"$B-B_{deseado}$", "$1-D$", "$B=1-Q$"],
               ["Disponibilidad", "Área $A$", "Factor $a(t)$ en $C=Npa$", "Tasa $c_m P_{mut}$",
                r"Factibilidad $\Phi_{fact}$"]],
              [1500, 1700, 1959, 1800, 1800], size=18,
              fuente="*Formulas_y_Desarrollo_Matematico.tex*, sección 3.")

    inf.titulo2("Alineación con el curso: evolución desde la propuesta inicial")
    inf.p(r"La propuesta inicial del proyecto, calibrada de forma preliminar con 300 réplicas en seis "
          r"escenarios, contenía constantes y exponentes sin interpretación directa en el curso. En la "
          r"versión final se sustituyeron por estructuras lineales y canónicas de los módulos 3 a 10 "
          r"(Tabla {tab:alineacion}), lo que redujo la sobreparametrización y permitió deducir "
          r"analíticamente equilibrios y constantes de tiempo.")
    inf.tabla("alineacion", "Evolución de la formulación: propuesta inicial frente a modelo final.",
              ["Aspecto", "Formulación inicial", "Formulación final", "Fundamento en el curso"],
              [[r"Flujo $I_D$", r"$4P_{mut}(1-P_{mut})(1-D)$",
                r"$r_{mut}(1-D)$, con $r_{mut}=c_m P_{mut}$", "Ajuste hacia meta (módulo 4, problema 4)"],
               [r"Flujo $I_Q$", r"$(k_q(1+0.25P_{cr})+k_{eq}E)D(1-Q)\Phi^{1.6558}$",
                r"$(k_{base}+k_{cr}P_{cr}+k_{el}E)\Phi_{fact}DB$",
                r"Disponibilidad lineal $C=Npa$ (módulo 10)"],
               [r"Flujo $O_Q$", r"$k_{load}(1-(1-P_{mut})^2)$" "\n" r"$\times(1-E/N)(Q-Q_{azar})$",
                r"$k_{carga}P_{mut}(1-E/N)$" "\n" r"$\times\max(0,Q-Q_{azar})$",
                "Enfriamiento hacia nivel ambiente (problema 5)"],
               ["Integración", r"Notación $t-\Delta t$ sin análisis de estabilidad",
                r"Euler $n\to n+1$ con prueba $|g_D|<1$ y $\Delta t=0.5$", "Módulo 8, problema 28"],
               ["Evaluación de políticas", "Solo RMSE en seis escenarios",
                "Equilibrios, constantes de tiempo, integral de brecha, extremos y sensibilidad ±10 %",
                "Módulos 7, 9 y 10, problema 29"]],
              [1150, 3050, 2850, 1709], size=16,
              fuente="hoja 00_Resumen_Ejecutivo y *Propuesta_Dinamica_de_Sistemas.md*.")


# =============================================================================
# Capitulo 5. Replica, simulacion y prueba piloto (Fases 3 y 4)
# =============================================================================
def simulacion(inf):
    inf.titulo1("Réplica, simulación y prueba piloto")
    inf.p(r"Las fases 3 y 4 implementaron el AG corregido y el modelo agregado en un libro de Excel "
          r"reproducible, verificaron la réplica en Python y fijaron, mediante una prueba piloto, el "
          r"horizonte de simulación y el paso de integración.")

    inf.titulo2("Libro de Excel y réplica en Python")
    inf.p(r"El modelo completo se implementó en el libro *Modelo_AG_Dinamica_de_Sistemas.xlsx*, generado "
          r"de forma reproducible por *generar_modelo_excel.py* y auditado visualmente (14 964 celdas y 7 "
          r"gráficas en Times New Roman, sin errores de fórmula). Sus ocho hojas siguen la secuencia de las "
          r"fases (Tabla {tab:hojas}). La réplica en Python implementa el mismo AG —decodificación, pena de "
          r"muerte, ruleta, cruce, mutación y elitismo— con la población inicial y los números aleatorios "
          r"exactos de la corrida de clase, lo que permite verificar la hoja del AG y generar las 500 "
          r"réplicas Monte Carlo por escenario.")
    inf.tabla("hojas", "Estructura del libro *Modelo_AG_Dinamica_de_Sistemas.xlsx*.",
              ["Hoja", "Fases", "Contenido principal"],
              [["00_Resumen_Ejecutivo", "Síntesis", "Indicadores clave enlazados y mapa de navegación"],
               ["01_Frontera_y_Modelo_DS", "1 y 2",
                "Frontera, diccionario dimensional, bucles B1–B4, Forrester e isomorfismo"],
               ["02_Parametros_y_Espacio", "2 y 3",
                "Panel de control, enumeración de los 256 cromosomas, equilibrios y estabilidad"],
               ["03_AG_Paso_a_Paso", "3", "Réplica corregida de *Mochila.xlsx* (generaciones 0 a 5)"],
               ["04_Simulacion_Euler_DS", "2 y 4",
                r"Euler explícito ($\Delta t=1$ y 0.5), solución analítica y estabilidad numérica"],
               ["05_Calibracion_Validacion", "5 y 6",
                "Calibración (5 escenarios) y validación fuera de muestra (3 escenarios)"],
               ["06_Escenarios_y_Sensibilidad", "6 y 7",
                "Políticas, integral de brecha, condiciones extremas y sensibilidad ±10 %"],
               ["07_Analisis_Estadistico", "7 y 8",
                "Pruebas no paramétricas, ANOVA de tres factores y contraste de hipótesis"]],
              [3000, 1000, 4759], size=18, alinear=["left", "center", "left"],
              fuente="hoja 00_Resumen_Ejecutivo del libro.")

    inf.titulo2("Algoritmo genético paso a paso")
    inf.p(r"La hoja 03_AG_Paso_a_Paso automatiza cada generación con fórmulas: población, fenotipo, "
          r"evaluación con pena de muerte, ruleta, cruce bit a bit, mutación, reemplazo elitista y métricas "
          r"hacia la dinámica de sistemas. La Tabla {tab:ag05} resume las seis generaciones simuladas. La "
          r"transición de la generación 0 a la 1 coincide con el registro de clase y reproduce el evento "
          r"de la hipótesis 2; en la generación 2 aparecen pequeñas diferencias con el libro original "
          r"($\bar{Z}=43.875$ frente a 44.25), atribuibles a que la selección, el cruce y el reemplazo "
          r"ahora se calculan por fórmula en lugar de transcribirse manualmente. Las generaciones 3 a 5 "
          r"continúan la corrida con semillas reproducibles.")
    inf.tabla("ag05", "Métricas del AG corregido, generaciones 0 a 5 ($C=140$).",
              ["Gen.", r"$Z_{max}$", r"$\bar{Z}$", r"$Q_{AG}$", r"$B_{AG}$", r"$\bar{W}$", r"$D_{AG}$",
               "Observación"],
              [["0", "61", "35.500", "0.5071", "0.4929", "44.000", "0.9375", "Población inicial de clase"],
               ["1", "67", "46.500", "0.6643", "0.3357", "59.125", "0.7969",
                "$Z=67$ nace en la pareja 4 sin cruce"],
               ["2", "67", "43.875", "0.6268", "0.3732", "55.250", "0.8672",
                "El elitismo sostiene $Z=67$"],
               ["3", "67", "40.250", "0.5750", "0.4250", "51.875", "0.8672", "Semilla 103"],
               ["4", "67", "48.375", "0.6911", "0.3089", "62.375", "0.8047", "Semilla 104"],
               ["5", "67", "54.500", "0.7786", "0.2214", "71.000", "0.5234", "Semilla 105"]],
              [600, 750, 900, 900, 900, 900, 900, 2909], size=18,
              alinear=["center"] * 7 + ["left"],
              fuente="hoja 03_AG_Paso_a_Paso. La factibilidad fue del 100 % en todas las generaciones.")
    inf.p(r"Una corrida aislada es ruidosa: la utilidad media baja en las generaciones 2 y 3 y sube en "
          r"la 4 y la 5, aunque el mejor individuo se mantiene en 67 gracias al elitismo. Por ello la "
          r"validación del modelo agregado se hace contra el promedio de 500 réplicas y no contra una "
          r"corrida individual.")

    inf.titulo2("Simulación por Euler explícito")
    inf.p(r"Siguiendo la regla de un solo instante del módulo 8, en cada paso $n$ los flujos se evalúan "
          r"exclusivamente con el estado $n$ y solo después se actualizan simultáneamente los niveles, "
          r"acotados físicamente al intervalo [0, 1]:", after=60)
    inf.ecuacion(r"D_{n+1}=\max\left(0,\min\left(1,\,D_n+\Delta t\left(I_{D,n}-O_{D,n}\right)\right)\right)",
                 after=60)
    inf.ecuacion(r"Q_{n+1}=\max\left(0,\min\left(1,\,Q_n+\Delta t\left(I_{Q,n}-O_{Q,n}\right)\right)\right)")
    inf.p(r"La Tabla {tab:euler} desarrolla los tres primeros pasos del escenario base con $\Delta t=1$, "
          r"partiendo de $D_0=0.875$ y $Q_0=0.5$; los valores coinciden al cuarto decimal con las filas 10 "
          r"a 13 de la hoja 04_Simulacion_Euler_DS. La diversidad desciende y la calidad asciende "
          r"rápidamente en las primeras generaciones, y ambas se aproximan a sus equilibrios $D^*=0.6715$ "
          r"y $Q^*=0.7470$.")
    inf.tabla("euler", r"Desarrollo de Euler en el escenario base, pasos $n=0$ a 2 ($\Delta t=1$).",
              ["$n$", "$D_n$", "$Q_n$", "$B_n$", "$I_{D,n}$", "$O_{D,n}$", "$I_{Q,n}$", "$O_{Q,n}$",
               "$D_{n+1}$", "$Q_{n+1}$"],
              [["0", "0.8750", "0.5000", "0.5000", "0.0510", "0.1747", "0.1324", "0.0000", "0.7513",
                "0.6324"],
               ["1", "0.7513", "0.6324", "0.3676", "0.1015", "0.1500", "0.0836", "0.0275", "0.7028",
                "0.6884"],
               ["2", "0.7028", "0.6884", "0.3116", "0.1213", "0.1403", "0.0663", "0.0392", "0.6838",
                "0.7155"]],
              [479] + [920] * 9, size=18, alinear=["center"] * 10,
              fuente=r"hoja 04_Simulacion_Euler_DS ($r_{mut}=0.4081$, $k_D=0.1996$, $k_Q=0.3026$, "
                     r"$k_{deg}=0.2081$, $Q_{azar}=0.5$).")

    inf.titulo2("Prueba piloto: horizonte y estabilidad numérica")
    inf.p(r"La prueba piloto fijó un horizonte de $T=30$ generaciones, que cubre holgadamente el "
          r"transitorio ($T>3\tau$) hasta el equilibrio, y verificó la estabilidad numérica de Euler. Al "
          r"escribir la recurrencia del error $e_{D,n}=D_n-D^*$ se obtiene:", after=60)
    inf.ecuacion(r"e_{D,n+1}=g_D\,e_{D,n},\qquad g_D=1-\left(r_{mut}+k_D\right)\Delta t")
    inf.p(r"La simulación es monótona estable si $0\le g_D<1$ ($\Delta t\le\tau_D$), oscilatoria estable "
          r"si $-1<g_D<0$ y oscilatoria inestable si $|g_D|\ge1$ ($\Delta t\ge2\tau_D$). La Tabla "
          r"{tab:estabilidad} muestra que el paso base de una generación es monótono estable "
          r"($g_D=0.3922$), al igual que el paso fino de media generación ($g_D=0.6961$), y que solo pasos "
          r"de dos o más generaciones degradan la simulación.")
    inf.tabla("estabilidad", "Estabilidad numérica de Euler en el escenario base.",
              [r"$\Delta t$ [gen]", r"$k_{tot}\Delta t$", "$g_D$", r"$|g_D|$", "Clasificación",
               "Recomendación"],
              [["0.50", "0.3039", "+0.6961", "0.6961", "Monótona estable", "Adecuado"],
               ["1.00", "0.6078", "+0.3922", "0.3922", "Monótona estable", "Adecuado"],
               ["2.00", "1.2155", "−0.2155", "0.2155", "Oscilatoria estable", r"Reducir $\Delta t$"],
               ["3.50", "2.1272", "−1.1272", "1.1272", "Oscilatoria inestable", r"Rechazar $\Delta t$"]],
              [1100, 1200, 1100, 1000, 2259, 2100], size=18, alinear=["center"] * 6,
              fuente=r"hoja 04_Simulacion_Euler_DS ($k_{tot}=r_{mut}+k_D=0.6078$ $\text{gen}^{-1}$, "
                     r"$\tau_D=1.6454$ gen).")


# =============================================================================
# Capitulo 6. Calibracion y validacion (Fases 5 y 6)
# =============================================================================
def validacion(inf):
    inf.titulo1("Calibración y validación")
    inf.p(r"Las fases 5 y 6 estimaron los parámetros de tasa del modelo y evaluaron su validez "
          r"estructural y de comportamiento con las pruebas del curso.")

    inf.titulo2("Calibración")
    inf.p(r"Los siete parámetros de tasa $(c_m,k_{sel},k_e,k_{base},k_{cr},k_{el},k_{carga})$ se "
          r"estimaron conjuntamente —un único vector para todos los escenarios— contra el promedio de 500 "
          r"réplicas Monte Carlo del AG exacto de 8 bits en cinco escenarios de entrenamiento (E1 a E5). "
          r"La función objetivo fue la suma de los cuadrados de $\text{RMSE}_Q$ y $\text{RMSE}_D$ de cada "
          r"escenario más una penalización del error máximo, minimizada por búsqueda de patrones "
          r"coordenada a coordenada con paso decreciente y restricción de no negatividad. Los errores se "
          r"miden como:", after=60)
    inf.ecuacion(r"\text{RMSE}_Q=\sqrt{\frac{1}{T+1}\sum_{t=0}^{T}\left(Q_{DS}(t)-\bar{Q}_{AG}(t)\right)^2},"
                 r"\qquad \text{RMSE}_D=\sqrt{\frac{1}{T+1}\sum_{t=0}^{T}\left(D_{DS}(t)-\bar{D}_{AG}(t)"
                 r"\right)^2}")
    inf.p(r"El vector calibrado es $c_m=2.7209$, $k_{sel}=0.0250$, $k_e=0.3971$, $k_{base}=0.1018$, "
          r"$k_{cr}=0.0218$, $k_{el}=0.1814$ y $k_{carga}=1.5854$. Todos son positivos y de magnitud "
          r"plausible; el bajo valor de $k_{cr}$ anticipa el escaso efecto del cruce que luego confirmó "
          r"el ANOVA.")

    inf.titulo2("Validación del comportamiento (RMSE)")
    inf.p(r"El modelo calibrado se evaluó fuera de muestra en tres escenarios independientes (E6 a E8) "
          r"que no participaron en la calibración. Los ocho escenarios cumplen el criterio RMSE ≤ 0.05 en "
          r"calidad y en diversidad (Tabla {tab:rmse}); el error máximo global es 0.0440, en el escenario "
          r"más exigente (sin mutación ni elitismo), y los errores promedio son 0.0273 en calidad y 0.0155 "
          r"en diversidad.")
    inf.tabla("rmse", "Resultados de calibración (E1 a E5) y validación fuera de muestra (E6 a E8).",
              ["ID", "Escenario", "Conjunto", r"$\text{RMSE}_Q$", r"$\text{RMSE}_D$", "Máx.", "Cumple"],
              [["E1", r"Base ($P_{mut}=0.15$, $E=1$)", "Calibración", "0.0186", "0.0137", "0.0186", "Sí"],
               ["E2", r"Sin mutación ($P_{mut}=0$, $E=1$)", "Calibración", "0.0145", "0.0072", "0.0145", "Sí"],
               ["E3", r"Baja mutación ($P_{mut}=0.05$, $E=1$)", "Calibración", "0.0372", "0.0187", "0.0372",
                "Sí"],
               ["E4", r"Sin elitismo ($P_{mut}=0.15$, $E=0$)", "Calibración", "0.0098", "0.0183", "0.0183",
                "Sí"],
               ["E5", r"Capacidad vinculante ($C=55$)", "Calibración", "0.0372", "0.0254", "0.0372", "Sí"],
               ["E6", r"Sin mutación ni elitismo ($P_{mut}=0$, $E=0$)", "Validación", "0.0440", "0.0152",
                "0.0440", "Sí"],
               ["E7", r"Alta mutación ($P_{mut}=0.25$, $E=1$)", "Validación", "0.0387", "0.0124", "0.0387",
                "Sí"],
               ["E8", r"Bajo cruce ($P_{cr}=0.50$, $P_{mut}=0.15$)", "Validación", "0.0186", "0.0130",
                "0.0186", "Sí"],
               ["Global", "Promedio (RMSE) y máximo", "8 escenarios", "0.0273", "0.0155", "0.0440", "8/8"]],
              [700, 3000, 1250, 950, 950, 909, 1000], size=18,
              alinear=["center", "left", "center", "center", "center", "center", "center"],
              fuente=r"hoja 05_Calibracion_Validacion (500 réplicas Monte Carlo por escenario). Salvo "
                     r"indicación, $P_{cr}=0.89$ y $C=140$.")

    inf.titulo2("Condiciones extremas")
    inf.p(r"La validación estructural se completó con cinco condiciones extremas (Tabla {tab:extremos}). "
          r"En todas, el modelo produce el comportamiento esperado, sin valores negativos ni mayores que 1 "
          r"y respetando los límites físicos de los niveles.")
    inf.tabla("extremos", "Verificación de condiciones extremas.",
              ["Prueba", "Condición", r"$\Phi_{fact}$", "$D^*$", "$Q^*$", "Comportamiento del modelo"],
              [["CE-1", r"Sin mutación ($P_{mut}=0$, $E=1$)", "1.0000", "0", "Estanc.",
                r"$D^*=0$ exacto; decaimiento exponencial de $D$, con $I_D=O_Q=0$"],
               ["CE-2", r"Mutación máxima ($P_{mut}=0.5$, $E=0$)", "1.0000", "0.9007", "0.5605",
                r"Diversidad muy alta y calidad arrastrada cerca de $Q_{azar}$"],
               ["CE-3", r"Sin elitismo ($P_{mut}=0.15$, $E=0$)", "1.0000", "0.7312", "0.6358",
                r"Toda la población expuesta a la carga mutacional; menor $Q^*$"],
               ["CE-4", r"Capacidad extrema ($C=14$)", "0.0781", "0.6715", "0.1132",
                r"$k_Q$ se reduce en proporción a $\Phi_{fact}$; ganancia casi nula"],
               ["CE-5", r"Capacidad nula ($C=0$)", "0.0039", "0.6715", "0.0038",
                r"Solo 00000000 es factible ($Z^*=0$); no hay mejora posible"]],
              [800, 2200, 950, 900, 900, 3009], size=18,
              alinear=["center", "left", "center", "center", "center", "left"],
              fuente="hoja 06_Escenarios_y_Sensibilidad.")

    inf.titulo2("Sensibilidad ante incertidumbre de ±10 %")
    inf.p(r"Las sensibilidades logarítmicas del equilibrio de diversidad se obtienen derivando $D^*$:",
          after=60)
    inf.ecuacion(r"S^{D^*}_{r_{mut}}=\frac{r_{mut}}{D^*}\frac{\partial D^*}{\partial r_{mut}}=1-D^*=+0.3285,"
                 r"\qquad S^{D^*}_{k_D}=\frac{k_D}{D^*}\frac{\partial D^*}{\partial k_D}=-\left(1-D^*\right)"
                 r"=-0.3285")
    inf.p(r"es decir, un cambio del 10 % en cualquiera de las dos tasas desplaza el equilibrio alrededor "
          r"de un 3.3 %. La Tabla {tab:sensibilidad} resume las perturbaciones de ±10 %: en el peor caso "
          r"conjunto el equilibrio cae a $D^*=0.6258$ (−6.8 %) y $Q^*=0.7134$ (−4.5 %), y en el mejor sube "
          r"a $D^*=0.7142$ (+6.4 %) y $Q^*=0.7797$ (+4.4 %). El equilibrio es robusto: ninguna "
          r"perturbación cambia el régimen cualitativo del sistema.")
    inf.tabla("sensibilidad", "Sensibilidad del equilibrio ante perturbaciones de ±10 %.",
              ["Escenario", "$D^*$", "$Q^*$", r"$\Delta D^*$", r"$\Delta Q^*$", "Interpretación"],
              [["Base", "0.6715", "0.7470", "0.0 %", "0.0 %", "Punto nominal"],
               [r"$r_{mut}$ +10 %", "0.6922", "0.7508", "+3.1 %", "+0.5 %", r"Más inyección eleva $D^*$"],
               [r"$r_{mut}$ −10 %", "0.6479", "0.7426", "−3.5 %", "−0.6 %", r"Menos inyección reduce $D^*$"],
               [r"$k_D$ +10 %", "0.6502", "0.7430", "−3.2 %", "−0.5 %", r"Más consumo reduce $D^*$"],
               [r"$k_D$ −10 %", "0.6943", "0.7512", "+3.4 %", "+0.6 %", r"Menos consumo eleva $D^*$"],
               [r"$k_Q$ +10 %", "0.6715", "0.7589", "0.0 %", "+1.6 %", r"Más ganancia eleva $Q^*$"],
               [r"$k_{deg}$ +10 %", "0.6715", "0.7351", "0.0 %", "−1.6 %",
                r"Más carga mutacional reduce $Q^*$"],
               ["Peor caso conjunto", "0.6258", "0.7134", "−6.8 %", "−4.5 %", "Cota inferior de robustez"],
               ["Mejor caso conjunto", "0.7142", "0.7797", "+6.4 %", "+4.4 %", "Cota superior de robustez"]],
              [2000, 950, 950, 1000, 1000, 2859], size=18,
              alinear=["left", "center", "center", "center", "center", "left"],
              fuente=r"hoja 06_Escenarios_y_Sensibilidad. Peor caso: −10 % en $r_{mut}$ y $k_Q$, +10 % "
                     r"en $k_D$ y $k_{deg}$; el mejor caso invierte los signos.")


# =============================================================================
# Capitulo 7. Experimentacion (Fase 7)
# =============================================================================
def experimentacion(inf):
    inf.titulo1("Experimentación y análisis estadístico")
    inf.p(r"Con el modelo validado, la fase 7 comparó políticas de operación del AG y contrastó "
          r"estadísticamente el efecto de la mutación, el cruce y el elitismo sobre réplicas "
          r"independientes del AG.")

    inf.titulo2("Comparación de políticas")
    inf.p(r"Las ocho configuraciones se compararon a lo largo del horizonte completo mediante la integral "
          r"acumulada de la brecha del módulo 10, que mide la exposición total a soluciones subóptimas:",
          after=60)
    inf.ecuacion(r"E_B=\int_0^T B(t)\,dt\approx\sum_{n=0}^{M-1}B_n\,\Delta t\quad[\text{gen}]")
    inf.p(r"La Tabla {tab:politicas} muestra el resultado. La política sin mutación (E2) alcanza la mayor "
          r"calidad media y la menor integral de brecha, pero lo hace agotando la diversidad "
          r"($D(30)=0.0011$): la población se vuelve clónica y pierde la capacidad de seguir explorando, "
          r"que es precisamente la convergencia prematura del régimen I. Entre las políticas que conservan "
          r"exploración, la de baja mutación con elitismo (E3) ofrece el mejor compromiso —$Q(30)=0.8193$, "
          r"$D(30)=0.4053$ y $E_B=6.29$ generaciones, frente a 8.08 en la configuración de clase (E1) y "
          r"11.30 sin elitismo (E4)—. La capacidad vinculante (E5) reduce la calidad alcanzable por la pena "
          r"de muerte, y el exceso de mutación (E7) eleva la diversidad a costa de una fuerte carga "
          r"mutacional.")
    inf.tabla("politicas", "Comparación de políticas con el modelo de dos niveles ($T=30$, $\\Delta t=1$).",
              ["ID", "Política", "$D^*$", r"$\tau_D$", "$Q(30)$", "$D(30)$", r"$\bar{Z}(30)$", "$E_B$",
               "Diagnóstico"],
              [["E1", "Base de clase", "0.6715", "1.65", "0.7470", "0.6715", "52.29", "8.08",
                "Equilibrio alto con exploración activa"],
               ["E2", "Sin mutación", "0", "5.01", "0.8810", "0.0011", "61.67", "4.88",
                r"Estancamiento: $D\to0$ anula $I_Q$"],
               ["E3", "Baja mutación", "0.4053", "2.98", "0.8193", "0.4053", "57.35", "6.29",
                "Mejor compromiso calidad–diversidad"],
               ["E4", "Sin elitismo", "0.7312", "1.79", "0.6358", "0.7312", "44.50", "11.30",
                r"Carga mutacional lleva $Q^*$ hacia $Q_{azar}$"],
               ["E5", "Capacidad $C=55$", "0.6715", "1.65", "0.6421", "0.6715", "29.54", "11.35",
                r"La pena de muerte reduce $\Phi_{fact}$ y $Z^*$"],
               ["E6", "Sin mutación ni elitismo", "0", "6.67", "0.7570", "0.0067", "52.99", "8.67",
                "Estancamiento prematuro (peor caso)"],
               ["E7", "Alta mutación", "0.7731", "1.14", "0.7014", "0.7731", "49.10", "9.28",
                "Alta diversidad con fuerte carga mutacional"],
               ["E8", "Bajo cruce", "0.6715", "1.65", "0.7435", "0.6715", "52.04", "8.18",
                "Menor velocidad inicial de ganancia"]],
              [500, 1450, 750, 600, 800, 800, 800, 700, 2359], size=16,
              alinear=["center", "left"] + ["center"] * 6 + ["left"],
              fuente=r"hoja 06_Escenarios_y_Sensibilidad. $\tau_D$ y $E_B$ en generaciones; $\bar{Z}$ en "
                     r"unidades de utilidad. Los parámetros de cada escenario son los de la Tabla "
                     r"{tab:rmse}.")

    inf.titulo2("Pruebas no paramétricas")
    inf.p(r"Las pruebas se ejecutaron con fórmulas nativas de Excel (distribuciones normal y "
          r"chi-cuadrado) sobre réplicas independientes del AG, con los estadísticos:", after=60)
    inf.ecuacion(r"Z_W=\frac{\left|W^+-\mu_W\right|}{\sigma_W},\qquad H=\frac{12}{N(N+1)}\sum_{j=1}^{k}"
                 r"\frac{R_j^2}{n_j}-3(N+1)", after=60)
    inf.ecuacion(r"Z_U=\frac{\left|U_1-n_1n_2/2\right|}{\sqrt{n_1n_2\left(n_1+n_2+1\right)/12}}")
    inf.p(r"La prueba de Wilcoxon, sobre las mismas 30 poblaciones iniciales con $P_{mut}=0.05$ y "
          r"$P_{mut}=0$, muestra que la mutación sostiene significativamente la diversidad: las 30 "
          r"diferencias son positivas ($W^+=450$, el máximo posible). Kruskal-Wallis confirma que el nivel "
          r"de mutación cambia la calidad final, que decrece de forma monótona con $P_{mut}$; la mayor "
          r"calidad media sin mutación se explica porque la población converge a copias de un buen "
          r"individuo, pero con diversidad prácticamente nula, es decir, con la búsqueda detenida. "
          r"Mann-Whitney confirma que el elitismo eleva la calidad final (Tabla {tab:noparam}).")
    inf.tabla("noparam", r"Resultados de las pruebas no paramétricas ($\alpha=0.05$, 30 réplicas por grupo).",
              ["Prueba", "Comparación", "Variable", "Medias por grupo", "Estadístico", "Valor $p$",
               "Decisión"],
              [["Wilcoxon (pareada)", r"$P_{mut}=0.05$ frente a $P_{mut}=0$", "$D(30)$",
                "0.4130 frente a 0.0039", r"$W^+=450$; $Z=4.474$", r"$7.69\times10^{-6}$",
                r"Rechazar $H_0$"],
               ["Kruskal-Wallis", r"$P_{mut}\in\{0;\,0.05;\,0.15;\,0.25\}$", "$Q(30)$",
                "0.910; 0.860; 0.720; 0.678", r"$H=77.44$", r"$1.08\times10^{-16}$", r"Rechazar $H_0$"],
               ["Mann-Whitney U", r"$E=1$ frente a $E=0$ ($P_{mut}=0.15$)", "$Q(30)$",
                "0.720 frente a 0.613", r"$U_1=721$; $Z=4.007$", r"$6.16\times10^{-5}$",
                r"Rechazar $H_0$"]],
              [1150, 1900, 950, 1400, 1350, 1100, 909], size=17,
              alinear=["left", "left", "center", "center", "center", "center", "center"],
              fuente="hoja 07_Analisis_Estadistico.")

    inf.titulo2("ANOVA de tres factores")
    inf.p(r"El ANOVA sobre $Q(30)$ en el diseño factorial $2\times2\times2$ ($P_{mut}\in\{0.05;\,0.15\}$, "
          r"$P_{cr}\in\{0.50;\,0.89\}$, $E\in\{0;\,1\}$, 20 réplicas por celda) descompone la variabilidad "
          r"total en efectos principales, interacciones y error, con $F_X=MS_X/MS_{Err}$ y "
          r"$\eta^2_X=SS_X/SS_T$ (Montgomery, 2017). El elitismo es el factor dominante (34.2 % de la "
          r"varianza), seguido de la mutación (23.4 %); el cruce no es significativo ($p=0.718$) y ninguna "
          r"interacción lo es ($p>0.25$), como muestra la Tabla {tab:anova}. La mejor celda es "
          r"$P_{mut}=0.05$, $P_{cr}=0.89$, $E=1$, con calidad media 0.892, y la peor es $P_{mut}=0.15$, "
          r"$P_{cr}=0.89$, $E=0$, con 0.609.")
    inf.tabla("anova", r"ANOVA de tres factores sobre $Q(30)$ ($N=160$, $\alpha=0.05$).",
              ["Fuente de variación", "SS", "g. l.", "MS", "$F$", "Valor $p$", r"$\eta^2$"],
              [[r"A: mutación ($P_{mut}$)", "0.5689", "1", "0.5689", "85.59", r"$1.92\times10^{-16}$",
                "23.4 %"],
               [r"B: cruce ($P_{cr}$)", "0.0009", "1", "0.0009", "0.13", "0.7179", "0.04 %"],
               [r"C: elitismo ($E$)", "0.8325", "1", "0.8325", "125.23", r"$1.39\times10^{-21}$", "34.2 %"],
               ["A × B", "0.0011", "1", "0.0011", "0.17", "0.6796", "0.05 %"],
               ["A × C", "0.0047", "1", "0.0047", "0.71", "0.4023", "0.19 %"],
               ["B × C", "0.0081", "1", "0.0081", "1.22", "0.2703", "0.33 %"],
               ["A × B × C", "0.0086", "1", "0.0086", "1.29", "0.2585", "0.35 %"],
               ["Error", "1.0104", "152", "0.0066", "–", "–", "41.5 %"],
               ["Total", "2.4352", "159", "–", "–", "–", "100 %"]],
              [2659, 1000, 800, 1000, 1000, 1300, 1000], size=18,
              alinear=["left"] + ["center"] * 6,
              fuente="hoja 07_Analisis_Estadistico.")

    inf.titulo2("Contraste de hipótesis y política recomendada")
    inf.p(r"**Hipótesis 1, confirmada.** El modelo de dos niveles reproduce el comportamiento medio del "
          r"AG con RMSE ≤ 0.05 en los ocho escenarios, incluidos tres fuera de muestra, con simulación "
          r"numéricamente estable; además, la solución analítica demuestra que sin mutación la calidad se "
          r"estanca en $Q_\infty<1$, y la prueba de Wilcoxon confirma que sin mutación la diversidad final "
          r"es prácticamente nula.")
    inf.p(r"**Hipótesis 2, confirmada.** En la generación 1 de clase el nuevo mejor individuo ($Z=67$) "
          r"surgió sin cruce, por mutación del bit 7. Estadísticamente, la mutación tiene un efecto "
          r"significativo y de doble signo: sostiene la diversidad (Wilcoxon) pero reduce la calidad media "
          r"cuando es excesiva (Kruskal-Wallis y factor A del ANOVA), tal como predice la dualidad "
          r"exploratoria–disruptiva.")
    inf.p(r"**Efecto del elitismo, significativo.** Mann-Whitney y el factor C del ANOVA muestran que "
          r"preservar una copia élite eleva la calidad estacionaria, al proteger una fracción $1/N$ de la "
          r"población de la carga mutacional y asegurar un piso para $Z_{max}$.")
    inf.p(r"**Efecto del cruce, no significativo.** El bajo coeficiente calibrado $k_{cr}=0.0218$ y el "
          r"factor B del ANOVA ($p=0.718$) coinciden: en este espacio pequeño (256 genotipos) la "
          r"recombinación aporta poco frente a la selección, la mutación y el elitismo. La ausencia de "
          r"interacciones significativas respalda, además, la separación lineal y aditiva de las tasas "
          r"$k_Q$ y $k_{deg}$ del modelo.")
    inf.p(r"**Política recomendada.** Para la mochila de 8 bits con $N=8$ se recomienda operar con "
          r"$P_{mut}=0.05$, $P_{cr}=0.89$ y elitismo $E=1$: es la configuración de mayor calidad media "
          r"entre las que conservan diversidad positiva ($D^*\approx0.41$), con la menor integral de "
          r"brecha de ese grupo y sin riesgo de estancamiento. La recomendación supone la población, la "
          r"codificación y los operadores de clase; para otros tamaños de población o capacidades deben "
          r"recalibrarse las tasas.")


# =============================================================================
# Capitulo 8. Conclusiones
# =============================================================================
def conclusiones(inf):
    inf.titulo1("Conclusiones")
    inf.p(r"Se cumplió el objetivo de modelar un AG como sistema dinámico: sus variables se clasificaron "
          r"en dos niveles (diversidad y calidad), cuatro flujos, una brecha o backlog, entradas, "
          r"parámetros y salidas, integrados en un diagrama causal de 17 enlaces y cuatro bucles de "
          r"balance, un diagrama de Forrester y ecuaciones dimensionalmente consistentes.")
    inf.p(r"La convergencia del AG se explica por su estructura de realimentación: depende del cambio de "
          r"dominancia entre la búsqueda asintótica (B1), el agotamiento de diversidad (B2), la reposición "
          r"mutacional (B3) y el freno por carga mutacional (B4), acoplados por el enlace diversidad–"
          r"ganancia. Sin mutación, B2 apaga a B1 y el algoritmo converge prematuramente; con mutación "
          r"moderada y elitismo alcanza un equilibrio.")
    inf.p(r"El modelo agregado de dos niveles es suficiente para reproducir el comportamiento medio del "
          r"AG (hipótesis 1): con un único vector de siete parámetros cumplió RMSE ≤ 0.05 en los ocho "
          r"escenarios (máximo 0.044), superó las pruebas dimensional, de condiciones extremas y de "
          r"sensibilidad, y su simulación por Euler es estable con paso de una generación.")
    inf.p(r"La mutación cumple una doble función (hipótesis 2): en la corrida de clase produjo "
          r"directamente al mejor individuo sin cruce, sostiene significativamente la diversidad y, en "
          r"exceso, reduce la calidad media por carga mutacional. El elitismo (34.2 % de la varianza) y la "
          r"mutación (23.4 %) gobiernan la calidad final; el cruce no es significativo. Se recomienda "
          r"$P_{mut}=0.05$, $P_{cr}=0.89$ y $E=1$, que maximiza la calidad entre las políticas que "
          r"conservan exploración; la configuración de clase ($P_{mut}=0.15$) mantiene una diversidad alta, "
          r"pero su carga mutacional limita la calidad media a 0.75.")
    inf.p(r"El modelo describe la trayectoria media, no la variabilidad de corridas individuales, y sus "
          r"tasas calibradas valen para 256 genotipos, 8 individuos y los operadores de clase. Como trabajo "
          r"futuro se propone recalibrarlo para poblaciones y cromosomas mayores, incorporar operadores "
          r"alternativos (torneo, cruce multipunto, mutación adaptativa) y explorar capacidades vinculantes "
          r"intermedias. De la fase 8 quedan pendientes las diapositivas de sustentación.")


# =============================================================================
# Anexo A. Bitacora de IA
# =============================================================================
BITACORA = [
    ("2026-11-17", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes)",
     "Solicitud de transcripción, organización y estructuración de los apuntes manuscritos escaneados del "
     "planteamiento del proyecto (SCAN_20261117_071824490.pdf).",
     "Transcripción y documentación formal del problema, hipótesis, objetivos, sistema experimental y marco "
     "metodológico en Proyecto_Dinamica_de_Sistemas.md.",
     "**Alto (Dirección y autoría intelectual):** Las ideas centrales, hipótesis, diseño del problema y "
     "selección del AG de clase fueron definidos por el estudiante y su compañero; la IA solo transcribió y "
     "estructuró en Markdown.",
     "Se aprendió a formalizar una idea inicial de investigación en un documento trazable separando "
     "problema, hipótesis contrastables, sistema experimental y fases metodológicas."),
    ("2026-11-17", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + Python (zipfile/XML)",
     "Solicitud de análisis técnico de los libros de Excel y scripts de Python trabajados durante el "
     "semestre en Sistemas Inteligentes Computacionales (Mochila.xlsx, Mochila-valoresDec.xlsx, "
     "mochilaAGVacio, mochilacorregida26.py).",
     r"Inventario y auditoría técnica en Analisis_Excels_AG.md: identificación de parámetros reales "
     r"($C=140$, $N=8$, $L=8$), detección de 5 errores estructurales en los Excel de clase y especificación "
     r"del nuevo diseño automatizado por iteración.",
     "**Medio-Alto (Curaduría y criterio técnico):** El estudiante suministró los archivos propios del "
     "semestre, delimitó cuáles comparar y validó la inconsistencia entre capacidad 14 vs 140 y los errores "
     "de ruleta.",
     r"Se comprendió con precisión el impacto de la pena de muerte en la ruleta proporcional, por qué "
     r"$C=140$ deja inactiva la restricción en 8 bits ($W_{max}=87$) y cómo evitar errores de arrastre y "
     r"fórmulas matriciales en Excel."),
    ("2026-11-17", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + Python",
     "Solicitud de síntesis de la propuesta formal integrando el marco conceptual, diagrama causal, diagrama "
     "de Forrester, ecuaciones preliminares y validación inicial.",
     r"Redacción de Propuesta_Dinamica_de_Sistemas.md con clasificación de variables de Dinámica de "
     r"Sistemas, bucles B1–B4, formulación de 2 niveles (Diversidad $D$ y Calidad $Q$) y primera "
     r"calibración en Python.",
     "**Medio (Supervisión conceptual):** El estudiante revisó la propuesta generada y detectó "
     "posteriormente que varias fórmulas contenían expresiones ajenas al curso (constantes y potencias "
     "arbitrarias) que debían corregirse.",
     "Se comprendió cómo traducir un algoritmo evolutivo discreto en variables de estado, flujos y brechas, "
     "y se desarrolló criterio crítico para detectar sobreparametrización no alineada con la teoría vista "
     "en clase."),
    ("2026-11-19", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + Python (simulación Monte Carlo)",
     "Solicitud de analizar el proyecto usando Propuesta_Dinamica_de_Sistemas.md, Analisis_Excels_AG.md y "
     "la base teórica exacta del curso (dinamica-de-sistemas-conocimiento.md), corrigiendo las fórmulas poco "
     "concordes con lo visto en clase.",
     r"Reformulación matemática del modelo DS de 2 niveles siguiendo estrictamente los Módulos 1–10 del "
     r"curso: eliminación de constantes ad-hoc (0.25, $\Phi^{1.6558}$, $1-(1-P_{mut})^2$) y sustitución "
     r"por ajuste hacia meta, drenaje proporcional, isomorfismo Tanque–TI–AG, equilibrio analítico ($D^*$, "
     r"$Q^*$, $\tau$), Euler explícito ($n\to n+1$) con estabilidad $|g_D|<1$ y sensibilidad ±10 %.",
     "**Alto (Criterio disciplinar y corrección metodológica):** El estudiante aportó la base de "
     "conocimiento completa del curso, detectó las fórmulas «raras» del plan previo y exigió estricto apego "
     "a los modelos lineales/canónicos de clase.",
     r"Se afianzó el dominio de los 10 módulos de Dinámica de Sistemas (consistencia dimensional, "
     r"equilibrio dinámico $dX/dt=0$, constante de tiempo $\tau=1/k$, regla de un solo instante en Euler y "
     r"criterio de estabilidad numérica de Euler)."),
    ("2026-11-19", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + Python (openpyxl) + Excel COM",
     "Solicitud de crear el libro de Excel aplicando todos los pasos del proyecto, manteniendo la forma de "
     "los Excels de clase y las fórmulas corregidas del curso.",
     r"Construcción de Modelo_AG_Dinamica_de_Sistemas.xlsx (8 hojas, 8,856 fórmulas, 7 gráficas nativas) y "
     r"su generador reproducible generar_modelo_excel.py: panel dinámico sobre $2^8=256$ estados, AG paso a "
     r"paso corregido (Gen 0–5), Euler ($\Delta t=1$ y 0.5), calibración/validación en 8 escenarios (500 "
     r"réplicas MC, 8/8 RMSE ≤ 0.05) y pruebas estadísticas (Wilcoxon, Kruskal-Wallis, Mann-Whitney y "
     r"ANOVA 3 factores).",
     "**Medio (Especificación y validación):** El estudiante definió la estructura esperada según "
     "Analisis_Excels_AG.md y validó el entregable abriéndolo directamente en Microsoft Excel.",
     "Se integraron en un único modelo auditable la simulación microscópica de un AG por iteración y su "
     "modelo macroscópico en Dinámica de Sistemas, verificando estadísticamente que elitismo y mutación "
     "gobiernan el desempeño estacionario."),
    ("2026-11-19", "Andrés Hernández",
     "Antigravity IDE / Gemini (Enjambre /swarm de 4 agentes) + Servidor MCP excel + Excel COM",
     "Solicitud en modo enjambre (/swarm) de corregir problemas visuales donde algunas letras no se veían "
     "bien en las casillas y estandarizar todo el libro en tipografía Times New Roman usando el MCP de "
     "Excel.",
     "Auditoría visual con capturas de pantalla vía MCP (excel_describe_sheets, excel_read_sheet, "
     "excel_screen_capture), eliminación de 24 celdas ##### en 03_AG_Paso_a_Paso, ajuste de anchos de "
     "columna y alturas de fila en las 8 hojas, migración de 14,964/14,964 celdas y 7/7 gráficas a Times "
     "New Roman, recolocación de leyendas sin solapamiento y verificación final con 0 errores de fórmula.",
     "**Medio-Alto (Control de calidad visual y requisito de formato):** El estudiante inspeccionó el "
     "archivo real en Excel, detectó los recortes de texto y fijó el estándar tipográfico (Times New Roman) "
     "y el uso del MCP.",
     "Se comprendió la importancia de auditar visualmente hojas complejas (evitando herencia de fuente "
     "scheme=\"minor\" hacia Calibri, errores ##### por ancho estrecho y solapamientos de leyendas en "
     "gráficas) y de separar roles de auditoría, edición y verificación."),
    ("2026-11-19", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + PowerShell / Python",
     "Reporte de que el libro aparecía bloqueado en modo solo lectura al intentar editarlo y solicitud de "
     "desbloquearlo y generar la bitácora formal de IA.",
     r"Diagnóstico de permisos/locks de Windows y XML del .xlsx, terminación del proceso huérfano "
     r"EXCEL.EXE /automation -Embedding que mantenía el bloqueo ~\$Modelo_AG_Dinamica_de_Sistemas.xlsx, "
     r"verificación de acceso lectura/escritura y generación de esta bitácora en "
     r"Bitacora_IA_Proyecto_Dinamica.md.",
     "**Medio (Pruebas de usuario y trazabilidad académica):** El estudiante detectó el bloqueo operativo "
     "al abrir el archivo para edición propia y definió el formato exacto de registro ético de uso de IA.",
     r"Se identificó cómo los procesos de automatización COM de Excel pueden dejar archivos temporales de "
     r"bloqueo (~\$...) que fuerzan el modo solo lectura, y se documentó de forma transparente la coautoría "
     r"asistida por IA."),
    ("2026-11-19", "Andrés Hernández", "Antigravity IDE / Gemini (Hermes) + PowerShell / LaTeX",
     "Solicitud de organizar el directorio del proyecto y generar un documento en LaTeX con todas las "
     "fórmulas explicadas y desarrolladas paso a paso.",
     r"Reorganización modular del repositorio (docs/, docs/referencias/, excel/, latex/, src/, README.md), "
     r"eliminación de copia redundante y redacción de latex/Formulas_y_Desarrollo_Matematico.tex (7 "
     r"secciones con deducciones analíticas de $D(t)$, $D^*$, $\tau_D$, $Q^*$, $Q_\infty$, desarrollo manual "
     r"de Euler $n=0,1,2$, estabilidad $|g_D|<1$, sensibilidad ±10 % y pruebas estadísticas).",
     "**Medio (Estructuración de entregables):** El estudiante definió la necesidad de ordenar el "
     "repositorio y disponer de un documento matemático formal en LaTeX para sustentar cada ecuación del "
     "Excel.",
     "Se consolidó la trazabilidad completa entre el modelo conceptual, las deducciones algebraicas paso a "
     "paso, la simulación numérica en Excel y la evidencia estadística en formato académico reproducible."),
    ("2026-11-20", "Andrés Hernández",
     "Claude Code (modelo Claude Opus 5.5, Anthropic) + Python (python-docx, matplotlib, latex2mathml) + "
     "Microsoft Word (COM)",
     "Solicitud de analizar todo el proyecto y crear un informe siguiendo estrictamente el formato "
     "institucional «Formato para presentación de trabajos ASI – Word», agregando la bitácora de IA, "
     "verificando el documento página por página (sin errores ni páginas en blanco) y creando el commit y "
     "el push al repositorio.",
     "Redacción y maquetación del informe final (informe/Informe_Final_Proyecto_Dinamica.docx y .pdf) a "
     "partir de los entregables de las fases 1 a 8, el libro de Excel y el compendio LaTeX; generación del "
     "diagrama de Forrester; conversión de las ecuaciones a formato nativo de Word; inclusión de esta "
     "bitácora como anexo; verificación visual página por página y versionamiento en Git (rama "
     "Rama_avance).",
     "**Medio (Especificación y control de calidad):** El estudiante definió la plantilla institucional, el "
     "alcance del informe, la inclusión de la bitácora y los criterios de verificación; el contenido "
     "técnico proviene de los entregables previos del proyecto y la revisión final del documento "
     "corresponde a los autores.",
     "Se consolidó en un único documento con formato institucional la trazabilidad entre la "
     "conceptualización causal, la formulación matemática, la simulación, la validación estadística y el "
     "registro ético del uso de IA, como base para la sustentación final del proyecto."),
]


def anexo_bitacora(inf):
    inf.titulo1("A. Anexo: Bitácora de uso de inteligencia artificial", numerado=False)
    inf.p(r"En cumplimiento de los lineamientos de uso ético y transparente de la inteligencia artificial "
          r"(IA), este anexo transcribe la bitácora formal del proyecto, que se mantiene en la carpeta de "
          r"documentación del repositorio. Cada registro documenta la fecha, el nombre del "
          r"estudiante, la herramienta utilizada, el resumen de la instrucción (prompt), la tarea en la que "
          r"ayudó la IA, el nivel de intervención humana y el resultado de aprendizaje (Tabla "
          r"{tab:bitacora}). El registro 9 corresponde a la elaboración de este informe.")
    filas = []
    campos = ["Fecha", "Nombre", "Herramienta utilizada", "Resumen del prompt", "Tarea ayudada",
              "Nivel de intervención humana", "Resultado de aprendizaje"]
    for i, reg in enumerate(BITACORA, start=1):
        filas.append({"grupo": f"Registro {i}"})
        for campo, valor in zip(campos, reg):
            filas.append([f"**{campo}**", valor])
    inf.tabla("bitacora", "Bitácora de uso de inteligencia artificial en el proyecto.",
              ["Campo", "Registro"], filas, [2300, 6459], size=17, unida="grupos",
              fuente="*docs/Bitacora_IA_Proyecto_Dinamica.md*.")
    inf.p(r"**Declaración de uso de IA.** Las herramientas de IA se emplearon como asistentes de "
          r"transcripción, análisis, programación, verificación y redacción. Las ideas centrales, las "
          r"hipótesis, el diseño del problema, la selección del sistema experimental y los criterios de "
          r"validación fueron definidos por los autores, a quienes corresponden la revisión y la "
          r"responsabilidad final del contenido.", after=0)


# =============================================================================
# Bibliografia (seccion independiente de la plantilla)
# =============================================================================
REFERENCIAS = [
    "Fisher, R. A. (1930). *The genetical theory of natural selection*. Clarendon Press.",
    "Forrester, J. W. (1961). *Industrial dynamics*. MIT Press.",
    "Goldberg, D. E. (1989). *Genetic algorithms in search, optimization, and machine learning*. "
    "Addison-Wesley.",
    "Holland, J. H. (1975). *Adaptation in natural and artificial systems*. University of Michigan Press.",
    "Kruskal, W. H., y Wallis, W. A. (1952). Use of ranks in one-criterion variance analysis. *Journal of "
    "the American Statistical Association, 47*(260), 583–621. https://doi.org/10.1080/01621459.1952.10483441",
    "Mann, H. B., y Whitney, D. R. (1947). On a test of whether one of two random variables is "
    "stochastically larger than the other. *The Annals of Mathematical Statistics, 18*(1), 50–60. "
    "https://doi.org/10.1214/aoms/1177730491",
    "Montgomery, D. C. (2017). *Design and analysis of experiments* (9.ª ed.). Wiley.",
    "Sterman, J. D. (2000). *Business dynamics: Systems thinking and modeling for a complex world*. "
    "Irwin/McGraw-Hill.",
    "Universidad Nacional de Colombia. (2026a). *Dinámica de sistemas: laboratorio para administración de "
    "sistemas informáticos* [Material del curso, módulos 1 a 10]. Sede Manizales.",
    "Universidad Nacional de Colombia. (2026b). *Algoritmos genéticos* [Diapositivas de clase, Sistemas "
    "Inteligentes Computacionales]. Sede Manizales.",
    "Wilcoxon, F. (1945). Individual comparisons by ranking methods. *Biometrics Bulletin, 1*(6), 80–83. "
    "https://doi.org/10.2307/3001968",
]


def bibliografia(inf):
    inf.titulo1("Bibliografía", numerado=False, salto=False)
    for ref in REFERENCIAS:
        inf.p(ref, jc="left", line=276, after=200, ind={"left": 709, "hanging": 709})


# =============================================================================
# Ensamble, actualizacion del indice y exportacion
# =============================================================================
def construir():
    FIG_FORRESTER.parent.mkdir(parents=True, exist_ok=True)
    dibujar_forrester(FIG_FORRESTER)
    inf = Informe(PLANTILLA)
    inf.preparar_plantilla(TITULO, AUTORES, CURSO, ANIO, TITULO_CORTO)
    inf.predefinir(FIGURAS, TABLAS)
    resumen(inf)
    inf.contenido()
    introduccion(inf)
    antecedentes(inf)
    metodologia(inf)
    conceptualizacion(inf)
    formulacion(inf)
    simulacion(inf)
    validacion(inf)
    experimentacion(inf)
    conclusiones(inf)
    anexo_bitacora(inf)
    inf.cerrar_seccion1()
    bibliografia(inf)
    inf.guardar(SALIDA)
    return SALIDA


def actualizar_y_exportar(docx_path):
    """Actualiza el indice con Microsoft Word, guarda y exporta a PDF."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    pdf = docx_path.with_suffix(".pdf")
    if pdf.exists():
        pdf.unlink()
    try:
        doc = word.Documents.Open(str(docx_path), ConfirmConversions=False, ReadOnly=False,
                                  AddToRecentFiles=False, Visible=False)
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        for toc in doc.TablesOfContents:
            toc.UpdatePageNumbers()
        doc.Save()
        paginas = doc.ComputeStatistics(2)
        doc.ExportAsFixedFormat(OutputFileName=str(pdf), ExportFormat=17, OpenAfterExport=False,
                                OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True, KeepIRM=True,
                                CreateBookmarks=1, DocStructureTags=True, BitmapMissingFonts=True,
                                UseISO19005_1=False)
        doc.Close(False)
    finally:
        word.Quit()
    if not pdf.exists():
        raise RuntimeError("Word no genero el PDF.")
    return pdf, paginas


if __name__ == "__main__":
    salida = construir()
    print("DOCX:", salida)
    if "--pdf" in sys.argv:
        pdf, paginas = actualizar_y_exportar(salida)
        print("PDF:", pdf, "| paginas:", paginas)
