"""
app.py
======
Buscador interactivo de Planes de Cuidados de Enfermería a partir de un
archivo Excel institucional.

Autor: Senior Python Developer (Data Processing)
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Configuración general
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Buscador de Planes de Enfermería",
    page_icon="🩺",
    layout="wide",
)

PLACEHOLDER = "No especificado en el documento original"

PALABRAS_CLAVE_ENCABEZADO = [
    "diagnostico", "riesgo", "dominio", "hallazgo", "resultado",
    "intervencion", "actividad", "evaluacion", "prevenir", "cuidado",
    "dx", "enferm",
]
MAX_FILAS_ESCANEO_ENCABEZADO = 25
MIN_COINCIDENCIAS_ENCABEZADO = 2

SINONIMOS_COLUMNAS = {
    "riesgo": ["riesgo a prevenir", "riesgo"],
    "dominio": ["dominio"],
    "hallazgos": ["hallazgo"],
    "diagnostico": ["diagnostico"],
    "diagnostico_alt": ["dx"],
    "resultado": ["resultado esperado", "resultado"],
    "intervencion": ["intervencion"],
    "actividad": ["actividad"],
    "evaluacion": ["evaluacion"],
}


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------

def _normalizar(texto: object) -> str:
    """Normaliza un valor a texto en minúsculas sin tildes, seguro ante NaN."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    s = str(texto).strip().lower()
    for origen, destino in {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n"}.items():
        s = s.replace(origen, destino)
    return s


def _normalizar_espacios(texto: str) -> str:
    """Colapsa saltos de línea y espacios múltiples en un solo espacio,
    para que un valor de celda con saltos de línea no rompa el diseño
    de un botón o una etiqueta de una sola línea."""
    return " ".join(texto.split())


def _es_valor_vacio(valor: object) -> bool:
    """Determina si un valor debe considerarse 'no especificado'."""
    if valor is None:
        return True
    if isinstance(valor, float) and pd.isna(valor):
        return True
    if isinstance(valor, str) and valor.strip() == "":
        return True
    return False


def _texto_limpio(valor: object) -> Optional[str]:
    """Devuelve el valor como texto limpio, sanitizando caracteres que rompen el HTML."""
    if _es_valor_vacio(valor):
        return None
    # FIX: Escapar caracteres matemáticos o flechas para no romper el DOM de React
    return str(valor).strip().replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------------------
# Detección automática de encabezado y columnas clave
# ---------------------------------------------------------------------------

def detectar_fila_encabezado(df_crudo: pd.DataFrame) -> Optional[int]:
    n_filas = min(MAX_FILAS_ESCANEO_ENCABEZADO, len(df_crudo))
    mejor_fila, mejor_puntaje = None, 0

    for i in range(n_filas):
        celdas = [_normalizar(v) for v in df_crudo.iloc[i].tolist()]
        puntaje = sum(
            1 for c in celdas if c and any(k in c for k in PALABRAS_CLAVE_ENCABEZADO)
        )
        no_vacias = sum(1 for c in celdas if c)
        if puntaje >= MIN_COINCIDENCIAS_ENCABEZADO and no_vacias >= puntaje:
            if puntaje > mejor_puntaje:
                mejor_puntaje, mejor_fila = puntaje, i

    return mejor_fila


def encontrar_columna(columnas: list[str], claves: list[str]) -> Optional[str]:
    for col in columnas:
        n = _normalizar(col)
        if any(clave in n for clave in claves):
            return col
    return None


def identificar_columnas_clave(columnas: list[str]) -> dict[str, Optional[str]]:
    mapa: dict[str, Optional[str]] = {}
    mapa["riesgo"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["riesgo"])
    mapa["dominio"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["dominio"])
    mapa["hallazgos"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["hallazgos"])
    mapa["resultado"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["resultado"])
    mapa["intervencion"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["intervencion"])
    mapa["actividad"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["actividad"])
    mapa["evaluacion"] = encontrar_columna(columnas, SINONIMOS_COLUMNAS["evaluacion"])

    diag = encontrar_columna(columnas, SINONIMOS_COLUMNAS["diagnostico"])
    if diag is None:
        diag = encontrar_columna(columnas, SINONIMOS_COLUMNAS["diagnostico_alt"])
    mapa["diagnostico"] = diag
    return mapa


# ---------------------------------------------------------------------------
# Estructuras de datos
# ---------------------------------------------------------------------------

@dataclass
class Cita:
    hoja: str
    columna: str
    fila: int

    def texto(self) -> str:
        return f"({self.hoja}, {self.columna}, Fila {self.fila})"


@dataclass
class CampoConCita:
    valor: Optional[str]
    cita: Optional[Cita]


@dataclass
class Actividad:
    valor: str
    cita: Cita


@dataclass
class Intervencion:
    texto: CampoConCita
    actividades: list[Actividad] = field(default_factory=list)


@dataclass
class Hallazgo:
    valor: str
    cita: Cita


@dataclass
class CasoEnfermeria:
    hoja: str
    fila_excel_inicio: int
    diagnostico: CampoConCita
    riesgo: CampoConCita
    dominio: CampoConCita
    resultado: CampoConCita
    evaluacion: CampoConCita
    hallazgos: list[Hallazgo] = field(default_factory=list)
    intervenciones: list[Intervencion] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Mini-ETL
# ---------------------------------------------------------------------------

def _origen_por_columna(serie: pd.Series) -> pd.Series:
    valido = serie.apply(lambda v: not _es_valor_vacio(v))
    posiciones = pd.Series(serie.index, index=serie.index)
    return posiciones.where(valido).ffill()


def _procesar_hoja(nombre_hoja: str, df_crudo: pd.DataFrame) -> tuple[list[CasoEnfermeria], list[dict], Optional[str]]:
    if df_crudo.empty:
        return [], [], f"Hoja '{nombre_hoja}' está vacía. Se omite."

    idx_encabezado = detectar_fila_encabezado(df_crudo)
    if idx_encabezado is None:
        return [], [], f"No se identificó un encabezado válido en la hoja '{nombre_hoja}'. Se omite."

    df_hoja = df_crudo.copy()
    df_hoja.columns = [
        str(c).strip() if not str(c).startswith("Unnamed") and pd.notna(c) else f"Columna_{i}"
        for i, c in enumerate(df_hoja.iloc[idx_encabezado].tolist())
    ]
    df_hoja = df_hoja.iloc[idx_encabezado + 1:].reset_index(drop=True)
    df_hoja = df_hoja.dropna(how="all").reset_index(drop=True)

    if df_hoja.empty:
        return [], [], f"La hoja '{nombre_hoja}' no tiene datos tras el encabezado. Se omite."

    columnas_clave = identificar_columnas_clave(list(df_hoja.columns))
    col_diag = columnas_clave["diagnostico"]
    if col_diag is None:
        return [], [], (
            f"La hoja '{nombre_hoja}' no tiene una columna reconocible de "
            f"'Diagnóstico de Enfermería'. Se omite."
        )

    fila_base_excel = idx_encabezado + 2
    origenes = {col: _origen_por_columna(df_hoja[col]) for col in df_hoja.columns}
    lleno = df_hoja.ffill(axis=0)

    def cita_para(col: Optional[str], fila_pos: int) -> Optional[Cita]:
        if col is None:
            return None
        origen_pos = origenes[col].get(fila_pos)
        if pd.isna(origen_pos):
            return None
        return Cita(hoja=nombre_hoja, columna=col, fila=fila_base_excel + int(origen_pos))

    def valor_para(col: Optional[str], fila_pos: int) -> Optional[str]:
        if col is None:
            return None
        return _texto_limpio(lleno.at[fila_pos, col])

    registros_planos: list[dict] = []
    for pos in range(len(df_hoja)):
        for col in df_hoja.columns:
            registros_planos.append({
                "hoja": nombre_hoja,
                "fila_excel": fila_base_excel + pos,
                "columna": col,
                "valor": df_hoja.at[pos, col],
            })

    diag_no_nulo = df_hoja[col_diag].apply(lambda v: not _es_valor_vacio(v))
    bloque_diag_id = diag_no_nulo.cumsum()

    col_interv = columnas_clave["intervencion"]
    if col_interv is not None:
        interv_no_nulo = df_hoja[col_interv].apply(lambda v: not _es_valor_vacio(v))
        bloque_interv_id = interv_no_nulo.cumsum()
    else:
        bloque_interv_id = bloque_diag_id

    df_trabajo = df_hoja.copy()
    df_trabajo["_bloque_diag"] = bloque_diag_id
    df_trabajo["_bloque_interv"] = bloque_interv_id
    df_trabajo["_fila_pos"] = range(len(df_trabajo))
    df_trabajo = df_trabajo[df_trabajo["_bloque_diag"] >= 1]

    casos: list[CasoEnfermeria] = []
    col_riesgo = columnas_clave["riesgo"]
    col_dominio = columnas_clave["dominio"]
    col_hallazgos = columnas_clave["hallazgos"]
    col_resultado = columnas_clave["resultado"]
    col_actividad = columnas_clave["actividad"]
    col_evaluacion = columnas_clave["evaluacion"]

    for _, grupo_diag in df_trabajo.groupby("_bloque_diag", sort=True):
        primera_pos = int(grupo_diag["_fila_pos"].iloc[0])

        diagnostico = CampoConCita(valor_para(col_diag, primera_pos), cita_para(col_diag, primera_pos))
        riesgo = CampoConCita(valor_para(col_riesgo, primera_pos), cita_para(col_riesgo, primera_pos))
        dominio = CampoConCita(valor_para(col_dominio, primera_pos), cita_para(col_dominio, primera_pos))
        resultado = CampoConCita(valor_para(col_resultado, primera_pos), cita_para(col_resultado, primera_pos))
        evaluacion = CampoConCita(valor_para(col_evaluacion, primera_pos), cita_para(col_evaluacion, primera_pos))

        hallazgos: list[Hallazgo] = []
        if col_hallazgos is not None:
            vistos = set()
            for _, fila in grupo_diag.iterrows():
                v = _texto_limpio(fila[col_hallazgos])
                if v and v not in vistos:
                    vistos.add(v)
                    hallazgos.append(Hallazgo(
                        valor=v,
                        cita=Cita(nombre_hoja, col_hallazgos, fila_base_excel + int(fila["_fila_pos"])),
                    ))

        intervenciones: list[Intervencion] = []
        for _, sub_grupo in grupo_diag.groupby("_bloque_interv", sort=True):
            pos_interv = int(sub_grupo["_fila_pos"].iloc[0])
            texto_interv = CampoConCita(
                valor_para(col_interv, pos_interv) if col_interv else None,
                cita_para(col_interv, pos_interv) if col_interv else None,
            )

            actividades: list[Actividad] = []
            if col_actividad is not None:
                vistas = set()
                for _, fila in sub_grupo.iterrows():
                    v = _texto_limpio(fila[col_actividad])
                    if v and v not in vistas:
                        vistas.add(v)
                        actividades.append(Actividad(
                            valor=v,
                            cita=Cita(nombre_hoja, col_actividad, fila_base_excel + int(fila["_fila_pos"])),
                        ))

            if texto_interv.valor or actividades:
                intervenciones.append(Intervencion(texto=texto_interv, actividades=actividades))

        casos.append(CasoEnfermeria(
            hoja=nombre_hoja,
            fila_excel_inicio=fila_base_excel + primera_pos,
            diagnostico=diagnostico,
            riesgo=riesgo,
            dominio=dominio,
            resultado=resultado,
            evaluacion=evaluacion,
            hallazgos=hallazgos,
            intervenciones=intervenciones,
        ))

    return casos, registros_planos, None


@st.cache_data(show_spinner=False)
def procesar_excel(archivo_bytes: bytes) -> tuple[list[CasoEnfermeria], pd.DataFrame, list[str]]:
    advertencias: list[str] = []
    todos_los_casos: list[CasoEnfermeria] = []
    todos_los_registros_planos: list[dict] = []

    excel_file = pd.ExcelFile(io.BytesIO(archivo_bytes), engine="openpyxl")

    for nombre_hoja in excel_file.sheet_names:
        try:
            df_crudo = excel_file.parse(sheet_name=nombre_hoja, header=None)
            casos, planos, aviso = _procesar_hoja(nombre_hoja, df_crudo)
            if aviso:
                advertencias.append(aviso)
            todos_los_casos.extend(casos)
            todos_los_registros_planos.extend(planos)
        except Exception as e:
            advertencias.append(f"Error procesando la hoja '{nombre_hoja}': {e}. Se omite esta hoja.")
            continue

    df_long = pd.DataFrame(
        todos_los_registros_planos, columns=["hoja", "fila_excel", "columna", "valor"]
    )
    return todos_los_casos, df_long, advertencias


# ---------------------------------------------------------------------------
# Renderizado UI
# ---------------------------------------------------------------------------

def _mostrar_campo(etiqueta: str, campo: CampoConCita) -> None:
    if campo.valor:
        cita_txt = f"  \n<sub>Cita: {campo.cita.texto()}</sub>" if campo.cita else ""
        st.markdown(f"**{etiqueta}:** {campo.valor}{cita_txt}", unsafe_allow_html=True)
    else:
        st.markdown(f"**{etiqueta}:** {PLACEHOLDER}")


def _renderizar_caso(caso: CasoEnfermeria, numero: int) -> None:
    st.markdown(f"##### Coincidencia {numero} — Diagnóstico: *{caso.diagnostico.valor or PLACEHOLDER}*")

    col_izq, col_der = st.columns(2)
    with col_izq:
        _mostrar_campo("Riesgo a prevenir", caso.riesgo)
        _mostrar_campo("Dominio", caso.dominio)
    with col_der:
        _mostrar_campo("Diagnóstico de Enfermería", caso.diagnostico)
        _mostrar_campo("Resultado esperado", caso.resultado)

    st.markdown("**Hallazgos:**")
    if caso.hallazgos:
        for h in caso.hallazgos:
            st.markdown(f"- {h.valor}  \n  <sub>Cita: {h.cita.texto()}</sub>", unsafe_allow_html=True)
    else:
        st.markdown(f"- {PLACEHOLDER}")

    _renderizar_intervenciones(caso.intervenciones)

    if caso.evaluacion.valor:
        st.markdown("**Evaluación del plan de cuidado:**")
        _mostrar_campo("Evaluación", caso.evaluacion)


def _grupos_por_caso(
    casos_hoja: list[CasoEnfermeria],
) -> list[tuple[str, list[CasoEnfermeria]]]:
    """Arma UNA cajita por cada caso (diagnóstico) de la hoja, con el título
    formado por sus propios hallazgos ya fusionados en un solo texto.

    FIX: antes se armaba una cajita POR CADA HALLAZGO. Cuando varios
    hallazgos apuntaban al mismo diagnóstico (ej. "Observación de facies de
    dolor", "Informe verbal de dolor..." y "Posturas para evitar el dolor"
    -> los 3 llevan a "Dolor"), esto producía 3-4 cajitas repetidas para el
    mismo diagnóstico. Ahora se arma una sola cajita por caso, con sus
    hallazgos fusionados en el título (separados por coma), tal como ya
    vienen agrupados en el Excel (mismo bloque de fila = mismo diagnóstico).

    Los pocos casos sin ningún hallazgo registrado se agrupan juntos al
    final bajo una única cajita "Sin hallazgo específico asociado", para no
    perderlos de vista.
    """
    grupos: list[tuple[str, list[CasoEnfermeria]]] = []
    casos_sin_hallazgo: list[CasoEnfermeria] = []

    for caso in casos_hoja:
        if not caso.hallazgos:
            casos_sin_hallazgo.append(caso)
            continue
        titulo = ", ".join(_normalizar_espacios(h.valor) for h in caso.hallazgos)
        if len(titulo) > 140:
            titulo = titulo[:137] + "..."
        grupos.append((titulo, [caso]))

    if casos_sin_hallazgo:
        grupos.append(("Sin hallazgo específico asociado", casos_sin_hallazgo))

    # Desambiguar por si dos cajitas distintas terminaran con el mismo título
    # (coincidencia rara, pero se evita mostrar dos cajitas idénticas).
    conteo: dict[str, int] = {}
    for titulo, _ in grupos:
        conteo[titulo.lower()] = conteo.get(titulo.lower(), 0) + 1

    grupos_finales: list[tuple[str, list[CasoEnfermeria]]] = []
    for titulo, lista_casos in grupos:
        if conteo[titulo.lower()] > 1:
            titulo = f"{titulo} — fila {lista_casos[0].fila_excel_inicio}"
        grupos_finales.append((titulo, lista_casos))

    return grupos_finales


def _renderizar_intervenciones(intervenciones: list[Intervencion]) -> None:
    """Muestra cada intervención como una cajita desplegable (st.expander)
    que, al abrirse, revela sus actividades. Se usa tanto en la vista
    guiada como en la búsqueda libre para mantener el mismo comportamiento."""
    st.markdown("**Intervenciones / Actividades:**")
    if not intervenciones:
        st.markdown(f"- {PLACEHOLDER}")
        return

    for interv in intervenciones:
        titulo = interv.texto.valor or PLACEHOLDER
        n_actividades = len(interv.actividades)
        sufijo_conteo = (
            f" · {n_actividades} actividad{'es' if n_actividades != 1 else ''}"
            if n_actividades else ""
        )
        with st.expander(f"🩹 {titulo}{sufijo_conteo}", expanded=False):
            if interv.texto.cita:
                st.markdown(f"<sub>Cita: {interv.texto.cita.texto()}</sub>", unsafe_allow_html=True)
            if interv.actividades:
                for act in interv.actividades:
                    st.markdown(
                        f"- {act.valor}  \n  <sub>Cita: {act.cita.texto()}</sub>",
                        unsafe_allow_html=True,
                    )
            else:
                st.markdown(PLACEHOLDER)


def _renderizar_resultado_e_intervenciones(caso: CasoEnfermeria) -> None:
    _mostrar_campo("Resultado esperado", caso.resultado)
    _renderizar_intervenciones(caso.intervenciones)


# ---------------------------------------------------------------------------
# Interfaz principal
# ---------------------------------------------------------------------------

def main() -> None:
    st.title("🩺 Buscador de Planes de Cuidados de Enfermería")
    st.caption(
        "Carga el Excel institucional más reciente. La búsqueda se hace "
        "sobre la columna 'Diagnóstico de Enfermería' de cada hoja."
    )

    archivo_subido = st.file_uploader(
        "Sube el archivo Excel (.xlsx / .xls)", type=["xlsx", "xls"], accept_multiple_files=False,
    )
    if archivo_subido is None:
        st.info("Esperando la carga de un archivo Excel para comenzar.")
        return

    try:
        contenido_bytes = archivo_subido.getvalue()
        casos, df_long, advertencias = procesar_excel(contenido_bytes)
    except ValueError as e:
        st.error(f"El archivo no pudo interpretarse como un Excel válido. Detalle técnico: {e}")
        return
    except Exception as e:
        st.error(f"Ocurrió un error inesperado al leer el archivo. Detalle: {e}")
        return

    if advertencias:
        with st.expander(f"⚠️ Avisos del procesamiento ({len(advertencias)})"):
            for a in advertencias:
                st.warning(a)

    if not casos:
        st.error(
            "No se pudo extraer ningún caso de enfermería del archivo. Revisa "
            "que las hojas tengan una columna de 'Diagnóstico de Enfermería'."
        )
        return

    st.success(f"Archivo procesado: {len(set(c.hoja for c in casos))} hoja(s), {len(casos)} caso(s) detectados.")

    st.markdown("### ¿Cómo quieres buscar?")
    modo = st.radio(
        "Modo de navegación",
        options=["Explorar por diagnóstico médico (guiado)", "Buscar directo por diagnóstico de enfermería"],
        label_visibility="collapsed",
    )
    st.divider()

    if modo.startswith("Explorar"):
        _vista_guiada_por_diagnostico_medico(casos)
    else:
        _vista_busqueda_libre(casos)

    with st.expander("Ver datos crudos de todas las hojas (tabla de citas completa)"):
        st.dataframe(
            df_long.rename(columns={
                "hoja": "Hoja", "columna": "Columna", "fila_excel": "Fila (Excel)", "valor": "Valor",
            }),
            use_container_width=True,
            hide_index=True,
        )


COLOR_INICIO = "#B37FE0"
COLOR_DX_MEDICO = "#1FAE6B"
COLOR_HALLAZGOS = "#FF3FA4"


def _inyectar_estilos_flujo() -> None:
    if st.session_state.get("_estilos_flujo_inyectados"):
        return
    st.session_state["_estilos_flujo_inyectados"] = True
    st.markdown(
        """
        <style>
        .pill-paso {
            display: inline-block;
            padding: 8px 18px;
            border-radius: 999px;
            color: white;
            font-weight: 700;
            margin: 6px 0 10px 0;
            font-size: 0.95rem;
            white-space: normal;
            overflow-wrap: break-word;
            max-width: 100%;
        }
        .caja-hallazgos {
            background: #FFE3F3;
            border-left: 6px solid #FF3FA4;
            border-radius: 10px;
            padding: 14px 18px;
            margin-bottom: 8px;
        }
        div[data-testid="stButton"] > button {
            white-space: normal !important;
            height: auto !important;
            min-height: 64px;
            padding: 12px 10px !important;
            border-radius: 14px !important;
            font-weight: 600 !important;
            border: 2px solid rgba(0,0,0,0.08) !important;
        }
        div[data-testid="stExpander"] {
            border: 1px solid rgba(0,0,0,0.08) !important;
            border-radius: 12px !important;
            margin-bottom: 8px;
            overflow: hidden;
        }
        div[data-testid="stExpander"] summary {
            font-weight: 600 !important;
            background: #F3F0FB;
        }
        """,
        unsafe_allow_html=True,
    )


def _pill(texto: str, color: str) -> None:
    st.markdown(f'<span class="pill-paso" style="background:{color};">{texto}</span>', unsafe_allow_html=True)


def _vista_guiada_por_diagnostico_medico(casos: list[CasoEnfermeria]) -> None:
    _inyectar_estilos_flujo()
    hojas_disponibles = sorted({c.hoja for c in casos}, key=_normalizar)

    _pill("🟣 Inicio", COLOR_INICIO)
    st.write("")
    _pill("1️⃣ Diagnóstico médico", COLOR_DX_MEDICO)
    # FIX: se usa index=None + placeholder en vez de meter "— Selecciona —"
    # como una opción más de la lista. Así el campo queda realmente vacío
    # (sin texto que borrar) y se puede escribir directo para filtrar.
    hoja_seleccionada = st.selectbox(
        "Selecciona la patología / diagnóstico médico:",
        options=hojas_disponibles,
        index=None,
        placeholder="Escribe o selecciona un diagnóstico médico...",
        label_visibility="collapsed",
        key="selector_dx_medico",
    )
    if hoja_seleccionada is None:
        st.caption("👆 Elige un diagnóstico médico para continuar.")
        return

    casos_hoja = [c for c in casos if c.hoja == hoja_seleccionada]

    _pill("2️⃣ Hallazgos → Diagnóstico de enfermería — despliega una cajita para ver su diagnóstico, resultado e intervenciones", COLOR_HALLAZGOS)

    grupos = _grupos_por_caso(casos_hoja)

    if not grupos:
        st.markdown(f'<div class="caja-hallazgos">{PLACEHOLDER}</div>', unsafe_allow_html=True)
        return

    # FIX: ya no hay botones ni "selección" con estado flotante. Cada cajita
    # es autocontenida: al abrirla ves hallazgos + diagnóstico + resultado
    # esperado + intervenciones, todo junto; al cerrarla se oculta todo por
    # igual. use `key=` para que Streamlit recuerde si estaba abierta o
    # cerrada entre una interacción y otra (por ejemplo, al abrir una
    # cajita de intervención dentro de ella).
    for i, (titulo, lista_casos) in enumerate(grupos):
        n_dx = len(lista_casos)
        sufijo = f"  ·  {n_dx} diagnósticos" if n_dx > 1 else ""
        with st.expander(f"🔎 {titulo}{sufijo}", expanded=False, key=f"caja_hallazgo::{hoja_seleccionada}::{i}"):
            for j, caso in enumerate(lista_casos):
                if j > 0:
                    st.divider()

                if caso.hallazgos:
                    for h in caso.hallazgos:
                        st.markdown(
                            f"- {h.valor}  \n  <sub>Cita: {h.cita.texto()}</sub>",
                            unsafe_allow_html=True,
                        )
                    st.write("")

                _pill(f"🩺 {caso.diagnostico.valor or PLACEHOLDER}", "#FF7FC0")
                if caso.diagnostico.cita:
                    st.markdown(f"<sub>Cita: {caso.diagnostico.cita.texto()}</sub>", unsafe_allow_html=True)
                _renderizar_resultado_e_intervenciones(caso)


def _vista_busqueda_libre(casos: list[CasoEnfermeria]) -> None:
    st.markdown("### 🔎 Buscador por diagnóstico")
    termino = st.text_input("Escribe un diagnóstico o palabra clave (ej. 'Dolor', 'Hepatitis C')", value="")

    termino_normalizado = _normalizar(termino)
    if not termino_normalizado:
        st.caption("Escribe un término para buscar en la columna 'Diagnóstico de Enfermería'.")
        return

    coincidencias = [
        c for c in casos
        if c.diagnostico.valor and termino_normalizado in _normalizar(c.diagnostico.valor)
    ]
    if not coincidencias:
        st.warning("No se encontraron coincidencias para ese término en ninguna hoja.")
        return

    st.markdown(f"**{len(coincidencias)} coincidencia(s) encontrada(s), agrupadas por hoja:**")
    st.divider()

    hojas_con_match = sorted({c.hoja for c in coincidencias})
    for hoja in hojas_con_match:
        casos_hoja = [c for c in coincidencias if c.hoja == hoja]
        st.markdown(f"### HOJA: {hoja.upper()}")
        for i, caso in enumerate(casos_hoja, start=1):
            _renderizar_caso(caso, i)
        st.divider()


if __name__ == "__main__":
    main()