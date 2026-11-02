"""ENESEM Servicios: revisión por Excel con cinco cargas independientes."""

from collections import defaultdict, Counter
from hashlib import sha256

import pandas as pd
import streamlit as st

from engine import evaluate, get_catalog, is_other, norm, summarize_company
from workflow import (LOADS, MergeConflict, assign_company, export_state, identity,
                      load_state, merge_states, reviewed_rows, update_review)


st.set_page_config(page_title="ENESEM · Servicios · Excel", page_icon="🔎", layout="wide")


def install_state(state: dict, message: str):
    st.session_state["workbook_state"] = state
    st.session_state["generation"] = st.session_state.get("generation", 0) + 1
    st.session_state.pop("download_bytes", None)
    st.session_state.pop("preferred_company", None)
    st.session_state.pop("preferred_cpc", None)
    st.session_state["notice"] = message
    st.rerun()


def table(items: list[dict]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Carga": r["load"], "Empresa": r["company_id"], "id_serv": r["service_id"],
        "Servicio": norm(r["source"].get("Descripción del servicio")),
        "CPC original": norm(r["source"].get("Código CPC")),
        "CPC vigente": r["validation"]["code"], "CPC4": r["validation"]["cpc4"],
        "CIIU coincide": "Sí" if r["validation"]["activity_match"] else "No",
        "Alertas": " | ".join(r["validation"]["alerts"]) or "Sin alertas",
        "Decisión": r["decision"], "Responsable": r["reviewer"],
    } for r in items])


CHECKS = {
    "Formato": "Código de servicio:",
    "Código Otros": "99999999 se reserva",
    "CIIU código": "CIIU del código fuera",
    "CPC código": "CPC fuera",
    "Par CIIU–CPC": "Combinación CIIU–CPC",
    "Actividad servicio": "Ninguna actividad de servicio",
    "División excluida": "Actividad en división no investigada",
    "CIIU actividad": "CIIU del servicio no coincide",
}


def alert_counts(items):
    return {name: sum(any(a.startswith(prefix) for a in r["validation"]["alerts"])
                      for r in items) for name, prefix in CHECKS.items()}


def summary_row(label, rows):
    checks = alert_counts(rows)
    return {"Grupo": label, "Servicios": len(rows),
            "Revisados": sum(r["decision"] in ("Aceptado", "Cambio") for r in rows),
            "Con alertas": sum(bool(r["validation"]["alerts"]) for r in rows),
            "Incumplimientos": sum(len(r["validation"]["alerts"]) for r in rows),
            **checks}


def company_summary(company, rows):
    source_rows = [r["source"] for r in rows]
    def first_value(field):
        return next((norm(source.get(field)) for source in source_rows
                     if norm(source.get(field))), "")
    return {"id_empresa": company, "ruc": first_value("ruc"),
            "nombre_comercial": first_value("nombre_comercial"),
            "razon_social": first_value("razon_social"),
            **{k: v for k, v in summary_row(company, rows).items() if k != "Grupo"}}


def edit_matrix(items: list[dict], state: dict, catalog, key: str):
    """Revisa varios servicios en una cuadrícula; los datos del informante son de lectura."""
    rows = []
    for r in items:
        rows.append({
            "Empresa": r["company_id"], "id_serv": r["service_id"],
            "Descripción": norm(r["source"].get("Descripción del servicio")),
            "CPC original": norm(r["source"].get("Código CPC")),
            "CPC corregido": r["edited_code"], "Alertas": " | ".join(r["validation"]["alerts"]),
            "Decisión": r["decision"], "Responsable": r["reviewer"],
            "Revisar": r["review_note"], "Observaciones": r["observations"],
        })
    st.caption("Edita las celdas blancas y pulsa Guardar cambios. El CPC original y los datos informados son de lectura.")
    frame = pd.DataFrame(rows)
    disabled = ["Empresa", "id_serv", "Descripción", "CPC original", "Alertas"]
    edited = st.data_editor(frame, key=f"grid_{key}", hide_index=True,
        use_container_width=True, height=min(560, 88 + 36 * len(rows)),
        disabled=disabled, num_rows="fixed",
        column_config={
            "Decisión": st.column_config.SelectboxColumn(options=["Pendiente", "Aceptado", "Cambio"], required=True),
            "CPC corregido": st.column_config.TextColumn(help="Código de ocho dígitos, sin modificar el original"),
            "Descripción": st.column_config.TextColumn(width="large"),
            "Alertas": st.column_config.TextColumn(width="large"),
            "Revisar": st.column_config.TextColumn(width="medium"),
            "Observaciones": st.column_config.TextColumn(width="medium"),
        })
    with st.expander("Justificaciones y revisión PC (opcionales)"):
        extras = pd.DataFrame([{
            "id_serv": r["service_id"], "Justificación CIIU": r["just_ciiu"],
            "Justificación CPC": r["just_cpc"],
            "Justificación correspondencia": r["just_match"],
            "Revisión PC": r["review_pc"],
        } for r in items])
        extra_edit = st.data_editor(extras, key=f"extras_{key}", hide_index=True,
            use_container_width=True, height=min(420, 88 + 36 * len(rows)),
            disabled=["id_serv"], num_rows="fixed")
    if st.button("Guardar cambios de la matriz", key=f"save_{key}", type="primary"):
        failures, saved = [], 0
        for idx, item in enumerate(items):
            main, extra = edited.iloc[idx], extra_edit.iloc[idx]
            changes = {
                "edited_code": main["CPC corregido"], "decision": main["Decisión"],
                "reviewer": main["Responsable"], "review_note": main["Revisar"],
                "observations": main["Observaciones"],
                "just_ciiu": extra["Justificación CIIU"], "just_cpc": extra["Justificación CPC"],
                "just_match": extra["Justificación correspondencia"], "review_pc": extra["Revisión PC"],
            }
            if all(norm(value) == norm(item[field]) for field, value in changes.items()):
                continue
            changes["reviewer"] = norm(changes["reviewer"]) or st.session_state.get("my_name", "")
            try:
                if update_review(state, identity(item), changes, catalog):
                    saved += 1
            except ValueError as exc:
                failures.append(f"id_serv {item['service_id']}: {exc}")
        if saved:
            st.session_state.pop("download_bytes", None)
            st.session_state["preferred_load"] = selected_load
            st.session_state["generation"] += 1
            st.session_state["notice"] = f"{saved} revisión(es) guardada(s) en esta sesión. Descarga el Excel para conservarlas."
        if failures:
            st.session_state["save_errors"] = failures
        if saved:
            st.rerun()
        if not failures:
            st.info("No hubo cambios para guardar.")


st.title("ENESEM · Tomo II · Servicios")
st.caption("Trabajo con Excel · revisión por empresa, CPC y servicio · cinco cargas")
catalog = get_catalog()

with st.sidebar:
    st.header("Archivos de trabajo")
    uploaded = st.file_uploader("Cargar base o avance (.xlsx)", type=["xlsx"], key="source_upload")
    if uploaded is not None:
        try:
            content = uploaded.getvalue()
            signature = sha256(content).hexdigest()
            if signature != st.session_state.get("uploaded_signature"):
                candidate = load_state(content)
                st.session_state["uploaded_signature"] = signature
                install_state(candidate, "Archivo cargado. Los avances previos y las cargas se restauraron.")
        except Exception as exc:
            st.error(f"No se pudo cargar: {exc}")
    st.divider()
    st.write("**Consolidar revisiones paralelas**")
    merge_files = st.file_uploader("Avances de los revisores (2 a 5)",
                                   type=["xlsx"], accept_multiple_files=True, key="merge_upload")
    if st.button("Consolidar archivos", disabled=len(merge_files) < 2):
        try:
            if len(merge_files) > 5:
                raise ValueError("Carga un máximo de cinco archivos por vez.")
            merged = merge_states([load_state(f.getvalue()) for f in merge_files])
            install_state(merged, "Avances consolidados. Descarga el Excel resultante.")
        except (ValueError, MergeConflict) as exc:
            st.error(str(exc))

state = st.session_state.get("workbook_state")
notice = st.session_state.pop("notice", None)
if notice:
    st.success(notice)
for error in st.session_state.pop("save_errors", []):
    st.error(error)
if state is None:
    st.info("Carga la base para distribuir empresas entre cinco cargas. Descarga el Excel maestro y entrega una copia a cada revisor. Cada persona selecciona su carga, revisa y descarga su avance. Finalmente, consolida los cinco archivos aquí.")
    st.stop()

for warning in state["warnings"]:
    st.warning(warning)
generation = st.session_state.get("generation", 0)
all_items = reviewed_rows(state, catalog)
active = [r for r in all_items if not r["validation"]["other"]]

with st.sidebar:
    st.text_input("Tu nombre", key="my_name")
    load_options = ["Todas (coordinación)", *LOADS]
    preferred_load = st.session_state.get("preferred_load", "Carga 1")
    selected_load = st.selectbox("Carga que revisarás", load_options,
                                  index=load_options.index(preferred_load) if preferred_load in load_options else 1,
                                  key=f"load_{generation}")
    totals = Counter(r["load"] for r in active)
    st.dataframe(pd.DataFrame([{"Carga": x, "Servicios": totals[x],
                               "Empresas": len({r['company_id'] for r in active if r['load'] == x})}
                              for x in LOADS]), hide_index=True, use_container_width=True)
    with st.expander("Reasignar empresa antes de iniciar"):
        company_move = st.selectbox("Empresa", sorted({r["company_id"] for r in all_items}),
                                    key=f"move_company_{generation}")
        new_load = st.selectbox("Nueva carga", list(LOADS), key=f"move_load_{generation}")
        if st.button("Asignar empresa"):
            try:
                assign_company(state, company_move, new_load)
                st.session_state.pop("download_bytes", None)
                st.session_state["notice"] = f"Empresa {company_move} asignada a {new_load}. Descarga el Excel para distribuirlo."
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

visible = active if selected_load.startswith("Todas") else [r for r in active if r["load"] == selected_load]
if not visible:
    st.info("Esta carga no tiene servicios evaluables.")
else:
    groups = defaultdict(list)
    for row in visible:
        groups[row["company_id"]].append(row)
    summary = {company: summarize_company(rows) for company, rows in groups.items()}
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Servicios en la carga", len(visible))
    c2.metric("Empresas", len(groups))
    c3.metric("Revisados", sum(r["decision"] in ("Aceptado", "Cambio") for r in visible))
    c4.metric("Con alertas", sum(bool(r["validation"]["alerts"]) for r in visible))
    c5.metric("Incumplimientos totales", sum(len(r["validation"]["alerts"]) for r in visible),
              help="Suma de alertas automáticas; un servicio puede incumplir varias reglas.")

    tab_e, tab_c, tab_s = st.tabs(["Por empresa", "Por código CPC", "Por servicio"])
    with tab_e:
        st.subheader("Resumen de empresas")
        st.caption("Cada columna de validación cuenta servicios que la incumplen. Incumplimientos suma todas las alertas.")
        q_company, only_alerts = st.columns([3, 1])
        company_query = q_company.text_input("Buscar por ID, RUC o nombre", key=f"company_query_{generation}")
        company_alerts = only_alerts.checkbox("Solo con alertas", key=f"company_alerts_{generation}")
        company_rows = [company_summary(company, rows) for company, rows in sorted(groups.items())]
        if company_query:
            company_rows = [r for r in company_rows if company_query.casefold() in
                " ".join(str(r[k]) for k in ("id_empresa", "ruc", "nombre_comercial", "razon_social")).casefold()]
        if company_alerts:
            company_rows = [r for r in company_rows if r["Con alertas"]]
        if company_rows:
            company_event = st.dataframe(pd.DataFrame(company_rows), hide_index=True,
                use_container_width=True, height=420, on_select="rerun", selection_mode="single-row",
                key=f"company_summary_{generation}_{selected_load}_{company_query}_{company_alerts}")
            selected = company_event.selection.rows
            if selected:
                st.session_state["preferred_company"] = company_rows[selected[0]]["id_empresa"]
            company = (company_rows[selected[0]]["id_empresa"] if selected else
                       st.session_state.get("preferred_company"))
            if company in {r["id_empresa"] for r in company_rows}:
                st.subheader(f"Empresa {company}")
                info = summary[company]
                st.write(f"Correspondencias: **{info['matches']} de {info['total']}** · "
                         f"mínimo: **{'Cumple' if info['minimum_met'] else 'No cumple'}** · "
                         f"revisados: **{info['reviewed']}**")
                edit_matrix(groups[company], state, catalog, f"company_{generation}_{company}")
            else:
                st.info("Selecciona una fila de la matriz para revisar esa empresa.")
        else:
            st.info("No hay empresas con esos filtros.")
    with tab_c:
        by_cpc = defaultdict(list)
        for row in visible:
            by_cpc[row["validation"]["cpc4"] or "Sin CPC válido"].append(row)
        st.subheader("Resumen por código CPC")
        cpc_rows = [{
            "CPC4": code, "Descripción": catalog.cpc_names.get(code, "Sin descripción"),
            "Empresas": len({r['company_id'] for r in rows}),
            **{k: v for k, v in summary_row(code, rows).items() if k != "Grupo"},
        } for code, rows in sorted(by_cpc.items())]
        cpc_query = st.text_input("Buscar código o descripción", key=f"cpc_query_{generation}")
        if cpc_query:
            cpc_rows = [r for r in cpc_rows if cpc_query.casefold() in
                        (r["CPC4"] + " " + r["Descripción"]).casefold()]
        if cpc_rows:
            cpc_event = st.dataframe(pd.DataFrame(cpc_rows), hide_index=True,
                use_container_width=True, height=420, on_select="rerun", selection_mode="single-row",
                key=f"cpc_summary_{generation}_{selected_load}_{cpc_query}")
            selected = cpc_event.selection.rows
            if selected:
                st.session_state["preferred_cpc"] = cpc_rows[selected[0]]["CPC4"]
            code = (cpc_rows[selected[0]]["CPC4"] if selected else
                    st.session_state.get("preferred_cpc"))
            if code in {r["CPC4"] for r in cpc_rows}:
                st.subheader(f"CPC {code}")
                st.caption(catalog.cpc_names.get(code, "Sin descripción"))
                edit_matrix(by_cpc[code], state, catalog, f"cpc_{generation}_{code}")
            else:
                st.info("Selecciona una fila de la matriz para revisar ese código CPC.")
        else:
            st.info("No hay códigos con ese filtro.")
    with tab_s:
        q1, q2, q3 = st.columns(3)
        query = q1.text_input("Empresa o servicio", key=f"query_{generation}")
        alert_filter = q2.selectbox("Alertas", ["Todas", "Con alertas", "Sin alertas"], key=f"alerts_{generation}")
        decision_filter = q3.selectbox("Decisión", ["Todas", "Pendiente", "Aceptado", "Cambio"], key=f"decision_filter_{generation}")
        selection = [r for r in visible if
            (not query or query.casefold() in (r["company_id"] + " " + norm(r["source"].get("Descripción del servicio"))).casefold())
            and (alert_filter == "Todas" or bool(r["validation"]["alerts"]) == (alert_filter == "Con alertas"))
            and (decision_filter == "Todas" or r["decision"] == decision_filter)]
        st.dataframe(table(selection), hide_index=True, use_container_width=True, height=380)
        if selection:
            service = st.selectbox("Servicio para revisar", selection, key=f"service_{generation}",
                format_func=lambda r: f"{r['company_id']} · {r['service_id']} · {norm(r['source'].get('Descripción del servicio'))[:65]}")
            edit_matrix([service], state, catalog, f"service_{generation}_{service['company_id']}_{service['service_id']}")

st.divider()
st.subheader("Descargar avance")
st.write("Descarga el archivo después de revisar. Contiene la base completa, la carga asignada, validaciones recalculadas y el historial. Al volverlo a cargar podrás continuar.")
if st.button("Preparar Excel de avance"):
    st.session_state["download_bytes"] = export_state(state, catalog)
if st.session_state.get("download_bytes"):
    st.download_button("Descargar Excel", data=st.session_state["download_bytes"],
                       file_name="ENESEM_Servicios_avance.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("El guardado en esta versión está en el Excel descargado. Si cierras la sesión antes de descargar, los cambios recientes se perderán.")
