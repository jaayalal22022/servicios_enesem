"""ENESEM Servicios: revisión por Excel con cinco cargas independientes."""

from collections import defaultdict, Counter

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


def picker(items: list[dict], key: str) -> dict:
    return st.selectbox("Elegir servicio", items, key=key,
        format_func=lambda r: (f"{r['company_id']} · {r['service_id']} · "
                               f"{norm(r['source'].get('Descripción del servicio'))[:85]}"))


def review_form(item: dict, state: dict, catalog, key: str):
    source, result = item["source"], item["validation"]
    st.caption(f"{item['load']} · Empresa {item['company_id']} · id_serv {item['service_id']}")
    st.write("**Actividades:**",
             "principal", norm(source.get("ciiu4_actividad_principal")) or "—", "·",
             "secundaria", norm(source.get("ciiu4_actividad_secundaria")) or "—", "·",
             "tercera", norm(source.get("ciiu4_tercera_actividad")) or "—")
    c1, c2, c3 = st.columns(3)
    c1.metric("CPC original", norm(source.get("Código CPC")))
    c2.metric("Código vigente", result["code"])
    c3.metric("Correspondencia con actividad", "Sí" if result["activity_match"] else "No")
    st.caption(f"CPC4 {result['cpc4'] or '—'} · {result['cpc_name'] or '—'}")
    if result["exceptions"]:
        st.info("Actividad de excepción: " + ", ".join(result["exceptions"]))
    for alert in result["alerts"]:
        st.warning(alert)
    if not result["alerts"]:
        st.success("Sin alertas automáticas.")
    prefix = f"{key}_{item['company_id']}_{item['service_id']}"
    with st.form(prefix):
        code = st.text_input("Código CPC corregido (8 dígitos)", value=item["edited_code"])
        decision = st.selectbox("Decisión", ["Pendiente", "Aceptado", "Cambio"],
            index=["Pendiente", "Aceptado", "Cambio"].index(item["decision"])
            if item["decision"] in ("Pendiente", "Aceptado", "Cambio") else 0)
        reviewer = st.text_input("Nombre del responsable",
                                 value=item["reviewer"] or st.session_state.get("my_name", ""))
        note = st.text_area("Revisar / acción necesaria", value=item["review_note"])
        a, b = st.columns(2)
        just_ciiu = a.text_area("Justificación CIIU", value=item["just_ciiu"])
        just_cpc = b.text_area("Justificación CPC", value=item["just_cpc"])
        just_match = st.text_area("Justificación de correspondencia", value=item["just_match"])
        observations = st.text_area("Observaciones", value=item["observations"])
        review_pc = st.text_input("Revisión PC", value=item["review_pc"])
        submitted = st.form_submit_button("Guardar en esta sesión", type="primary")
    if submitted:
        try:
            changed = update_review(state, identity(item), {
                "edited_code": code, "review_note": note, "reviewer": reviewer,
                "decision": decision, "just_ciiu": just_ciiu, "just_cpc": just_cpc,
                "just_match": just_match, "observations": observations,
                "review_pc": review_pc,
            }, catalog)
        except ValueError as exc:
            st.error(str(exc))
            return
        if changed:
            st.session_state.pop("download_bytes", None)
            st.session_state["notice"] = "Revisión guardada en esta sesión. Descarga el Excel para conservarla."
            st.rerun()
        st.info("No hubo cambios en esta revisión.")


st.title("ENESEM · Tomo II · Servicios")
st.caption("Trabajo con Excel · revisión por servicio, empresa y CPC · cinco cargas")
catalog = get_catalog()

with st.sidebar:
    st.header("Archivos de trabajo")
    uploaded = st.file_uploader("Cargar base o avance (.xlsx)", type=["xlsx"], key="source_upload")
    if st.button("Abrir archivo", disabled=uploaded is None):
        try:
            candidate = load_state(uploaded.getvalue())
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
    selected_load = st.selectbox("Carga que revisarás", ["Todas (coordinación)", *LOADS],
                                  index=1, key=f"load_{generation}")
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
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Servicios en la carga", len(visible))
    c2.metric("Empresas", len(groups))
    c3.metric("Revisados", sum(r["decision"] in ("Aceptado", "Cambio") for r in visible))
    c4.metric("Con alertas", sum(bool(r["validation"]["alerts"]) for r in visible))

    tab_s, tab_e, tab_c = st.tabs(["Por servicio", "Por empresa", "Por código CPC"])
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
            review_form(picker(selection, f"service_{generation}"), state, catalog, f"service_{generation}")
    with tab_e:
        company = st.selectbox("Empresa", sorted(groups), key=f"company_{generation}")
        info = summary[company]
        st.write(f"Correspondencias: **{info['matches']} de {info['total']}** · "
                 f"mínimo: **{'Cumple' if info['minimum_met'] else 'No cumple'}** · "
                 f"revisados: **{info['reviewed']}**")
        st.dataframe(table(groups[company]), hide_index=True, use_container_width=True)
        review_form(picker(groups[company], f"company_record_{generation}"), state, catalog, f"company_{generation}")
    with tab_c:
        by_cpc = defaultdict(list)
        for row in visible:
            by_cpc[row["validation"]["cpc4"] or "Sin CPC válido"].append(row)
        st.dataframe(pd.DataFrame([{"CPC4": code,
            "Descripción": catalog.cpc_names.get(code, "Sin descripción"),
            "Servicios": len(rows), "Empresas": len({r['company_id'] for r in rows}),
            "Con alertas": sum(bool(r["validation"]["alerts"]) for r in rows)}
            for code, rows in sorted(by_cpc.items())]), hide_index=True, use_container_width=True)
        code = st.selectbox("Código CPC4", sorted(by_cpc), key=f"cpc_{generation}")
        st.dataframe(table(by_cpc[code]), hide_index=True, use_container_width=True)
        review_form(picker(by_cpc[code], f"cpc_record_{generation}"), state, catalog, f"cpc_{generation}")

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
