"""ENESEM Servicios: revisión por Excel con cinco cargas independientes."""

from collections import defaultdict, Counter
from hashlib import sha256

import pandas as pd
import streamlit as st

from engine import ciiu_for_code, get_catalog, norm, summarize_company
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
        "Carga": r["load"], "Empresa": r["company_id"],
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
RULE_HELP = {
    "Formato": "El código informado debe tener ocho dígitos: cuatro de CIIU y cuatro de CPC.",
    "Código Otros": "99999999 está reservado para la fila Otros; no debe codificar otro servicio.",
    "CIIU código": "Los cuatro primeros dígitos del código no figuran entre los CIIU de servicios o excepción del catálogo.",
    "CPC código": "Los cuatro últimos dígitos del código no figuran entre los CPC de servicios o excepción del catálogo.",
    "Par CIIU–CPC": "La combinación de los ocho dígitos no aparece en la tabla de correspondencias; requiere revisión.",
    "Actividad servicio": "Ninguna de las tres actividades de la empresa se reconoce como servicio o actividad de excepción.",
    "División excluida": "Hay una actividad en las divisiones K64, K66, Q88 o S94, que requieren revisar su tratamiento.",
    "CIIU actividad": "Los cuatro dígitos CIIU del servicio no coinciden con los de ninguna actividad informada por la empresa.",
}


def alert_counts(items):
    return {name: sum(any(a.startswith(prefix) for a in r["validation"]["alerts"])
                      for r in items) for name, prefix in CHECKS.items()}


def summary_row(label, rows):
    checks = alert_counts(rows)
    return {"Grupo": label, "Servicios": len(rows),
            "Revisados": sum(r["decision"] in ("Justificación", "Cambio") for r in rows),
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


def company_profile(company, rows, catalog):
    """Datos informados por la empresa, sin tomar valores de la revisión."""
    sources = [r["source"] for r in rows]
    def field(name):
        return next((norm(source.get(name)) for source in sources
                     if norm(source.get(name))), "")
    st.subheader(f"Información de la empresa {company}")
    with st.container(border=True):
        a, b, c, d = st.columns(4)
        a.write("**RUC**")
        a.write(field("ruc") or "—")
        b.write("**Nombre comercial**")
        b.write(field("nombre_comercial") or "—")
        c.write("**Razón social**")
        c.write(field("razon_social") or "—")
        d.write("**Zonal**")
        d.write(field("inec_zonal") or "—")
        st.write("**Actividades y CIIU informados**")
        activities = (
            ("Principal", "actividad_principal", "ciiu4_actividad_principal", "desc_actividad_principal"),
            ("Secundaria", "actividad_secundaria", "ciiu4_actividad_secundaria", "desc_actividad_secundaria"),
            ("Tercera", "tercera_actividad", "ciiu4_tercera_actividad", "desc_tercera_actividad"),
        )
        for column, (label, activity, ciiu, description) in zip(st.columns(3), activities):
            with column:
                st.write(f"**{label}**")
                st.write(f"CIIU: {field(ciiu) or '—'}")
                if field(description):
                    st.caption(field(description))
                elif field(activity):
                    st.caption(field(activity))
        with st.expander("Consultar CPC posibles según las actividades de esta empresa"):
            st.caption("Correspondencias del catálogo para cada CIIU completo. Son opciones de referencia; la descripción del servicio informado determina la codificación.")
            for label, activity, ciiu, description in activities:
                full_activity = field(ciiu).upper().replace(".", "")
                if not full_activity:
                    continue
                possible = catalog.activity_pairs.get(full_activity, ())
                st.write(f"**{label}: {full_activity} · {len(possible)} código(s)**")
                if possible:
                    st.dataframe(pd.DataFrame([{
                        "Código CIIU–CPC": code,
                        "CPC4": code[4:],
                        "Descripción del CPC": desc or catalog.cpc_names.get(code[4:], ""),
                    } for code, desc in possible]), hide_index=True,
                        use_container_width=True, height=min(260, 38 * (len(possible) + 1)))
                else:
                    st.caption("Sin correspondencias para este CIIU completo en el catálogo.")


def reverse_lookup(items, catalog, key):
    with st.expander("Consultar actividades CIIU de un código registrado"):
        st.caption("Selecciona un servicio para ver los CIIU asociados a su código de ocho dígitos.")
        choices = pd.DataFrame([{"Empresa": r["company_id"],
            "Servicio informado": norm(r["source"].get("Descripción del servicio")),
            "Código original": norm(r["source"].get("Código CPC")),
            "Código vigente": r["validation"]["code"]} for r in items])
        event = st.dataframe(choices, hide_index=True, use_container_width=True,
            height=min(250, 88 + 36 * len(items)), on_select="rerun",
            selection_mode="single-row", key=f"reverse_pick_{key}")
        if not event.selection.rows:
            st.info("Selecciona una fila para consultar el código.")
            return
        item = items[event.selection.rows[0]]
        code = item["validation"]["code"]
        original = norm(item["source"].get("Código CPC"))
        if original != code:
            code = st.radio("Código a consultar", [original, code], horizontal=True,
                            key=f"reverse_code_{key}_{item['service_id']}")
        broader = st.checkbox("Mostrar también otros CIIU del mismo CPC4", key=f"reverse_broad_{key}")
        matches = ciiu_for_code(catalog, code, include_same_cpc=broader)
        if not matches:
            st.warning("El catálogo no contiene correspondencias para este código.")
            return
        source = item["source"]
        own = {
            norm(source.get("ciiu4_actividad_principal")).upper().replace(".", ""): "Principal",
            norm(source.get("ciiu4_actividad_secundaria")).upper().replace(".", ""): "Secundaria",
            norm(source.get("ciiu4_tercera_actividad")).upper().replace(".", ""): "Tercera",
        }
        st.dataframe(pd.DataFrame([{"CIIU": activity, "Descripción de actividad": description,
            "Código CIIU–CPC": combined,
            "En esta empresa": own.get(activity, "No consta")}
            for activity, description, combined in matches]), hide_index=True,
            use_container_width=True, height=min(350, 88 + 36 * len(matches)))


def edit_matrix(items: list[dict], state: dict, catalog, key: str,
                show_activities: bool = False, show_identity: bool = False):
    """Revisa varios servicios en una cuadrícula; los datos del informante son de lectura."""
    rows = []
    for r in items:
        source = r["source"]
        row = {"Empresa": r["company_id"]}
        if show_identity:
            row.update({"RUC": norm(source.get("ruc")),
                        "Nombre comercial": norm(source.get("nombre_comercial")),
                        "Razón social": norm(source.get("razon_social"))})
        if show_activities:
            row.update({"CIIU principal": norm(source.get("ciiu4_actividad_principal")),
                        "CIIU secundaria": norm(source.get("ciiu4_actividad_secundaria")),
                        "CIIU tercera": norm(source.get("ciiu4_tercera_actividad"))})
        row.update({
            "Descripción": norm(r["source"].get("Descripción del servicio")),
            "CPC original": norm(r["source"].get("Código CPC")),
            "CPC corregido": r["edited_code"],
            "Descripción del CPC vigente": r["validation"]["cpc_name"],
            "Alertas": " | ".join(r["validation"]["alerts"]),
            "Decisión": r["decision"], "Responsable": r["reviewer"],
            "Revisar": r["review_note"], "Observaciones": r["observations"],
        })
        rows.append(row)
    st.caption("Edita las celdas blancas y pulsa Guardar cambios. El CPC original y los datos informados son de lectura.")
    frame = pd.DataFrame(rows)
    disabled = ["Empresa", "Descripción", "CPC original", "Descripción del CPC vigente", "Alertas",
                "RUC", "Nombre comercial", "Razón social", "CIIU principal", "CIIU secundaria", "CIIU tercera"]
    edited = st.data_editor(frame, key=f"grid_{key}", hide_index=True,
        use_container_width=True, height=min(560, 88 + 36 * len(rows)),
        disabled=[column for column in disabled if column in frame.columns], num_rows="fixed",
        column_config={
            "Decisión": st.column_config.SelectboxColumn(options=["Pendiente", "Justificación", "Cambio"], required=True),
            "CPC corregido": st.column_config.TextColumn(help="Código de ocho dígitos, sin modificar el original"),
            "Descripción": st.column_config.TextColumn(width="large"),
            "Descripción del CPC vigente": st.column_config.TextColumn(width="large", help="Descripción del CPC del código vigente. Se actualiza después de guardar una corrección."),
            "Alertas": st.column_config.TextColumn(width="large"),
            "Revisar": st.column_config.TextColumn(width="medium"),
            "Observaciones": st.column_config.TextColumn(width="medium"),
        })
    with st.expander("Justificaciones y revisión PC (opcionales)"):
        extras = pd.DataFrame([{
            "Servicio": norm(r["source"].get("Descripción del servicio")),
            "Justificación CIIU": r["just_ciiu"],
            "Justificación CPC": r["just_cpc"],
            "Justificación correspondencia": r["just_match"],
            "Revisión PC": r["review_pc"],
        } for r in items])
        extra_edit = st.data_editor(extras, key=f"extras_{key}", hide_index=True,
            use_container_width=True, height=min(420, 88 + 36 * len(rows)),
            disabled=["Servicio"], num_rows="fixed")
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
    reverse_lookup(items, catalog, key)


st.markdown("""<style>
.enesem-hero {background:linear-gradient(105deg,#19395f,#2f6ee8);color:white;
  padding:24px 30px;border-radius:16px;margin:0 0 18px 0}
.enesem-hero h1 {color:white;font-size:2.1rem;line-height:1.2;margin:0 0 9px}
.enesem-hero p {color:#edf4ff;margin:0}
div[data-testid="stMetric"] {background:#eaf2ff;border:1px solid #d6e5ff;
  border-radius:12px;padding:12px 16px}
</style><div class="enesem-hero"><h1>ENESEM · Tomo II · Servicios</h1>
<p>Revisión por empresa, código CPC y servicio · cinco cargas · avances en Excel</p></div>""",
    unsafe_allow_html=True)
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
    c3.metric("Revisados", sum(r["decision"] in ("Justificación", "Cambio") for r in visible))
    c4.metric("Con alertas", sum(bool(r["validation"]["alerts"]) for r in visible))
    c5.metric("Incumplimientos totales", sum(len(r["validation"]["alerts"]) for r in visible),
              help="Suma de alertas automáticas; un servicio puede incumplir varias reglas.")
    with st.expander("Alertas por regla"):
        st.caption("Cantidad de servicios de la carga que incumplen cada validación; se excluyen las filas Otros. Coloca el cursor sobre ⓘ para ver la regla.")
        counts = alert_counts(visible)
        for start in range(0, len(CHECKS), 4):
            for column, (name, count) in zip(st.columns(4), list(counts.items())[start:start + 4]):
                column.metric(name, count, help=RULE_HELP[name])

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
                company_profile(company, groups[company], catalog)
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
                f1, f2 = st.columns(2)
                only_cpc_alerts = f1.checkbox("Solo con alertas en este CPC",
                    key=f"only_cpc_alerts_{generation}_{code}")
                show_identity = f2.checkbox("Mostrar RUC y nombres de empresa",
                    key=f"show_identity_{generation}_{code}")
                code_items = [r for r in by_cpc[code]
                              if not only_cpc_alerts or r["validation"]["alerts"]]
                st.caption(f"{len(code_items)} de {len(by_cpc[code])} servicios visibles")
                if code_items:
                    edit_matrix(code_items, state, catalog,
                        f"cpc_{generation}_{code}_{only_cpc_alerts}_{show_identity}",
                        show_activities=True, show_identity=show_identity)
                else:
                    st.info("Este código no tiene servicios con alertas.")
            else:
                st.info("Selecciona una fila de la matriz para revisar ese código CPC.")
        else:
            st.info("No hay códigos con ese filtro.")
    with tab_s:
        q1, q2, q3 = st.columns(3)
        query = q1.text_input("Empresa o servicio", key=f"query_{generation}")
        alert_filter = q2.selectbox("Alertas", ["Todas", "Con alertas", "Sin alertas"], key=f"alerts_{generation}")
        decision_filter = q3.selectbox("Decisión", ["Todas", "Pendiente", "Justificación", "Cambio"], key=f"decision_filter_{generation}")
        selection = [r for r in visible if
            (not query or query.casefold() in (r["company_id"] + " " + norm(r["source"].get("Descripción del servicio"))).casefold())
            and (alert_filter == "Todas" or bool(r["validation"]["alerts"]) == (alert_filter == "Con alertas"))
            and (decision_filter == "Todas" or r["decision"] == decision_filter)]
        st.dataframe(table(selection), hide_index=True, use_container_width=True, height=380)
        if selection:
            service = st.selectbox("Servicio para revisar", selection, key=f"service_{generation}",
                format_func=lambda r: f"{r['company_id']} · {r['validation']['code']} · {norm(r['source'].get('Descripción del servicio'))[:65]}")
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
