"""Estado de revisión transportado en Excel, sin base de datos."""

from __future__ import annotations

from collections import Counter
from datetime import datetime, timezone
from hashlib import sha256
from io import BytesIO
import json
from uuid import uuid4

from openpyxl import load_workbook

from engine import as_json, evaluate, is_other, norm, parse_upload, review_defaults
from exporter import export_excel


LOADS = tuple(f"Carga {n}" for n in range(1, 6))
REVIEW_KEYS = (
    "edited_code", "review_note", "reviewer", "decision", "just_ciiu",
    "just_cpc", "just_match", "observations", "review_pc",
)
REVIEW_COLUMNS = {
    "cpc_act_8dígitos": "edited_code", "Revisar": "review_note",
    "Nombre del responsable": "reviewer", "Decisión": "decision",
    "Justificación CIIU": "just_ciiu", "Justificación CPC": "just_cpc",
    "Justificación Correspondencia": "just_match",
    "Observaciones": "observations", "Revisiòn PC": "review_pc",
}
ORIGINAL_FIELDS = (
    "inec_identificador_empresa", "inec_zonal", "efectividad", "provincia", "canton",
    "parroquia", "tamanio", "actividad_principal", "desc_actividad_principal",
    "ciiu4_actividad_principal", "actividad_secundaria", "desc_actividad_secundaria",
    "ciiu4_actividad_secundaria", "tercera_actividad", "desc_tercera_actividad",
    "ciiu4_tercera_actividad", "ruc", "nombre_comercial", "razon_social",
    "Descripción del servicio", "Código CPC", "ventas_nac", "ventas_extranj",
    "id_serv", "servicio",
)


def identity(row: dict) -> tuple[str, str]:
    return row["company_id"], row["service_id"]


def fingerprint(records: list[dict]) -> str:
    source = [(identity(r), [norm(r["source"].get(key)) for key in ORIGINAL_FIELDS])
              for r in records]
    payload = as_json(sorted(source, key=lambda x: x[0])).encode("utf-8")
    return sha256(payload).hexdigest()


def distribute(records: list[dict], assignments: dict[str, str] | None = None):
    """Balancea cantidad de servicios evaluables; cada empresa permanece unida."""
    assignments = dict(assignments or {})
    sizes = Counter(r["company_id"] for r in records if not is_other(r["source"]))
    counts = Counter({load: 0 for load in LOADS})
    for company, load in assignments.items():
        if load not in LOADS:
            raise ValueError(f"Carga desconocida para la empresa {company}: {load}")
        counts[load] += sizes[company]
    for company in sorted(sizes, key=lambda x: (-sizes[x], x)):
        if company not in assignments:
            load = min(LOADS, key=lambda x: (counts[x], x))
            assignments[company] = load
            counts[load] += sizes[company]
    # Las filas 'Otros' siguen a su empresa, aunque no se asignen para validación.
    for company in sorted({r["company_id"] for r in records}):
        if company not in assignments:
            assignments[company] = min(LOADS, key=lambda x: (counts[x], x))
    for row in records:
        row["load"] = assignments[row["company_id"]]


def _sheet(book, required: set[str]):
    for sheet in book:
        headings = [norm(c.value) for c in sheet[1]]
        if required.issubset(set(headings)):
            return sheet
    raise ValueError("No se encuentra la hoja de Servicios en el Excel.")


def load_state(content: bytes) -> dict:
    parsed, headings, warnings = parse_upload(content)
    book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    base = _sheet(book, {"inec_identificador_empresa", "id_serv", "Código CPC"})
    headers = {norm(cell.value): cell.column - 1 for cell in base[1] if norm(cell.value)}
    assignment_by_company = {}
    if "Carga asignada" in headers:
        index = headers["Carga asignada"]
        loads_by_row = {number: norm(values[index]) for number, values in enumerate(
            base.iter_rows(min_row=2, max_col=index + 1, values_only=True), 2)}
        for row in parsed:
            load = loads_by_row.get(row["source_row"], "")
            if load:
                prev = assignment_by_company.setdefault(row["company_id"], load)
                if prev != load:
                    raise ValueError(f"La empresa {row['company_id']} figura en dos cargas distintas.")
    distribute(parsed, assignment_by_company)
    for row in parsed:
        row["review"] = review_defaults(row["source"])
        row["version"] = 0
    lookup = {identity(r): r for r in parsed}
    if "Revision_app" in book.sheetnames:
        sheet = book["Revision_app"]
        h = {norm(v): i for i, v in enumerate(next(sheet.iter_rows(max_row=1, values_only=True))) if norm(v)}
        if {"empresa", "id_serv", "Versión revisión"}.issubset(h):
            for values in sheet.iter_rows(min_row=2, values_only=True):
                key = norm(values[h["empresa"]]), norm(values[h["id_serv"]])
                if key in lookup:
                    lookup[key]["version"] = int(values[h["Versión revisión"]] or 0)
                    for field in REVIEW_KEYS:
                        header = f"review_{field}"
                        if header in h:
                            lookup[key]["review"][field] = norm(values[h[header]])
                    if lookup[key]["review"]["decision"] == "Aceptado":
                        lookup[key]["review"]["decision"] = "Justificación"
    history = []
    if "Historial" in book.sheetnames:
        sheet = book["Historial"]
        for values in sheet.iter_rows(min_row=2, values_only=True):
            if not values[0]:
                continue
            if len(values) < 7 or not values[6]:
                raise ValueError("El historial no tiene ID evento; use un archivo exportado por esta versión.")
            history.append({"company_id": norm(values[0]), "service_id": norm(values[1]),
                            "changed_at": norm(values[2]), "changed_by": norm(values[3]),
                            "previous_json": norm(values[4]), "new_json": norm(values[5]),
                            "event_id": norm(values[6])})
    source_hash = fingerprint(parsed)
    if "Control" in book.sheetnames:
        control = book["Control"]
        saved_hash = norm(control["B2"].value)
        if saved_hash != source_hash:
            raise ValueError("Los datos originales cambiaron respecto al archivo de avance.")
    exported = "Control" in book.sheetnames
    book.close()
    return {"rows": parsed, "headings": headings, "history": history,
            "fingerprint": source_hash, "warnings": warnings,
            "exported": exported}


def assign_company(state: dict, company: str, load: str):
    if load not in LOADS:
        raise ValueError("Carga inválida.")
    if any(r["company_id"] == company and r["version"] > 0 for r in state["rows"]):
        raise ValueError("Esta empresa ya tiene revisiones. Consolide los avances antes de reasignarla.")
    for row in state["rows"]:
        if row["company_id"] == company:
            row["load"] = load


def update_review(state: dict, key: tuple[str, str], changes: dict, catalog):
    row = next((r for r in state["rows"] if identity(r) == key), None)
    if row is None or is_other(row["source"]):
        raise ValueError("Servicio no evaluable o no encontrado.")
    if set(changes) != set(REVIEW_KEYS):
        raise ValueError("Faltan campos de revisión.")
    if changes["decision"] not in ("Pendiente", "Justificación", "Cambio"):
        raise ValueError("La decisión debe ser Pendiente, Justificación o Cambio.")
    candidate = evaluate(row["source"], changes["edited_code"], catalog)
    edited = norm(changes["edited_code"])
    if edited and (not candidate["code_ok"] or edited == "99999999"):
        raise ValueError("El CPC corregido debe contener ocho dígitos y no puede ser 99999999.")
    if changes["decision"] == "Cambio" and (not edited or edited == norm(row["source"].get("Código CPC"))):
        raise ValueError("Para marcar Cambio, escribe un CPC distinto del original.")
    if not norm(changes["reviewer"]):
        raise ValueError("Indica el nombre del responsable.")
    if changes["decision"] == "Justificación" and not any(
        norm(changes[k]) for k in ("just_ciiu", "just_cpc", "just_match", "observations")
    ):
        raise ValueError("Registra la justificación o una observación para esta decisión.")
    if changes["decision"] == "Cambio" and candidate["alerts"] and not any(
        norm(changes[k]) for k in ("just_ciiu", "just_cpc", "just_match", "observations")
    ):
        raise ValueError("Hay alertas. Registra una justificación u observación.")
    cleaned = {k: norm(changes[k]) for k in REVIEW_KEYS}
    if cleaned == row["review"]:
        return False
    state["history"].append({"company_id": row["company_id"], "service_id": row["service_id"],
                             "changed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                             "changed_by": cleaned["reviewer"], "previous_json": as_json(row["review"]),
                             "new_json": as_json(cleaned), "event_id": str(uuid4())})
    row["review"] = cleaned
    row["version"] += 1
    return True


def reviewed_rows(state: dict, catalog) -> list[dict]:
    return [{**r, **r["review"], "validation": evaluate(r["source"], r["review"]["edited_code"], catalog)}
            for r in state["rows"]]


def export_state(state: dict, catalog) -> bytes:
    items = reviewed_rows(state, catalog)
    dataset = {"headers_json": as_json(state["headings"])}
    first = export_excel(dataset, items, state["history"])
    book = load_workbook(BytesIO(first))
    base = book["Servicios revisados"]
    existing = {norm(cell.value): cell.column for cell in base[1] if norm(cell.value)}
    next_col = max(55, base.max_column + 1)
    for header, field in REVIEW_COLUMNS.items():
        if header not in existing:
            existing[header] = next_col
            base.cell(1, next_col, header)
            next_col += 1
        for row_no, record in enumerate(state["rows"], 2):
            base.cell(row_no, existing[header], record["review"][field])
    col = existing.get("Carga asignada", next_col)
    base.cell(1, col, "Carga asignada")
    for row_no, record in enumerate(state["rows"], 2):
        base.cell(row_no, col, record["load"])
    detail = book["Revision_app"]
    detail_cols = detail.max_column
    detail.cell(1, detail_cols + 1, "Carga asignada")
    detail.cell(1, detail_cols + 2, "Versión revisión")
    for idx, key in enumerate(REVIEW_KEYS, detail_cols + 3):
        detail.cell(1, idx, f"review_{key}")
    number = 2
    for record in state["rows"]:
        if is_other(record["source"]):
            continue
        detail.cell(number, detail_cols + 1, record["load"])
        detail.cell(number, detail_cols + 2, record["version"])
        for idx, key in enumerate(REVIEW_KEYS, detail_cols + 3):
            detail.cell(number, idx, record["review"][key])
        number += 1
    hist = book["Historial"]
    hist.cell(1, 7, "ID evento")
    for row_no, event in enumerate(state["history"], 2):
        hist.cell(row_no, 7, event["event_id"])
    control = book.create_sheet("Control")
    control.append(["Campo", "Valor"])
    control.append(["Huella datos originales", state["fingerprint"]])
    control.append(["Formato", "ENESEM Servicios Excel 1"])
    control.append(["Cargas", ", ".join(LOADS)])
    for sheet in (base, detail, hist, control):
        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = "A2"
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


class MergeConflict(ValueError):
    pass


def merge_states(states: list[dict]) -> dict:
    """Une archivos de la misma fuente; nunca resuelve conflictos por última fecha."""
    if len(states) < 2:
        raise ValueError("Carga por lo menos dos avances para consolidar.")
    if not all(s.get("exported") for s in states):
        raise MergeConflict("Consolida solo avances descargados desde este tablero.")
    reference = states[0]
    expected = {identity(r) for r in reference["rows"]}
    for other in states[1:]:
        if other["fingerprint"] != reference["fingerprint"] or {identity(r) for r in other["rows"]} != expected:
            raise MergeConflict("Los archivos no proceden de la misma base original.")
        if {identity(r): r["load"] for r in other["rows"]} != {identity(r): r["load"] for r in reference["rows"]}:
            raise MergeConflict("Las cargas asignadas difieren entre archivos. Unifíquelas antes de consolidar.")
    import copy
    merged = copy.deepcopy(reference)
    indexes = [{identity(row): row for row in state["rows"]} for state in states]
    conflicts = []
    for row in merged["rows"]:
        key = identity(row)
        versions = [index[key] for index in indexes if index[key]["version"] > 0]
        if not versions:
            continue
        reviews = {as_json(r["review"]) for r in versions}
        if len(reviews) > 1:
            conflicts.append(f"{key[0]} / {key[1]}")
            continue
        best = max(versions, key=lambda r: r["version"])
        row["review"] = dict(best["review"])
        row["version"] = best["version"]
    if conflicts:
        raise MergeConflict("Hay revisiones distintas del mismo servicio: " + ", ".join(conflicts[:10]))
    events = {event["event_id"]: event for state in states for event in state["history"]}
    merged["history"] = sorted(events.values(), key=lambda e: (e["changed_at"], e["event_id"]))
    merged["warnings"] = []
    return merged
