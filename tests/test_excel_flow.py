import io
import sys
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import ciiu_for_code, get_catalog, summarize_company
from workflow import (MergeConflict, assign_company, export_state, identity,
                      load_state, merge_states, reviewed_rows, update_review)


def input_excel():
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Servicios"
    sheet.append(["inec_identificador_empresa", "ciiu4_actividad_principal",
                  "ciiu4_actividad_secundaria", "ciiu4_tercera_actividad",
                  "Descripción del servicio", "Código CPC", "id_serv", "servicio",
                  "ventas_nac", "ventas_extranj", "Decisión", "Correspondencia=1"])
    sheet.append(["56", "F439020", "", "", "ALQUILER DE GRÚAS", "43905452", "562", "Principal", 90, 10, "Pendiente", 2])
    sheet.append(["27", "G451001", "", "", "VENTA DE VEHÍCULOS", "45106118", "271", "Principal", 200, 0, "Pendiente", 2])
    sheet.append(["56", "F439020", "", "", "OTROS", "99999999", "564", "Otros", 20, 5, "Pendiente", 2])
    data = io.BytesIO()
    wb.save(data)
    return data.getvalue()


def changes(record, name, comment):
    updated = dict(record["review"])
    updated.update({"decision": "Justificación", "reviewer": name, "observations": comment})
    return updated


class ExcelFlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalog = get_catalog()

    def test_activity_cpc_reference_uses_complete_ciiu_and_unique_codes(self):
        pairs = self.catalog.activity_pairs
        self.assertEqual(len(pairs["C102002"]), 5)
        self.assertEqual([code for code, _ in pairs["G463033"]], ["46306112"])
        self.assertEqual({code for code, _ in pairs["H521000"]},
                         {"52106721", "52106722", "52106729"})

    def test_reverse_reference_distinguishes_exact_code_from_same_cpc(self):
        exact = ciiu_for_code(self.catalog, "45208714")
        self.assertTrue(exact)
        self.assertTrue(all(code == "45208714" for _, _, code in exact))
        self.assertTrue(any(activity == "G452001" for activity, _, _ in exact))
        broad = ciiu_for_code(self.catalog, "45208714", include_same_cpc=True)
        self.assertGreaterEqual(len(broad), len(exact))
        self.assertTrue(all(code.endswith("8714") for _, _, code in broad))

    def test_old_accepted_status_is_renamed_on_import(self):
        wb = load_workbook(io.BytesIO(input_excel()))
        wb["Servicios"]["K2"] = "Aceptado"
        buffer = io.BytesIO()
        wb.save(buffer)
        self.assertEqual(load_state(buffer.getvalue())["rows"][0]["review"]["decision"],
                         "Justificación")

    def test_roundtrip_preserves_assignments_decisions_and_values(self):
        state = load_state(input_excel())
        rows = state["rows"]
        self.assertEqual(len({r["load"] for r in rows if r["service_id"] != "564"}), 2)
        self.assertEqual(rows[0]["load"], rows[2]["load"])
        update_review(state, ("56", "562"), changes(rows[0], "Ana", "Servicio de excepción revisado."), self.catalog)
        file = export_state(state, self.catalog)
        again = load_state(file)
        self.assertTrue(again["exported"])
        self.assertEqual(again["fingerprint"], state["fingerprint"])
        self.assertEqual(again["rows"][0]["review"]["reviewer"], "Ana")
        self.assertEqual(again["rows"][0]["version"], 1)
        self.assertEqual(again["rows"][2]["load"], again["rows"][0]["load"])
        self.assertEqual(len(again["history"]), 1)
        wb = load_workbook(io.BytesIO(file), data_only=True)
        s = wb["Servicios revisados"]
        head = {c.value: c.column for c in s[1] if c.value}
        self.assertEqual(s.cell(2, head["ventas_nac"]).value, 90)
        self.assertEqual(s.cell(3, head["ventas_extranj"]).value, 0)
        self.assertIsNone(s.cell(4, head["Correspondencia=1"]).value)
        self.assertEqual(s.cell(2, head["Nombre del responsable"]).value, "Ana")

    def test_parallel_merge_and_conflict(self):
        base = input_excel()
        a, b = load_state(base), load_state(base)
        update_review(a, ("56", "562"), changes(a["rows"][0], "Ana", "Aprobado por excepción."), self.catalog)
        update_review(b, ("27", "271"), changes(b["rows"][1], "Beto", "Revisado con informante."), self.catalog)
        merged = merge_states([load_state(export_state(a, self.catalog)),
                               load_state(export_state(b, self.catalog))])
        found = {identity(row): row for row in merged["rows"]}
        self.assertEqual(found[("56", "562")]["review"]["reviewer"], "Ana")
        self.assertEqual(found[("27", "271")]["review"]["reviewer"], "Beto")
        self.assertEqual(len(merged["history"]), 2)
        self.assertEqual(merged["fingerprint"], a["fingerprint"])
        c = load_state(base)
        update_review(c, ("56", "562"), changes(c["rows"][0], "Cris", "Corrección diferente."), self.catalog)
        with self.assertRaisesRegex(MergeConflict, "revisiones distintas"):
            merge_states([load_state(export_state(a, self.catalog)),
                          load_state(export_state(c, self.catalog))])

    def test_assignment_change_keeps_company_together(self):
        state = load_state(input_excel())
        assign_company(state, "56", "Carga 5")
        self.assertTrue(all(r["load"] == "Carga 5" for r in state["rows"] if r["company_id"] == "56"))
        self.assertEqual(load_state(export_state(state, self.catalog))["rows"][0]["load"], "Carga 5")


if __name__ == "__main__":
    unittest.main()
