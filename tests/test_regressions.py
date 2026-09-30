"""Regression checks; fixtures exercise code paths and are never application data."""
import ast
import copy
import io
from pathlib import Path
import unittest
from unittest.mock import Mock, patch

import pandas as pd
import requests
from streamlit.testing.v1 import AppTest

import PV_Blink_Login_Form as login

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "pv_blink_dashboard.py"


def dashboard_helpers():
    """Load existing helpers without executing the page or copying its logic."""
    tree = ast.parse(APP.read_text(encoding="utf-8"))
    nodes = []
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            nodes.append(node)
        elif isinstance(node, ast.FunctionDef):
            node = copy.deepcopy(node)
            node.decorator_list = []
            nodes.append(node)
        elif isinstance(node, ast.Assign):
            try:
                ast.literal_eval(node.value)
            except (ValueError, TypeError):
                continue
            nodes.append(node)
    namespace = {"__file__": str(APP)}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP), "exec"), namespace)
    return namespace


class LoginTests(unittest.TestCase):
    def response(self, status=200, data=None):
        response = Mock(status_code=status, ok=200 <= status < 400)
        response.json.return_value = data
        return response

    def test_success_requires_successful_http_response(self):
        body = {"data": {"error": False, "accessToken": "test-only-token"}}
        for status, expected in [(200, True), (401, False), (500, False)]:
            with self.subTest(status=status), patch.object(
                login.requests, "post", return_value=self.response(status, body)
            ) as post:
                self.assertEqual(login.login_with_api("test@example.invalid", "test")[0], expected)
                self.assertEqual(post.call_args.kwargs["timeout"], 20)

    def test_network_errors_are_safe(self):
        for error in [requests.ConnectionError(), requests.Timeout(), requests.RequestException("private-detail")]:
            with self.subTest(error=type(error).__name__), patch.object(login.requests, "post", side_effect=error):
                ok, message, _ = login.login_with_api("test@example.invalid", "test")
                self.assertFalse(ok)
                self.assertNotIn("private-detail", message)

    def test_invalid_or_unsuccessful_json(self):
        for body in [None, [], {}, {"data": {"error": False}}, {"data": {"error": True}}]:
            with self.subTest(body=body), patch.object(login.requests, "post", return_value=self.response(data=body)):
                self.assertFalse(login.login_with_api("test@example.invalid", "test")[0])
        response = self.response()
        response.json.side_effect = ValueError("not json")
        with patch.object(login.requests, "post", return_value=response):
            self.assertFalse(login.login_with_api("test@example.invalid", "test")[0])


class DataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.helpers = dashboard_helpers()

    def test_sales_without_optional_rate_or_representative(self):
        upload = io.BytesIO(
            b"Invoice Date,Ledger Name,Item Group,Item Name,QTY,Amount\n"
            b"2026-09-01,Test ledger,Inverter,3 KW inverter,2,100\n"
        )
        upload.name = "regression.csv"
        result, summary = self.helpers["process_data"](upload)
        banga_result, banga_summary = self.helpers["process_data"](upload, "Banga Solar Pvt Ltd")
        self.assertEqual(banga_result["Representative Ref"].tolist(), ["Banga Solar Pvt Ltd"])
        self.assertEqual(banga_result["Representative User"].tolist(), ["Banga Solar Pvt Ltd"])
        self.assertEqual(banga_result["Amount"].sum(), result["Amount"].sum())
        self.assertEqual(banga_summary, summary)
        self.assertEqual(summary["final"], 1)
        self.assertEqual(result["Amount"].sum(), 100)
        self.assertEqual(result["QTY"].sum(), 2)
        self.assertEqual(result.iloc[0]["Representative Ref"], "PV-Blink")

    def test_empty_sales_upload_is_rejected(self):
        upload = io.BytesIO(b"Invoice Date,Ledger Name,Item Group,Item Name,QTY,Amount\n")
        upload.name = "empty.csv"
        with self.assertRaises(ValueError):
            self.helpers["process_data"](upload)

    def test_service_product_groups_preserve_phase_and_mppt(self):
        catalog = self.helpers["PRODUCT_CHART_CATALOG"]
        frame = pd.DataFrame({"Product": catalog + [n.replace("(G)", "").lower() for n in catalog]})
        counts = self.helpers["service_product_counts"](frame).set_index("Product")["Count"]
        self.assertEqual(counts.sum(), 58)
        self.assertEqual(counts["5 KW 1 Phase 1 MPPT"], 2)
        self.assertEqual(counts["5 KW 1 Phase 2 MPPT"], 4)
        self.assertEqual(counts["5 KW 3 Phase 3 MPPT"], 2)

    def test_mixed_dates(self):
        result = self.helpers["parse_mixed_date_series"](pd.Series(["21-Sep-2026", "2026-09-22", "-", None]))
        self.assertEqual(result.notna().sum(), 2)

    def test_field_workbook_and_empty_activity(self):
        for rows in [[{"Emp_Name": "Test employee", "Date": "2026-09-01", "Visit": 2}], []]:
            with self.subTest(rows=len(rows)):
                upload = io.BytesIO()
                with pd.ExcelWriter(upload, engine="openpyxl") as writer:
                    pd.DataFrame(rows, columns=["Emp_Name", "Date", "Visit"]).to_excel(writer, sheet_name="Dashboard", index=False)
                upload.seek(0)
                result, _ = self.helpers["process_field_dashboard"](upload)
                self.assertEqual(len(result), len(rows))
                self.assertEqual(result["Visit"].sum(), 2 if rows else 0)

    def test_multiple_monthly_field_files(self):
        uploads = []
        for filename, employee, date, visits in [
            ("May.xlsx", "Test employee", "2026-05-01", 2),
            ("June.xlsx", "TEST EMPLOYEE", "2026-06-01", 3),
        ]:
            upload = io.BytesIO()
            upload.name = filename
            with pd.ExcelWriter(upload, engine="openpyxl") as writer:
                pd.DataFrame([{"Emp_Name": employee, "Date": date, "Visit": visits}]).to_excel(
                    writer, sheet_name="Dashboard", index=False
                )
            uploads.append(upload)
        result, info, errors = self.helpers["load_field_files"](uploads)
        self.assertEqual(errors, [])
        self.assertEqual(result["Visit"].sum(), 5)
        self.assertEqual(info["employees"], 1)
        self.assertEqual(info["dates"], 2)
        self.assertEqual(set(result["Source File"]), {"May.xlsx", "June.xlsx"})
        summary = self.helpers["field_employee_summary"](result)
        self.assertEqual(summary.iloc[0]["Visit"], 5)
        self.assertEqual(result.loc[result["Date"].dt.month.eq(6), "Visit"].sum(), 3)
        self.assertEqual(len(info["coverage"]), 2)
        broken = io.BytesIO(b"not an Excel file")
        broken.name = "broken.xlsx"
        partial, _, errors = self.helpers["load_field_files"]([uploads[0], broken])
        self.assertEqual(partial["Visit"].sum(), 2)
        self.assertEqual(len(errors), 1)
        self.assertIn("broken.xlsx", errors[0])
        empty, info, errors = self.helpers["load_field_files"]([])
        self.assertTrue(empty.empty)
        self.assertEqual(info, {})
        self.assertEqual(errors, [])

    def test_multiple_sales_and_return_registers(self):
        def csv_file(name, date_column, id_column, month, amount):
            upload = io.BytesIO(
                f"{date_column},{id_column},Ledger Name,Item Group,Item Name,QTY,Amount\n"
                f"21-{month}-2026,{name},Test ledger,Inverter,3 KW inverter,1,{amount}\n".encode()
            )
            upload.name = name
            return upload
        sales = [csv_file("May.csv", "Invoice Date", "Invoice No#", "May", 100),
                 csv_file("June.csv", "Invoice Date", "Invoice No#", "Jun", 200)]
        returns = [csv_file("Return-May.csv", "Sales Return Date", "Order No#", "May", 10),
                   csv_file("Return-June.csv", "Sales Return Date", "Order No#", "Jun", 20)]
        sf, audit, errors = self.helpers["load_register_files"](sales, self.helpers["process_data"])
        rf, raudit, rerrors = self.helpers["load_register_files"](returns, self.helpers["process_return_data"])
        self.assertEqual(errors + rerrors, [])
        self.assertEqual(sf["Amount"].sum(), 300)
        self.assertEqual(rf["Amount"].sum(), 30)
        self.assertEqual(self.helpers["net_metrics"](sf, rf)["net_amount"], 270)
        self.assertEqual(audit["final"], 2)
        self.assertEqual(len(audit["files"]), 2)
        self.assertEqual(raudit["return_amount"], 30)
        self.assertEqual(raudit["validation_status"], "PASS")
        self.assertEqual(sf["Invoice Date"].dt.month.tolist(), [5, 6])
        self.assertEqual(set(sf["Source File"]), {"May.csv", "June.csv"})
        broken = io.BytesIO(b"bad file")
        broken.name = "bad.xlsx"
        partial, _, errors = self.helpers["load_register_files"]([sales[0], broken], self.helpers["process_data"])
        self.assertEqual(partial["Amount"].sum(), 100)
        self.assertEqual(len(errors), 1)

    def test_service_multiple_workbooks(self):
        uploads = []
        for name, ticket, status in [("May.xlsx", "T1", "In-Progress"), ("June.xlsx", "T2", "Closed")]:
            upload = io.BytesIO()
            upload.name = name
            with pd.ExcelWriter(upload, engine="openpyxl") as writer:
                pd.DataFrame([{"Ticket No": ticket, "MTCE Status": status}]).to_excel(
                    writer, startrow=6, index=False
                )
            uploads.append(upload)
        frame, audit = self.helpers["load_service_files"](uploads)
        self.assertEqual(len(frame), 2)
        self.assertEqual(len(audit), 2)
        self.assertEqual(set(frame["Status"]), {"Open", "Closed"})
        self.assertEqual(set(frame["Source File"]), {"May.xlsx", "June.xlsx"})



class StartupTests(unittest.TestCase):
    def test_company_switch_updates_brand_and_clears_company_data(self):
        app = AppTest.from_file(str(APP), default_timeout=45)
        app.session_state["pv_authenticated"] = True
        app.run()
        self.assertEqual(app.selectbox(key="dashboard_company").value, "PV-Blink Inverter")
        app.session_state["sales_dates_v9"] = "old-company-filter"
        app.selectbox(key="dashboard_company").set_value("Banga Solar Pvt Ltd").run()
        self.assertFalse(app.exception, [e.message for e in app.exception])
        self.assertTrue(app.session_state["pv_authenticated"])
        self.assertNotIn("sales_dates_v9", app.session_state)
        markup = "\n".join(m.value for m in app.markdown)
        self.assertIn("Banga Solar Pvt Ltd</div>", markup)
        self.assertIn("--pv-accent:#166534", markup)
        self.assertNotIn("Companion of Clean Energy", markup)
        for page in app.radio[0].options:
            app.radio[0].set_value(page).run()
            self.assertFalse(app.exception, [e.message for e in app.exception])
        app.selectbox(key="dashboard_company").set_value("PV-Blink Inverter").run()
        markup = "\n".join(m.value for m in app.markdown)
        self.assertIn("PV-Blink Inverter</div>", markup)
        self.assertIn("--pv-accent:#f97316", markup)
        state = {"dashboard_company": "Banga Solar Pvt Ltd", "pv_authenticated": True,
                 "sales_upload": ["pv-sales"], "return_upload": ["pv-returns"],
                 "service_upload": ["pv-service"], "field_dashboard_upload": ["pv-field"]}
        with patch("streamlit.session_state", state):
            dashboard_helpers()["reset_company_data"]()
        self.assertEqual(state, {"dashboard_company": "Banga Solar Pvt Ltd", "pv_authenticated": True})

    def test_pages_with_multiple_field_files(self):
        uploads = []
        for month in [5, 6]:
            upload = io.BytesIO()
            upload.name = f"month-{month}.xlsx"
            with pd.ExcelWriter(upload, engine="openpyxl") as writer:
                pd.DataFrame([{"Emp_Name": "Test employee", "Date": pd.Timestamp(2026, month, 1), "Visit": month}]).to_excel(
                    writer, sheet_name="Dashboard", index=False
                )
            uploads.append(upload)

        def uploader(*args, **kwargs):
            if kwargs.get("key") == "field_dashboard_upload":
                self.assertTrue(kwargs["accept_multiple_files"])
                self.assertTrue(callable(kwargs["on_change"]))
                return uploads
            return [] if kwargs.get("key") == "service_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP), default_timeout=45)
            app.session_state["pv_authenticated"] = True
            app.run()
            for page in ["Field KPI / KRA", "Quality & Export"]:
                app.radio[0].set_value(page).run()
                self.assertFalse(app.exception, [e.message for e in app.exception])
            app.radio[0].set_value("Field KPI / KRA").run()
            date_filter = next(d for d in app.date_input if d.label == "Field Date Range")
            self.assertEqual(date_filter.value[0].month, 5)
            self.assertEqual(date_filter.value[1].month, 6)
        state = {"field_emp_filter": ["Test employee"], "field_date_filter_v26": (), "keep": True}
        with patch("streamlit.session_state", state):
            dashboard_helpers()["reset_field_upload_filters"]()
        self.assertEqual(state, {"keep": True})

    def test_login_and_all_pages_without_uploads(self):
        app = AppTest.from_file(str(APP), default_timeout=45).run()
        self.assertFalse(app.exception)
        # Test navigation directly; authentication itself is covered above.
        app.session_state["pv_authenticated"] = True
        app.run()
        self.assertFalse(app.exception)
        for page in app.radio[0].options:
            with self.subTest(page=page):
                app.radio[0].set_value(page).run()
                self.assertFalse(app.exception, [e.message for e in app.exception])
        app.button(key="pv_logout_button").click().run()
        self.assertFalse(app.exception)
        self.assertNotIn("pv_authenticated", app.session_state)

    def test_pages_with_sales_and_returns(self):
        sales = io.BytesIO(
            b"Invoice Date,Invoice No#,Ledger Name,Item Group,Item Name,QTY,Amount\n"
            b"2026-09-01,INV-1,Test ledger,Inverter,3 KW inverter,2,100\n"
        )
        sales.name = "regression-sales.csv"
        returns = io.BytesIO(
            b"Sales Return Date,Order No#,Ledger Name,Item Group,Item Name,QTY,Amount\n"
            b"2026-09-02,RET-1,Test ledger,Inverter,3 KW inverter,1,50\n"
        )
        returns.name = "regression-returns.csv"
        sales_next = io.BytesIO(sales.getvalue().replace(b"2026-09-01", b"2026-10-21").replace(b"INV-1", b"INV-2"))
        sales_next.name = "next-sales.csv"
        returns_next = io.BytesIO(returns.getvalue().replace(b"2026-09-02", b"2026-10-22").replace(b"RET-1", b"RET-2"))
        returns_next.name = "next-returns.csv"

        def uploader(*args, **kwargs):
            if kwargs.get("key") in {"sales_upload", "return_upload", "service_upload"}:
                self.assertTrue(kwargs["accept_multiple_files"])
            return {"sales_upload": [sales, sales_next], "return_upload": [returns, returns_next], "service_upload": []}.get(kwargs.get("key"))

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP), default_timeout=45)
            app.session_state["pv_authenticated"] = True
            app.run()
            self.assertFalse(app.exception, [e.message for e in app.exception])
            for company in ["PV-Blink Inverter", "Banga Solar Pvt Ltd"]:
                app.selectbox(key="dashboard_company").set_value(company).run()
                for page in app.radio[0].options:
                    with self.subTest(company=company, page=page):
                        app.radio[0].set_value(page).run()
                        self.assertFalse(app.exception, [e.message for e in app.exception])
                        self.assertFalse(app.error, [e.value for e in app.error])


if __name__ == "__main__":
    unittest.main()
