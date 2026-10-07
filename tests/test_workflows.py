import io
import unittest
from contextlib import closing
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfWriter
from restore import restore

from server import (add_months, amount, connect, equipment_view, import_billing,
                    import_equipment, init_db, review_certificate, save_equipment,
                    upload_pdf, backup_bytes, archive_equipment)


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.db_path = Path(__file__).resolve().parent / "_workflow.sqlite3"
        self.db_path.unlink(missing_ok=True)
        init_db(self.db_path)
        self.db = connect(self.db_path)

    def tearDown(self):
        self.db.close()
        self.db_path.unlink(missing_ok=True)
        self.db_path.with_name(self.db_path.name + "-wal").unlink(missing_ok=True)
        self.db_path.with_name(self.db_path.name + "-shm").unlink(missing_ok=True)

    def cert(self, equipment_id, cert_no, cal_date, result="合格", agreed="100.00"):
        cursor = self.db.execute("""INSERT INTO certificate(
            equipment_id,cert_no,file_name,file_path,cert_hash,cal_date,result,agreed_amount,created_at
        ) VALUES(?,?,?,?,?,?,?,?,?)""", (equipment_id, cert_no, "sample.pdf", "sample.pdf",
                                        cert_no, cal_date, result, agreed, "2026-01-01"))
        review_certificate(self.db, cursor.lastrowid, {}, "Tester")
        return cursor.lastrowid

    def test_old_certificate_cannot_replace_newer_and_abnormal_is_sticky(self):
        with self.db:
            device = save_equipment(self.db, {"asset_code": "00123", "name": "温度计",
                                                   "calibration_cycle": 12}, "Tester")
            self.cert(device["id"], "NEW", "2026-01-01")
            self.cert(device["id"], "OLD", "2025-01-01")
            row = self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone()
            view = equipment_view(self.db, row)
            self.assertEqual(view["asset_code"], "00123")
            self.assertEqual(view["last_cal_date"], "2026-01-01")
            self.assertEqual(view["next_cal_date"], "2027-01-01")
            save_equipment(self.db, {"abnormal_flag": "异常未检", "abnormal_note": "人工待确认"}, "Tester", device["id"])
            view = equipment_view(self.db, self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone())
            self.assertEqual(view["status"], "异常")
            self.assertEqual(view["abnormal_note"], "人工待确认")

    def test_rejected_and_failed_certificate_are_not_valid_calibration(self):
        with self.db:
            device = save_equipment(self.db, {"asset_code": "A-1", "name": "压力表"}, "Tester")
            self.cert(device["id"], "FAIL", "2026-01-01", "不合格")
            view = equipment_view(self.db, self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone())
            self.assertIsNone(view["last_cal_date"])
            self.assertEqual(view["status"], "异常")
            self.assertEqual(view["abnormal_flag"], "计量不过")

    def test_billing_keeps_duplicates_and_checks_real_price(self):
        from server import reconcile_line
        with self.db:
            device = save_equipment(self.db, {"asset_code": "0007", "name": "卡尺"}, "Tester")
            self.cert(device["id"], "CERT-7", "2026-02-01")
            self.assertEqual(reconcile_line(self.db, "0007", "CERT-7", amount("105"))[0], "金额差异")
            self.assertEqual(reconcile_line(self.db, "0008", "CERT-7", amount("100"))[0], "设备缺失")
            wb = Workbook()
            ws = wb.active
            ws.append(["资产编号", "证书编号", "金额"])
            ws.append(["0007", "CERT-7", 105])
            ws.append(["0007", "CERT-7", 105])
            out = io.BytesIO()
            wb.save(out)
            import base64
            result = import_billing(self.db, {"file_name": "bill.xlsx",
                "content": base64.b64encode(out.getvalue()).decode()}, "Tester")
            self.assertEqual([item["finding"] for item in result["findings"]], ["金额差异", "疑似重复"])
            self.assertEqual(self.db.execute("SELECT COUNT(*) FROM billing_record").fetchone()[0], 2)

    def test_excel_import_and_audit(self):
        import base64
        wb = Workbook()
        ws = wb.active
        ws.append(["资产编号", "设备名称", "ABC分类", "检定周期(月)"])
        ws.append(["0009", "电子秤", "A", 6])
        ws.append(["0009", "电子秤新", "B", 12])
        out = io.BytesIO()
        wb.save(out)
        result = import_equipment(self.db, {"file_name": "equipment.xlsx",
            "content": base64.b64encode(out.getvalue()).decode()}, "Tester")
        self.assertEqual(result["imported"], 2)
        self.assertEqual(self.db.execute("SELECT name FROM equipment WHERE asset_code='0009'").fetchone()[0], "电子秤新")
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0], 2)

    def test_month_end(self):
        from datetime import date
        self.assertEqual(add_months(date(2026, 1, 31), 1), date(2026, 2, 28))

    def test_scanned_pdf_pending_duplicate_archive_and_backup(self):
        import base64
        import zipfile
        pdf_dir = Path(__file__).resolve().parent / "_pdfs"
        pdf_dir.mkdir(exist_ok=True)
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        stream = io.BytesIO()
        writer.write(stream)
        payload = {"file_name": "scan.pdf", "content": base64.b64encode(stream.getvalue()).decode()}
        try:
            with self.db:
                device = save_equipment(self.db, {"asset_code": "0011", "name": "扫描件设备"}, "Tester")
                cert = upload_pdf(self.db, payload, "Tester", pdf_dir)
                self.assertEqual(cert["state"], "pending")
                self.assertIn("人工录入", cert["review_note"])
                with self.assertRaisesRegex(ValueError, "已上传"):
                    upload_pdf(self.db, payload, "Tester", pdf_dir)
                archive_equipment(self.db, device["id"], "设备报废", "Tester")
                self.assertEqual(self.db.execute("SELECT deleted FROM equipment WHERE id=?", (device["id"],)).fetchone()[0], 1)
            backup = backup_bytes(self.db_path, pdf_dir)
            with zipfile.ZipFile(io.BytesIO(backup)) as archive:
                self.assertIn("register.sqlite3", archive.namelist())
                self.assertTrue(any(path.startswith("certificates/") for path in archive.namelist()))
            restore_archive = Path(__file__).resolve().parent / "_backup_test.zip"
            restore_db = Path(__file__).resolve().parent / "_restored.sqlite3"
            restore_pdfs = Path(__file__).resolve().parent / "_restored_pdfs"
            restore_archive.write_bytes(backup)
            try:
                restore(restore_archive, restore_db, restore_pdfs, guard_server=False)
                with closing(connect(restore_db)) as restored:
                    self.assertEqual(restored.execute("SELECT asset_code FROM equipment").fetchone()[0], "0011")
                self.assertTrue(any(restore_pdfs.glob("*.pdf")))
            finally:
                restore_archive.unlink(missing_ok=True)
                restore_db.unlink(missing_ok=True)
                for file in restore_pdfs.glob("*.pdf"):
                    file.unlink()
                restore_pdfs.rmdir()
        finally:
            for file in pdf_dir.glob("*.pdf"):
                file.unlink()
            pdf_dir.rmdir()


if __name__ == "__main__":
    unittest.main()
