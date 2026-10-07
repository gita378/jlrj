import io
import unittest
from contextlib import closing
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfWriter
from restore import restore

from server import (add_months, amount, connect, equipment_view, import_billing,
                    import_equipment, init_db, review_certificate, save_equipment,
                    upload_pdf, backup_bytes, archive_equipment, parse_certificate_text,
                    save_submission, return_submission, extract_pdf)


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
            with self.assertRaisesRegex(ValueError, "处理原因"):
                save_equipment(self.db, {"abnormal_flag": ""}, "Tester", device["id"])
            cleared = save_equipment(self.db, {"abnormal_flag": "", "closure_note": "复核完成"}, "Tester", device["id"])
            self.assertEqual(cleared["status"], "正常")
            self.assertIn("复核完成", cleared["abnormal_note"])

    def test_rejected_and_failed_certificate_are_not_valid_calibration(self):
        with self.db:
            device = save_equipment(self.db, {"asset_code": "A-1", "name": "压力表"}, "Tester")
            self.cert(device["id"], "FAIL", "2026-01-01", "不合格")
            view = equipment_view(self.db, self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone())
            self.assertIsNone(view["last_cal_date"])
            self.assertEqual(view["status"], "异常")
            self.assertIn("不合格", view["certificate_issue"])
            self.cert(device["id"], "PASS", "2026-02-01", "合格")
            view = equipment_view(self.db, self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone())
            self.assertEqual(view["status"], "正常")

    def test_billing_keeps_duplicates_and_checks_real_price(self):
        from server import reconcile_line
        with self.db:
            device = save_equipment(self.db, {"asset_code": "0007", "name": "卡尺"}, "Tester")
            self.cert(device["id"], "CERT-7", "2026-02-01")
            self.cert(device["id"], "CERT-8", "2026-02-02")
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
            result = import_billing(self.db, {"file_name": "bill.xlsx", "calibration_month": "2026-02",
                "content": base64.b64encode(out.getvalue()).decode()}, "Tester")
            self.assertEqual([item["finding"] for item in result["findings"]], ["金额差异", "疑似重复", "清单缺失"])
            self.assertEqual(self.db.execute("SELECT COUNT(*) FROM billing_record").fetchone()[0], 3)

    def test_submission_is_counted_by_actual_send_date(self):
        with self.db:
            device = save_equipment(self.db, {"asset_code": "SEND-1", "name": "探头"}, "Tester")
            sent = save_submission(self.db, {"asset_code": "SEND-1", "sent_date": "2026-01-12"}, "Tester")
            self.assertEqual(sent["equipment_id"], device["id"])
            returned = return_submission(self.db, sent["id"], {"returned_date": "2026-01-20"}, "Tester")
            self.assertEqual(returned["returned_date"], "2026-01-20")
            self.assertEqual(self.db.execute("SELECT COUNT(*) FROM audit_log WHERE entity='submission'").fetchone()[0], 2)

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

    @staticmethod
    def text_pdf(lines):
        content = b"BT /F1 12 Tf 20 260 Td\n"
        for line in lines:
            content += b"(" + line.encode("ascii") + b") Tj 0 -20 Td\n"
        content += b"ET"
        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
        ]
        pdf = b"%PDF-1.4\n"
        offsets = [0]
        for index, obj in enumerate(objects, 1):
            offsets.append(len(pdf))
            pdf += f"{index} 0 obj\n".encode() + obj + b"\nendobj\n"
        start = len(pdf)
        pdf += f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode()
        for offset in offsets[1:]:
            pdf += f"{offset:010d} 00000 n \n".encode()
        return pdf + f"trailer\n<< /Root 1 0 R /Size {len(offsets)} >>\nstartxref\n{start}\n%%EOF\n".encode()

    def test_text_pdf_auto_updates_exact_equipment(self):
        import base64
        pdf_dir = Path(__file__).resolve().parent / "_pdfs_auto"
        pdf_dir.mkdir(exist_ok=True)
        try:
            with self.db:
                device = save_equipment(self.db, {"asset_code": "EQ-001", "name": "数字表"}, "Tester")
                raw = self.text_pdf(["Asset Code: EQ-001", "Certificate No: CERT-001",
                                     "Calibration Date: 2026-02-01", "Valid Until: 2027-02-01",
                                     "Calibration Result: Pass"])
                cert = upload_pdf(self.db, {"file_name": "text.pdf",
                    "content": base64.b64encode(raw).decode()}, "Tester", pdf_dir)
                self.assertEqual(cert["state"], "confirmed")
                view = equipment_view(self.db, self.db.execute("SELECT * FROM equipment WHERE id=?", (device["id"],)).fetchone())
                self.assertEqual(view["last_cal_date"], "2026-02-01")
                self.assertEqual(view["next_cal_date"], "2027-02-01")
                self.assertEqual(view["status"], "正常")
                ambiguous = self.text_pdf(["Asset Code: EQ-001", "Asset Code: EQ-002",
                    "Certificate No: CERT-002", "Calibration Date: 2026-02-01", "Calibration Result: Pass"])
                other = upload_pdf(self.db, {"file_name": "ambiguous.pdf",
                    "content": base64.b64encode(ambiguous).decode()}, "Tester", pdf_dir)
                self.assertEqual(other["state"], "pending")
        finally:
            for file in pdf_dir.glob("*.pdf"):
                file.unlink()
            pdf_dir.rmdir()

    def test_chinese_certificate_labels(self):
        fields, _ = parse_certificate_text("设备编号：EQ-002\n证书编号：C-002\n校准日期：2026年2月1日\n校准结论：合格")
        self.assertEqual(fields["asset_code"], "EQ-002")
        self.assertEqual(fields["cert_no"], "C-002")
        self.assertEqual(fields["cal_date"], "2026-02-01")
        self.assertEqual(fields["result"], "合格")

    def test_chinese_text_pdf_when_reportlab_available(self):
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            from reportlab.pdfgen import canvas
        except ImportError:
            self.skipTest("报表生成库未安装；中文文本解析已由独立测试覆盖")
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        output = io.BytesIO()
        pdf = canvas.Canvas(output)
        pdf.setFont("STSong-Light", 12)
        for index, line in enumerate(("设备编号：EQ-003", "证书编号：C-003", "校准日期：2026年2月1日", "校准结论：合格")):
            pdf.drawString(40, 780 - index * 30, line)
        pdf.save()
        found, _ = extract_pdf(output.getvalue())
        self.assertEqual(found.get("asset_code"), "EQ-003")
        self.assertEqual(found.get("cert_no"), "C-003")

    def test_pdf_without_text_pending_duplicate_archive_and_backup(self):
        import base64
        import zipfile
        pdf_dir = Path(__file__).resolve().parent / "_pdfs"
        pdf_dir.mkdir(exist_ok=True)
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        stream = io.BytesIO()
        writer.write(stream)
        payload = {"file_name": "blank.pdf", "content": base64.b64encode(stream.getvalue()).decode()}
        try:
            with self.db:
                device = save_equipment(self.db, {"asset_code": "0011", "name": "空白PDF设备"}, "Tester")
                cert = upload_pdf(self.db, payload, "Tester", pdf_dir)
                self.assertEqual(cert["state"], "pending")
                self.assertIn("核对", cert["review_note"])
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
