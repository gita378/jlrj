"""Native Windows desktop interface for the offline calibration register."""

from __future__ import annotations

import base64
import json
import os
import sqlite3
import tkinter as tk
from contextlib import closing
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from server import (
    DATA, DB_FILE, PDF_DIR, add_months, archive_equipment, backup_bytes,
    connect, equipment_view, import_billing, import_equipment, init_db,
    review_certificate, save_equipment, save_submission, return_submission,
    upload_pdf, workbook_bytes,
)
from restore import restore


NAVY = "#12332f"
TEAL = "#13745b"
PALE = "#f3f7f5"
WHITE = "#ffffff"
INK = "#20332d"
MUTED = "#64756d"
RED = "#b94b42"
AMBER = "#a56a13"
FONT = ("Microsoft YaHei UI", 10)
HEAD_FONT = ("Microsoft YaHei UI", 21, "bold")
NAV = [
    ("dashboard", "总览"), ("equipment", "设备台账"), ("submissions", "送检记录"),
    ("certificates", "计量证书"), ("billing", "结账对账"), ("audit", "审计与备份"),
]


def payload_for(path):
    raw = Path(path).read_bytes()
    return {"file_name": Path(path).name, "content": base64.b64encode(raw).decode("ascii")}


class Editor(tk.Toplevel):
    """Small ERP form dialog with native controls."""

    def __init__(self, parent, title, fields, values=None, on_save=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=WHITE)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.vars = {}
        self.on_save = on_save
        tk.Label(self, text=title, bg=WHITE, fg=INK,
                 font=("Microsoft YaHei UI", 17, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=24, pady=(22, 16))
        for index, (key, label, choices) in enumerate(fields, 1):
            tk.Label(self, text=label, bg=WHITE, fg=MUTED, font=FONT).grid(
                row=index, column=0, sticky="w", padx=(24, 12), pady=7)
            value = (values or {}).get(key)
            var = tk.StringVar(value="" if value is None else str(value))
            self.vars[key] = var
            if choices:
                widget = ttk.Combobox(self, textvariable=var, values=choices,
                                      state="readonly", width=35)
            else:
                widget = ttk.Entry(self, textvariable=var, width=38)
            widget.grid(row=index, column=1, sticky="ew", padx=(0, 24), pady=7)
            if index == 1:
                widget.focus_set()
        buttons = tk.Frame(self, bg=WHITE)
        buttons.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="e", padx=24, pady=20)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=(9, 0))
        ttk.Button(buttons, text="保存", style="Accent.TButton", command=self.save).pack(side="right")
        self.bind("<Return>", lambda _: self.save())
        self.bind("<Escape>", lambda _: self.destroy())
        self.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")

    def save(self):
        values = {key: var.get().strip() for key, var in self.vars.items()}
        if self.on_save and self.on_save(values):
            self.destroy()


class DesktopApp(tk.Tk):
    def __init__(self):
        super().__init__()
        init_db()
        PDF_DIR.mkdir(parents=True, exist_ok=True)
        self.title("计量设备管理")
        self.geometry("1380x850")
        self.minsize(1050, 650)
        self.configure(bg=PALE)
        self.operator = tk.StringVar(value="本机用户")
        self.view = "dashboard"
        self.nav_buttons = {}
        self.search = tk.StringVar()
        self._configure_style()
        self._shell()
        self.show("dashboard")

    def _configure_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=PALE)
        style.configure("TLabel", background=PALE, foreground=INK, font=FONT)
        style.configure("TButton", font=FONT, padding=(13, 8), background="#e6efea", foreground=INK)
        style.map("TButton", background=[("active", "#d5e5dc")])
        style.configure("Accent.TButton", background=TEAL, foreground=WHITE)
        style.map("Accent.TButton", background=[("active", "#0b604a")])
        style.configure("Treeview", font=FONT, rowheight=31, background=WHITE,
                        fieldbackground=WHITE, foreground=INK, borderwidth=0)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"),
                        background="#eaf1ed", foreground=INK, padding=8)
        style.map("Treeview", background=[("selected", "#d3e9de")], foreground=[("selected", INK)])
        style.configure("TEntry", padding=5)

    def _shell(self):
        sidebar = tk.Frame(self, bg=NAVY, width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="计量设备管理", bg=NAVY, fg=WHITE,
                 font=("Microsoft YaHei UI", 17, "bold"), anchor="w").pack(fill="x", padx=22, pady=(30, 4))
        tk.Label(sidebar, text="离线桌面版", bg=NAVY, fg="#a9d4c4",
                 font=FONT, anchor="w").pack(fill="x", padx=22, pady=(0, 35))
        for key, label in NAV:
            button = tk.Button(sidebar, text=label, font=FONT, fg="#d1e8dd", bg=NAVY,
                               activebackground="#1d5449", activeforeground=WHITE,
                               anchor="w", relief="flat", borderwidth=0, padx=22, pady=13,
                               command=lambda name=key: self.show(name))
            button.pack(fill="x")
            self.nav_buttons[key] = button
        tk.Label(sidebar, text="本机数据 · SQLite", bg=NAVY, fg="#89b6a5",
                 font=("Microsoft YaHei UI", 9)).pack(side="bottom", anchor="w", padx=22, pady=22)
        self.right = tk.Frame(self, bg=PALE)
        self.right.pack(side="left", expand=True, fill="both")
        top = tk.Frame(self.right, bg=WHITE, height=72)
        top.pack(fill="x")
        top.pack_propagate(False)
        self.page_title = tk.Label(top, text="", bg=WHITE, fg=INK, font=HEAD_FONT)
        self.page_title.pack(side="left", padx=28)
        tk.Label(top, text="操作人", bg=WHITE, fg=MUTED, font=FONT).pack(side="right", padx=(0, 28))
        ttk.Entry(top, textvariable=self.operator, width=18).pack(side="right", padx=8)
        self.body = tk.Frame(self.right, bg=PALE)
        self.body.pack(fill="both", expand=True, padx=26, pady=22)
        self.status = tk.StringVar(value="就绪")
        tk.Label(self.right, textvariable=self.status, bg=WHITE, fg=MUTED,
                 font=("Microsoft YaHei UI", 9), anchor="w").pack(fill="x", ipady=7, padx=0)

    def show(self, name):
        self.view = name
        for key, button in self.nav_buttons.items():
            button.configure(bg="#1c5749" if key == name else NAVY,
                             fg=WHITE if key == name else "#d1e8dd")
        self.page_title.configure(text=dict(NAV)[name])
        for child in self.body.winfo_children():
            child.destroy()
        getattr(self, f"page_{name}")()

    def _call(self, fn, success="操作完成", refresh=True):
        try:
            with closing(connect()) as db, db:
                result = fn(db)
            self.status.set(success)
            if refresh:
                self.show(self.view)
            return result
        except (ValueError, KeyError, OSError, sqlite3.Error) as exc:
            messagebox.showerror("操作未完成", str(exc), parent=self)
            self.status.set(str(exc))
            return None

    def _operator(self):
        return self.operator.get().strip() or "本机用户"

    def _toolbar(self, *buttons):
        bar = tk.Frame(self.body, bg=PALE)
        bar.pack(fill="x", pady=(0, 18))
        for index, (label, action) in enumerate(buttons):
            ttk.Button(bar, text=label, command=action,
                       style="Accent.TButton" if index == 0 else "TButton").pack(side="left", padx=(0, 8))
        return bar

    def _table(self, columns, rows, height=15, on_double=None):
        area = tk.Frame(self.body, bg=WHITE, highlightbackground="#d9e5df", highlightthickness=1)
        area.pack(fill="both", expand=True)
        ids = [item[0] for item in columns]
        tree = ttk.Treeview(area, columns=ids, show="headings", height=height, selectmode="browse")
        for key, title, width in columns:
            tree.heading(key, text=title)
            tree.column(key, width=width, minwidth=75, stretch=True)
        vs = ttk.Scrollbar(area, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(area, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        area.columnconfigure(0, weight=1)
        area.rowconfigure(0, weight=1)
        tree.tag_configure("bad", foreground=RED)
        tree.tag_configure("warn", foreground=AMBER)
        for row_id, values, tag in rows:
            tree.insert("", "end", iid=str(row_id), values=[str(v if v is not None else "") for v in values],
                        tags=(tag,) if tag else ())
        if on_double:
            tree.bind("<Double-1>", lambda _: on_double())
        return tree

    def _selection(self, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("请选择记录", "请先在表格中选中一条记录。", parent=self)
            return None
        return int(selected[0])

    def _summary(self, label, value, color=TEAL):
        card = tk.Frame(self.cardbar, bg=WHITE, highlightbackground="#d9e5df", highlightthickness=1)
        card.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(card, text=label, bg=WHITE, fg=MUTED, font=FONT, anchor="w").pack(fill="x", padx=16, pady=(14, 2))
        tk.Label(card, text=str(value), bg=WHITE, fg=color,
                 font=("Microsoft YaHei UI", 25, "bold"), anchor="w").pack(fill="x", padx=16, pady=(0, 14))

    def page_dashboard(self):
        with closing(connect()) as db:
            devices = [equipment_view(db, row) for row in db.execute("SELECT * FROM equipment WHERE deleted=0")]
            month = date.today().strftime("%Y-%m")
            sent = db.execute("SELECT COUNT(DISTINCT equipment_id) FROM submission WHERE sent_date LIKE ?",
                              (month + "-%",)).fetchone()[0]
            calibrated = db.execute("SELECT COUNT(DISTINCT equipment_id) FROM certificate WHERE state='confirmed' AND cal_date LIKE ?",
                                    (month + "-%",)).fetchone()[0]
            waiting = db.execute("SELECT COUNT(*) FROM certificate WHERE state='pending'").fetchone()[0]
            current = [d for d in devices if d["status"] in ("异常", "超期")]
            due_start = add_months(date.today().replace(day=1), 1).isoformat()
            due_end = add_months(date.fromisoformat(due_start), 1).isoformat()
            due = [d for d in devices if d["next_cal_date"] and due_start <= d["next_cal_date"] < due_end]
        self.cardbar = tk.Frame(self.body, bg=PALE)
        self.cardbar.pack(fill="x", pady=(0, 22))
        for label, value, color in (("设备总数", len(devices), TEAL), ("本月已送检", sent, TEAL),
                                    ("本月已校准", calibrated, TEAL), ("异常与超期", len(current), RED),
                                    ("证书待确认", waiting, AMBER)):
            self._summary(label, value, color)
        tk.Label(self.body, text="需关注的设备", bg=PALE, fg=INK,
                 font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w", pady=(0, 10))
        self._table([("code", "资产编号", 135), ("name", "设备名称", 220),
                     ("due", "到期日期", 140), ("status", "计量状态", 110),
                     ("reason", "异常原因", 280)],
                    [(d["id"], (d["asset_code"], d["name"], d["next_cal_date"], d["status"],
                                  d["abnormal_note"] or d["certificate_issue"]), "bad")
                     for d in (current + due)[:300]])
        self.status.set(f"下月到期 {len(due)} 台；异常或超期 {len(current)} 台")

    def page_equipment(self):
        bar = self._toolbar(("新增设备", lambda: self.edit_equipment()),
                            ("编辑设备", lambda: self.edit_equipment(self._selection(self.grid))),
                            ("归档设备", lambda: self.archive_selected()),
                            ("导入 Excel/CSV", self.import_equipment_file),
                            ("导出台账", self.export_equipment))
        ttk.Entry(bar, textvariable=self.search, width=28).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="搜索", command=lambda: self.show("equipment")).pack(side="right")
        query = f"%{self.search.get().strip()}%"
        with closing(connect()) as db:
            devices = [equipment_view(db, row) for row in db.execute("""SELECT * FROM equipment
                WHERE deleted=0 AND (asset_code LIKE ? OR name LIKE ?)
                ORDER BY asset_code LIMIT 5000""", (query, query))]
        self.grid = self._table([("code", "资产编号", 125), ("name", "设备名称", 180),
                                 ("model", "型号", 150), ("category", "分类", 80),
                                 ("department", "部门", 125), ("cycle", "周期（月）", 100),
                                 ("last", "上次校准", 130), ("next", "下次到期", 130),
                                 ("status", "状态", 100)],
                                [(d["id"], (d["asset_code"], d["name"], d["model"], d["category"],
                                            d["department"], d["calibration_cycle"], d["last_cal_date"],
                                            d["next_cal_date"], d["status"]),
                                  "bad" if d["status"] in ("异常", "超期") else "warn" if d["status"] == "待校准" else "")
                                 for d in devices], on_double=lambda: self.edit_equipment(self._selection(self.grid)))
        self.status.set(f"显示 {len(devices)} 台设备；双击可编辑")

    def edit_equipment(self, equipment_id=None):
        old = None
        if equipment_id:
            with closing(connect()) as db:
                old = dict(db.execute("SELECT * FROM equipment WHERE id=?", (equipment_id,)).fetchone())
        fields = [("asset_code", "资产编号 *", None), ("name", "设备名称 *", None),
                  ("model", "型号", None), ("serial_no", "出厂编号", None),
                  ("manufacturer", "生产厂家", None), ("category", "ABC 分类", ("A", "B", "C")),
                  ("department", "使用部门", None), ("location", "存放地点", None),
                  ("calibration_cycle", "检定周期（月）", None),
                  ("stopped", "使用状态 0使用 / 1停用", ("0", "1")),
                  ("abnormal_flag", "异常标记", ("", "计量不过", "异常未检", "其他")),
                  ("abnormal_note", "异常备注", None)]
        initial = old or {"category": "C", "calibration_cycle": "12", "stopped": "0"}

        def save(values):
            if old and old["abnormal_flag"] and not values["abnormal_flag"]:
                reason = self._ask_text("解除异常", "处理原因 *")
                if reason is None:
                    return False
                values["closure_note"] = reason
            return self._call(lambda db: save_equipment(db, values, self._operator(), equipment_id),
                              success="设备已保存") is not None

        Editor(self, "编辑设备" if old else "新建设备", fields, initial, save)

    def _ask_text(self, title, label, default=""):
        from tkinter.simpledialog import askstring
        return askstring(title, label, initialvalue=default, parent=self)

    def archive_selected(self):
        equipment_id = self._selection(self.grid)
        if not equipment_id:
            return
        note = self._ask_text("归档设备", "归档原因 *")
        if note is not None:
            self._call(lambda db: archive_equipment(db, equipment_id, note, self._operator()), "设备已归档")

    def import_equipment_file(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("表格", "*.xlsx *.csv")])
        if not path:
            return
        result = self._call(lambda db: import_equipment(db, payload_for(path), self._operator()),
                            success="台账导入完成")
        if result is not None:
            messagebox.showinfo("导入结果", f"成功 {result['imported']} 行，错误 {len(result['errors'])} 行。\n" +
                                "\n".join(f"第 {e['line']} 行：{e['error']}" for e in result["errors"][:12]), parent=self)

    def _save_file(self, raw, name, extension):
        path = filedialog.asksaveasfilename(parent=self, initialfile=name, defaultextension=extension,
                                            filetypes=[("文件", "*" + extension)])
        if path:
            Path(path).write_bytes(raw)
            self.status.set(f"已保存到 {path}")

    def export_equipment(self):
        with closing(connect()) as db:
            devices = [equipment_view(db, row) for row in db.execute("SELECT * FROM equipment WHERE deleted=0 ORDER BY asset_code")]
        columns = ("asset_code", "name", "model", "serial_no", "manufacturer", "category",
                   "department", "location", "calibration_cycle", "last_cal_date",
                   "next_cal_date", "status", "abnormal_flag", "abnormal_note")
        headers = ("资产编号", "设备名称", "型号", "出厂编号", "厂家", "分类", "部门", "地点",
                   "周期（月）", "上次校准", "下次到期", "状态", "异常", "备注")
        self._save_file(workbook_bytes(headers, [[item.get(key) for key in columns] for item in devices]),
                        "设备台账.xlsx", ".xlsx")

    def page_submissions(self):
        self._toolbar(("登记送检", self.new_submission),
                      ("登记返回", lambda: self.mark_return(self._selection(self.grid))))
        with closing(connect()) as db:
            rows = list(db.execute("""SELECT s.*,e.asset_code,e.name FROM submission s
                JOIN equipment e ON e.id=s.equipment_id ORDER BY s.id DESC LIMIT 5000"""))
        self.grid = self._table([("code", "资产编号", 140), ("name", "设备名称", 220),
                                 ("sent", "送检日期", 150), ("returned", "返回日期", 150),
                                 ("note", "备注", 350)],
                                [(r["id"], (r["asset_code"], r["name"], r["sent_date"],
                                             r["returned_date"], r["note"]),
                                  "warn" if not r["returned_date"] else "") for r in rows],
                                on_double=lambda: self.mark_return(self._selection(self.grid)))
        self.status.set(f"送检记录 {len(rows)} 条")

    def new_submission(self):
        fields = [("asset_code", "资产编号 *", None), ("sent_date", "送检日期 YYYY-MM-DD *", None),
                  ("note", "备注", None)]
        Editor(self, "登记送检", fields, {"sent_date": date.today().isoformat()},
               lambda values: self._call(lambda db: save_submission(db, values, self._operator()),
                                         "送检已登记") is not None)

    def mark_return(self, record_id):
        if not record_id:
            return
        value = self._ask_text("登记返回", "返回日期 YYYY-MM-DD *", date.today().isoformat())
        if value is not None:
            self._call(lambda db: return_submission(db, record_id, {"returned_date": value},
                                                    self._operator()), "返回已登记")

    def page_certificates(self):
        self._toolbar(("导入 PDF", self.import_certificates),
                      ("审核 / 更正", lambda: self.edit_certificate(self._selection(self.grid))),
                      ("打开原件", lambda: self.open_certificate(self._selection(self.grid))))
        with closing(connect()) as db:
            rows = list(db.execute("""SELECT c.*,e.asset_code FROM certificate c
                LEFT JOIN equipment e ON e.id=c.equipment_id ORDER BY c.id DESC LIMIT 5000"""))
        self.grid = self._table([("name", "文件名", 230), ("cert", "证书编号", 160),
                                 ("code", "资产编号", 150), ("date", "校准日期", 135),
                                 ("result", "结论", 100), ("state", "审核状态", 110),
                                 ("note", "提示", 280)],
                                [(r["id"], (r["file_name"], r["cert_no"], r["asset_code"] or
                                              json.loads(r["parsed_data"]).get("asset_code", ""),
                                              r["cal_date"], r["result"],
                                              {"pending": "待确认", "confirmed": "已确认", "rejected": "已驳回"}[r["state"]],
                                              r["review_note"]),
                                  "warn" if r["state"] == "pending" else "bad" if r["result"] == "不合格" else "")
                                 for r in rows],
                                on_double=lambda: self.edit_certificate(self._selection(self.grid)))
        self.status.set(f"证书 {len(rows)} 份；待确认 {sum(r['state']=='pending' for r in rows)} 份")

    def import_certificates(self):
        paths = filedialog.askopenfilenames(parent=self, filetypes=[("PDF 证书", "*.pdf")])
        if not paths:
            return
        confirmed = pending = 0
        errors = []
        for path in paths:
            result = self._call(lambda db, p=path: upload_pdf(db, payload_for(p), self._operator()),
                                refresh=False)
            if result is None:
                errors.append(Path(path).name)
            elif result["state"] == "confirmed":
                confirmed += 1
            else:
                pending += 1
        self.show("certificates")
        messagebox.showinfo("导入结果", f"已自动关联 {confirmed} 份，待确认 {pending} 份，失败 {len(errors)} 份。" +
                            ("\n失败文件：" + "、".join(errors[:10]) if errors else ""), parent=self)

    def edit_certificate(self, certificate_id):
        if not certificate_id:
            return
        with closing(connect()) as db:
            row = db.execute("""SELECT c.*,e.asset_code FROM certificate c LEFT JOIN equipment e
                ON e.id=c.equipment_id WHERE c.id=?""", (certificate_id,)).fetchone()
            if not row:
                return
            cert = dict(row)
        initial = {**cert, "asset_code": cert["asset_code"] or
                   json.loads(cert["parsed_data"]).get("asset_code", "")}
        fields = [("asset_code", "资产编号 *", None), ("cert_no", "证书编号 *", None),
                  ("cal_org", "校准机构", None), ("cal_date", "校准日期 YYYY-MM-DD *", None),
                  ("valid_until", "有效期至 YYYY-MM-DD", None),
                  ("result", "结论 *", ("", "合格", "不合格")),
                  ("agreed_amount", "本次约定金额", None), ("review_note", "审核备注", None)]
        Editor(self, f"确认 / 更正证书 #{certificate_id}", fields, initial,
               lambda values: self._call(lambda db: review_certificate(db, certificate_id,
                   {**values, "state": "confirmed"}, self._operator()), "证书已确认") is not None)

    def open_certificate(self, certificate_id):
        if not certificate_id:
            return
        with closing(connect()) as db:
            row = db.execute("SELECT file_path FROM certificate WHERE id=?", (certificate_id,)).fetchone()
        if row and Path(row["file_path"]).exists():
            os.startfile(row["file_path"])
        else:
            messagebox.showerror("文件不存在", "证书原件未找到，请检查备份或存储目录。", parent=self)

    def page_billing(self):
        bar = self._toolbar(("导入清单并核对", self.import_billing_file),
                            ("导出差异报告", self.export_billing))
        tk.Label(bar, text="核对校准月份（可选）", bg=PALE, fg=MUTED, font=FONT).pack(side="left", padx=(18, 6))
        if not hasattr(self, "billing_month"):
            self.billing_month = tk.StringVar(value=date.today().strftime("%Y-%m"))
        ttk.Entry(bar, textvariable=self.billing_month, width=10).pack(side="left")
        with closing(connect()) as db:
            rows = list(db.execute("SELECT * FROM billing_record ORDER BY id DESC LIMIT 5000"))
        self.grid = self._table([("batch", "导入批次", 185), ("line", "原始行", 90),
                                 ("code", "资产编号", 145), ("cert", "证书编号", 150),
                                 ("amount", "金额", 105), ("finding", "核对结果", 115),
                                 ("detail", "详情", 315)],
                                [(r["id"], (r["batch_no"], r["line_no"], r["asset_code"],
                                             r["cert_no"], r["amount"], r["finding"], r["detail"]),
                                  "" if r["finding"] == "匹配" else "warn" if r["finding"] == "金额待核对" else "bad")
                                 for r in rows])
        self.status.set(f"对账记录 {len(rows)} 条；金额差异需录入证书对应服务的约定金额")

    def import_billing_file(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("表格", "*.xlsx *.csv")])
        if not path:
            return
        data = {**payload_for(path), "calibration_month": self.billing_month.get().strip()}
        result = self._call(lambda db: import_billing(db, data, self._operator()), "对账完成")
        if result is not None:
            unmatched = sum(row["finding"] != "匹配" for row in result["findings"])
            messagebox.showinfo("核对结果", f"清单 {result['rows']} 行，差异或待核对 {unmatched} 条。",
                                parent=self)

    def export_billing(self):
        with closing(connect()) as db:
            rows = list(db.execute("""SELECT batch_no,line_no,asset_code,cert_no,amount,finding,detail
                FROM billing_record ORDER BY id"""))
        self._save_file(workbook_bytes(("批次", "行号", "资产编号", "证书编号", "金额", "结果", "详情"), rows),
                        "对账差异.xlsx", ".xlsx")

    def page_audit(self):
        self._toolbar(("导出审计", self.export_audit),
                      ("下载完整备份", self.export_backup), ("从备份恢复", self.restore_backup))
        with closing(connect()) as db:
            rows = list(db.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 5000"))
        self.grid = self._table([("time", "操作时间", 210), ("entity", "业务对象", 135),
                                 ("id", "记录 ID", 95), ("action", "操作", 130),
                                 ("operator", "操作人", 135), ("detail", "变更摘要", 400)],
                                [(r["id"], (r["timestamp"], r["entity"], r["entity_id"],
                                             r["action"], r["operator"], r["new_data"][:180]), "") for r in rows])
        self.status.set(f"审计记录 {len(rows)} 条；数据目录：{DATA}")

    def export_audit(self):
        with closing(connect()) as db:
            rows = list(db.execute("""SELECT timestamp,entity,entity_id,action,operator,old_data,new_data
                FROM audit_log ORDER BY id"""))
        self._save_file(workbook_bytes(("时间", "实体", "ID", "操作", "操作人", "原值", "新值"), rows),
                        "审计记录.xlsx", ".xlsx")

    def export_backup(self):
        try:
            self._save_file(backup_bytes(), "计量台账备份.zip", ".zip")
        except (OSError, sqlite3.Error) as exc:
            messagebox.showerror("备份失败", str(exc), parent=self)

    def restore_backup(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("备份 ZIP", "*.zip")])
        if not path:
            return
        if not messagebox.askyesno("恢复备份", "将用所选备份替换当前数据。程序会先保存当前数据的恢复包。继续吗？", parent=self):
            return
        try:
            recovery = restore(path, guard_server=False)
            init_db()
            self.show("dashboard")
            messagebox.showinfo("恢复完成", f"原有数据恢复包：{recovery}", parent=self)
        except (OSError, ValueError, sqlite3.Error) as exc:
            messagebox.showerror("恢复失败", str(exc), parent=self)


def main():
    DesktopApp().mainloop()


if __name__ == "__main__":
    main()
