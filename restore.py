"""Restore a local backup while the server is stopped."""

import hashlib
import io
import os
import socket
import sqlite3
import sys
import zipfile
from contextlib import closing
from datetime import datetime
from pathlib import Path

from server import DATA, DB_FILE, PDF_DIR, backup_bytes


def restore(archive_path, db_path=DB_FILE, pdf_dir=PDF_DIR, guard_server=True):
    if guard_server:
        with socket.socket() as sock:
            sock.settimeout(0.3)
            if sock.connect_ex(("127.0.0.1", 8765)) == 0:
                raise ValueError("请先关闭运行中的服务，再恢复备份")
    db_path.parent.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path) as bundle:
        entries = set(bundle.namelist())
        if "register.sqlite3" not in entries or any(
            name != "register.sqlite3" and not (
                name.startswith("certificates/") and name.count("/") == 1
                and name.endswith(".pdf") and len(Path(name).stem) == 64
            ) for name in entries
        ):
            raise ValueError("备份内容不符合本软件格式")
        staged = db_path.parent / "_restore_source.sqlite3"
        staged.write_bytes(bundle.read("register.sqlite3"))
        try:
            with closing(sqlite3.connect(staged)) as db, db:
                if db.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise ValueError("备份数据库损坏")
                certificates = db.execute("SELECT id,cert_hash FROM certificate").fetchall()
                needed = {f"certificates/{digest}.pdf" for _, digest in certificates}
                if not needed.issubset(entries):
                    raise ValueError("备份缺少证书文件")
                for _, digest in certificates:
                    raw = bundle.read(f"certificates/{digest}.pdf")
                    if hashlib.sha256(raw).hexdigest() != digest:
                        raise ValueError("证书哈希校验失败")
                for cert_id, digest in certificates:
                    db.execute("UPDATE certificate SET file_path=? WHERE id=?",
                               (str(pdf_dir / f"{digest}.pdf"), cert_id))
            recovery = None
            if db_path.exists():
                recovery = db_path.parent / f"pre-restore-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
                recovery.write_bytes(backup_bytes(db_path, pdf_dir))
            for _, digest in certificates:
                (pdf_dir / f"{digest}.pdf").write_bytes(bundle.read(f"certificates/{digest}.pdf"))
            os.replace(staged, db_path)
            return recovery
        finally:
            staged.unlink(missing_ok=True)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("用法：python restore.py <备份.zip>（先关闭服务）")
    result = restore(Path(sys.argv[1]))
    print("恢复完成。原有数据的恢复包：", result or "无")
