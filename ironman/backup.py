# -*- coding: utf-8 -*-
"""Otomatik ve elle veritabanı yedekleme."""
from __future__ import annotations

import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

from .config import BACKUP_DIR, BACKUP_KEEP, DB_PATH


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def create_backup(db_path: Path | str = DB_PATH, *, tag: str = "auto") -> Path | None:
    """WAL güvenli yedek alır (sqlite backup API'si)."""
    db_path = Path(db_path)
    if not db_path.exists():
        return None
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    target = BACKUP_DIR / f"ironman_{_stamp()}_{tag}.db"
    src = sqlite3.connect(str(db_path))
    try:
        dst = sqlite3.connect(str(target))
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    prune()
    return target


def list_backups() -> list[dict]:
    if not BACKUP_DIR.exists():
        return []
    out = []
    for p in sorted(BACKUP_DIR.glob("ironman_*.db"), reverse=True):
        stat = p.stat()
        out.append({
            "name": p.name,
            "path": str(p),
            "size_kb": round(stat.st_size / 1024, 1),
            "created": datetime.fromtimestamp(stat.st_mtime),
        })
    return out


def prune(keep: int = BACKUP_KEEP) -> int:
    backups = sorted(BACKUP_DIR.glob("ironman_*.db"), reverse=True) if BACKUP_DIR.exists() else []
    removed = 0
    for old in backups[keep:]:
        try:
            old.unlink()
            removed += 1
        except OSError:
            pass
    return removed


def restore_backup(name: str, db_path: Path | str = DB_PATH) -> bool:
    """Geri yüklemeden önce mevcut halin bir yedeğini alır."""
    source = BACKUP_DIR / Path(name).name
    if not source.exists() or source.parent.resolve() != BACKUP_DIR.resolve():
        return False
    create_backup(db_path, tag="restore-oncesi")
    db_path = Path(db_path)
    for suffix in ("-wal", "-shm"):
        side = Path(str(db_path) + suffix)
        if side.exists():
            side.unlink()
    shutil.copy2(source, db_path)
    return True


def maybe_auto_backup(enabled: bool = True) -> Path | None:
    """Uygulama açılışında günde bir kez yedek alır."""
    if not enabled:
        return None
    today = datetime.now().strftime("%Y%m%d")
    if BACKUP_DIR.exists() and any(BACKUP_DIR.glob(f"ironman_{today}_*_auto.db")):
        return None
    return create_backup(tag="auto")
