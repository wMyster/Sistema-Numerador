"""
Lock de sessão em SQLite (transacional).
"""
from __future__ import annotations

import os
import sqlite3
import threading
import time
from datetime import datetime, timedelta

import settings

def _connect_lock_db() -> sqlite3.Connection | None:
    """Conecta ao banco atual usado pelo sistema para guardar sessões ativas."""
    if not settings.get_rede_path():
        return None

    con = sqlite3.connect(settings.get_db_path(), timeout=10)
    con.row_factory = sqlite3.Row
    return con

def _ensure_session_table(con: sqlite3.Connection) -> None:
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS active_sessions (
            user_name TEXT PRIMARY KEY,
            status TEXT NOT NULL,
            last_ping TEXT NOT NULL,
            host_name TEXT,
            pid INTEGER
        )
        """
    )

def _upsert_session(con: sqlite3.Connection, user_name: str, status: str) -> None:
    now_iso = datetime.now().isoformat()
    con.execute(
        """
        INSERT INTO active_sessions(user_name, status, last_ping, host_name, pid)
        VALUES(?, ?, ?, ?, ?)
        ON CONFLICT(user_name) DO UPDATE SET
            status = excluded.status,
            last_ping = excluded.last_ping,
            host_name = excluded.host_name,
            pid = excluded.pid
        """,
        (user_name, status, now_iso, os.environ.get("COMPUTERNAME", ""), os.getpid()),
    )

def _is_session_active(row: sqlite3.Row, timeout_seconds: int) -> bool:
    if not row:
        return False
    if (row["status"] or "").lower() != "ocupado":
        return False
    try:
        last_ping = datetime.fromisoformat(row["last_ping"])
    except Exception:
        return False
    return datetime.now() - last_ping < timedelta(seconds=timeout_seconds)


def acquire_lock(user_name: str, timeout_seconds: int = 25) -> tuple[bool, str]:
    """
    Tenta marcar o usuário como ocupado de forma transacional.
    """
    con = _connect_lock_db()
    if con is None:
        return True, ""

    try:
        con.execute("BEGIN IMMEDIATE")
        _ensure_session_table(con)

        row = con.execute("SELECT * FROM active_sessions WHERE user_name = ?", (user_name,)).fetchone()
        if row and _is_session_active(row, timeout_seconds):
            con.rollback()
            return (
                False,
                f"O usuário '{user_name}' já está logado no Sistema em outro computador.\n\n"
                f"Aguarde ele fechar o programa (ou espere {timeout_seconds}s se houve queda) para entrar.",
            )

        _upsert_session(con, user_name, "ocupado")
        con.commit()
        return True, ""
    except Exception as e:
        try:
            con.rollback()
        except Exception:
            pass
        return False, f"Falha ao registrar login atual: {e}"
    finally:
        con.close()


def force_lock(user_name: str, data: dict | None = None) -> tuple[bool, str]:
    _ = data
    con = _connect_lock_db()
    if con is None:
        return True, ""

    try:
        con.execute("BEGIN IMMEDIATE")
        _ensure_session_table(con)
        _upsert_session(con, user_name, "ocupado")
        con.commit()
        return True, ""
    except Exception as e:
        try:
            con.rollback()
        except Exception:
            pass
        return False, f"Falha ao registrar login atual: {e}"
    finally:
        con.close()


def release_lock(user_name: str) -> None:
    """Marca o usuário como livre."""
    if not user_name:
        return

    con = _connect_lock_db()
    if con is None:
        return

    try:
        con.execute("BEGIN IMMEDIATE")
        _ensure_session_table(con)
        con.execute(
            """
            UPDATE active_sessions
            SET status = 'livre', last_ping = ?
            WHERE user_name = ?
            """,
            (datetime.now().isoformat(), user_name),
        )
        con.commit()
    except Exception:
        try:
            con.rollback()
        except Exception:
            pass
    finally:
        con.close()


def _heartbeat_worker(user_name: str, interval: int = 10) -> None:
    """Atualiza periodicamente o status ocupado do usuário logado."""
    while True:
        con = _connect_lock_db()
        if con is not None:
            try:
                con.execute("BEGIN IMMEDIATE")
                _ensure_session_table(con)
                _upsert_session(con, user_name, "ocupado")
                con.commit()
            except Exception:
                try:
                    con.rollback()
                except Exception:
                    pass
            finally:
                con.close()
        time.sleep(interval)


def start_lock_heartbeat(user_name: str) -> None:
    """Inicia thread daemon de heartbeat."""
    t = threading.Thread(target=_heartbeat_worker, args=(user_name,), daemon=True)
    t.start()
