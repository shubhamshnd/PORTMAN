import json
from contextlib import contextmanager
import psycopg2
import psycopg2.extras
from config import DATABASE_URL

MODULE_TABLE_REGISTRY = {
    'TM01': 'tide_master',
}


def _request_user():
    """Logged-in username for the current Flask request, else None
    (background workers, scripts, CLI)."""
    try:
        from flask import has_request_context, session
        return session.get('username') if has_request_context() else None
    except Exception:
        return None


def get_db():
    conn = psycopg2.connect(DATABASE_URL)
    # The data_audit trigger reads app.user to record who made each change.
    # Committed straight away so a later rollback on this connection can't undo it.
    user = _request_user()
    if user:
        with conn.cursor() as cur:
            cur.execute("SELECT set_config('app.user', %s, false)", (str(user),))
        conn.commit()
    return conn


def get_cursor(conn):
    """Return a RealDictCursor for dict-like row access."""
    return conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)


def get_module_table(module_code):
    return MODULE_TABLE_REGISTRY.get(module_code)


@contextmanager
def get_db_connection():
    conn = get_db()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_user_permissions(user_id, module_code):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT can_read, can_add, can_edit, can_delete
        FROM module_permissions
        WHERE user_id = %s AND module_code = %s
    ''', [user_id, module_code])
    row = cur.fetchone()
    conn.close()
    if row:
        return {'can_read': row['can_read'], 'can_add': row['can_add'],
                'can_edit': row['can_edit'], 'can_delete': row['can_delete']}
    return {'can_read': 0, 'can_add': 0, 'can_edit': 0, 'can_delete': 0}


def get_module_config(module_code):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT config_json FROM module_config WHERE module_code = %s', [module_code])
    row = cur.fetchone()
    conn.close()
    if row:
        return json.loads(row['config_json'])
    return {}


def save_module_config(module_code, config):
    conn = get_db()
    cur = get_cursor(conn)
    config_json = json.dumps(config)
    cur.execute('''
        INSERT INTO module_config (module_code, config_json) VALUES (%s, %s)
        ON CONFLICT(module_code) DO UPDATE SET config_json = %s
    ''', [module_code, config_json, config_json])
    conn.commit()
    conn.close()
