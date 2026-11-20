"""data_audit trigger: records who changed what, skips no-op saves, and never
copies secrets or file bytes. Applies migration dataaudit001 inside one
transaction on the local DB and rolls everything back — nothing persists.
Run: python test_data_audit.py"""
import importlib.util
import psycopg2.extras
from database import get_db

spec = importlib.util.spec_from_file_location(
    'mig', 'alembic/versions/dataaudit001_row_level_data_audit.py')
mig = importlib.util.module_from_spec(spec)


class Op:
    def __init__(self, cur):
        self.cur = cur

    def execute(self, sql):
        self.cur.execute(sql)


def audit_rows(cur, table):
    cur.execute("SELECT * FROM data_audit WHERE table_name=%s ORDER BY id", (table,))
    return cur.fetchall()


def main():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    try:
        spec.loader.exec_module(mig)
        mig.op = Op(cur)
        mig.upgrade()
        cur.execute("SELECT set_config('app.user', 'tester', false)")

        # Edit → one row with only the changed field, full old row, the app user.
        cur.execute("SELECT id, material_po FROM mbc_customer_details ORDER BY id LIMIT 1")
        line = cur.fetchone()
        assert line, 'needs at least one mbc_customer_details row'
        cur.execute("UPDATE mbc_customer_details SET material_po=material_po WHERE id=%s", (line['id'],))
        assert audit_rows(cur, 'mbc_customer_details') == [], 'no-op save must not be logged'
        cur.execute("UPDATE mbc_customer_details SET material_po='PO-TEST' WHERE id=%s", (line['id'],))
        [r] = audit_rows(cur, 'mbc_customer_details')
        assert r['op'] == 'U' and r['row_id'] == str(line['id']) and r['changed_by'] == 'tester'
        assert r['new_data'] == {'material_po': 'PO-TEST'}
        assert r['old_data']['material_po'] == line['material_po'] and 'mbc_id' in r['old_data']

        # Secrets: a password-only change leaves nothing to log; password never stored.
        cur.execute("SELECT id FROM users ORDER BY id LIMIT 1")
        uid = cur.fetchone()['id']
        cur.execute("UPDATE users SET password=password || 'x' WHERE id=%s", (uid,))
        assert audit_rows(cur, 'users') == [], 'password-only change must not be logged'

        # Binaries: proof-doc metadata audited, file bytes stripped.
        cur.execute("SELECT id FROM mbc_proof_documents ORDER BY id LIMIT 1")
        doc = cur.fetchone()
        if doc:
            cur.execute("DELETE FROM mbc_proof_documents WHERE id=%s", (doc['id'],))
            [d] = audit_rows(cur, 'mbc_proof_documents')
            assert d['op'] == 'D' and 'file_bytes' not in d['old_data']

        # Excluded tables get no trigger.
        cur.execute("SELECT 1 FROM information_schema.triggers "
                    "WHERE trigger_name='data_audit' AND event_object_table='approval_log'")
        assert cur.fetchone() is None

        # AUD01 formatting of the recorded edit.
        from modules.AUD01.views import _data_rows
        from datetime import datetime
        rows = [x for x in _data_rows(cur, datetime(2000, 1, 1), datetime(2100, 1, 1), 50)
                if x['module_code'] == 'mbc_customer_details']
        assert rows[0]['action'] == 'Edited' and 'material_po:' in rows[0]['comment']
        assert 'PO-TEST' in rows[0]['comment'] and 'mbc_id' in rows[0]['ref']
    finally:
        conn.rollback()
        conn.close()
    print('data audit: OK')


if __name__ == '__main__':
    main()
