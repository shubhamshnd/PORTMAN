"""Billing lock: an operational doc freezes once its cargo is billed, and
permanently once invoiced. Drives source_lock_state with a fake cursor that
answers each probe in order. Run: python test_billing_lock.py"""
from modules.FIN01.model import source_lock_state, lock_message


class FakeCursor:
    """Answers each execute() with the next scripted row. A row of None means
    'no match'. Records the SQL so we can assert which probe ran."""

    def __init__(self, answers):
        self.answers = list(answers)
        self.sql = []
        self.params = []
        self._row = None

    def execute(self, sql, params=None):
        self.sql.append(' '.join(sql.split()))
        self.params.append(params)
        self._row = self.answers.pop(0) if self.answers else None

    def fetchone(self):
        return self._row


def test_unbilled_is_open():
    # MBC probes: 1 invoiced check, 1 billed check — both miss.
    cur = FakeCursor([None, None])
    assert source_lock_state(cur, 'MBC', 42) == ''
    assert lock_message('') == ''


def test_billed_blocks():
    cur = FakeCursor([None, {'?column?': 1}])
    assert source_lock_state(cur, 'MBC', 42) == 'billed'
    assert 'Remove or reject the bill' in lock_message('billed', 'MBC')


def test_invoiced_blocks_permanently():
    # First probe hits, so the billed probe never runs.
    cur = FakeCursor([{'?column?': 1}])
    assert source_lock_state(cur, 'MBC', 42) == 'invoiced'
    assert 'permanent' in lock_message('invoiced', 'MBC')
    assert 'invoice_bill_mapping' in cur.sql[0]
    assert len(cur.sql) == 1


def test_invoiced_ignores_bill_status():
    # The invoiced probe must not filter on bill_status — a cancelled bill
    # still means a customer was once invoiced off this cargo.
    cur = FakeCursor([{'?column?': 1}])
    source_lock_state(cur, 'MBC', 42)
    assert 'bill_status' not in cur.sql[0]


def test_vcn_checks_both_import_and_export():
    # VCN has two declaration tables; an unbilled call probes all four times.
    cur = FakeCursor([None, None, None, None])
    assert source_lock_state(cur, 'VCN', 7) == ''
    joined = ' '.join(cur.sql)
    assert 'vcn_cargo_declaration' in joined
    assert 'vcn_export_cargo_declaration' in joined


def test_export_only_billing_still_locks():
    # Import clean, export billed → locked.
    cur = FakeCursor([None, None, None, {'?column?': 1}])
    assert source_lock_state(cur, 'VCN', 7) == 'billed'


def test_unknown_kind_and_missing_id_are_open():
    assert source_lock_state(FakeCursor([]), 'BARGE', 1) == ''
    assert source_lock_state(FakeCursor([]), 'MBC', None) == ''


def test_line_scope_narrows_to_one_row():
    # MBC01 Change Customer: a billed sibling must not lock this line, so both
    # probes filter on the line id and bind it right after the mbc_id.
    cur = FakeCursor([None, None])
    assert source_lock_state(cur, 'MBC', 42, line_id=9) == ''
    for sql, params in zip(cur.sql, cur.params):
        assert 'd.id = %s' in sql
        assert params[1:3] == [42, 9]
    assert sql.count('%s') == len(params)


def test_released_line_is_open_despite_old_invoice():
    # Admin → Release Cargo zeroes the counters but leaves the old invoice rows;
    # FIN01 offers it for billing again, so MBC01 must not call it invoiced.
    from modules.MBC01.model import _line_lock
    cur = FakeCursor([{'?column?': 1}])     # the old invoice link still exists
    released = {'id': 9, 'mbc_id': 42, 'is_billed': 0, 'billed_quantity': 0, 'bill_id': None}
    assert _line_lock(cur, released) == ''
    assert cur.sql == []                    # counters decide; no probe runs
    held = dict(released, is_billed=1, billed_quantity=7816, bill_id=5)
    assert _line_lock(FakeCursor([{'?column?': 1}]), held) == 'invoiced'
    assert _line_lock(FakeCursor([None, None]), held) == 'billed'   # cutover mark


def test_change_line_reports_only_what_moved():
    from modules.MBC01.model import line_changes
    line = {'customer_name': 'JSW STEEL LIMITED', 'material_po': '4300025732'}
    assert line_changes(line, 'JSW STEEL LIMITED', '4300025732') == ('', [])
    action, changes = line_changes(line, 'JSW STEEL LIMITED', '4500256823')
    assert action == 'PO Changed' and changes == ['PO: 4300025732 → 4500256823']
    action, _ = line_changes(line, 'OTHER', '4300025732')
    assert action == 'Customer Changed'
    action, changes = line_changes(line, 'OTHER', '4500256823')
    assert action == 'Customer / PO Changed' and len(changes) == 2
    assert line_changes({'customer_name': None, 'material_po': None}, 'X', 'Y')[0] == 'Customer / PO Changed'


def test_line_scope_refused_for_vcn():
    try:
        source_lock_state(FakeCursor([]), 'VCN', 7, line_id=1)
    except AssertionError:
        return
    raise AssertionError('VCN line scope should be refused')


if __name__ == '__main__':
    for name, fn in sorted(globals().items()):
        if name.startswith('test_'):
            fn()
    print('billing lock: OK')
