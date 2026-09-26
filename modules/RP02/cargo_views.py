"""RP02 Backdated Cargo Handling routes."""
import json as _json
from functools import wraps
from flask import render_template, request, jsonify, session, redirect, url_for, Response
from database import get_user_permissions
from . import bp
from . import cargo_handling


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), 'RP02')


def _can_upload():
    perms = get_perms()
    return bool(perms.get('can_add') or perms.get('can_edit'))


def read_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if not get_perms().get('can_read'):
            return render_template('no_access.html'), 403
        return f(*args, **kwargs)
    return decorated


def upload_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Not logged in'}), 401
        if not _can_upload():
            return jsonify({'error': 'No upload permission'}), 403
        return f(*args, **kwargs)
    return decorated


@bp.route('/module/RP02/cargo-backdated/', strict_slashes=False)
@bp.route('/module/RP02/cargo-backdated', strict_slashes=False)
@read_required
def cargo_backdated_index():
    return render_template(
        'cargo_backdated.html',
        username=session.get('username'),
        status=cargo_handling.get_status(),
        can_upload=_can_upload(),
        fields=cargo_handling.FIELDS
    )


@bp.route('/api/module/RP02/cargo-backdated/template', strict_slashes=False)
@read_required
def cargo_backdated_template():
    fmt = request.args.get('format', 'excel').lower()
    if fmt == 'csv':
        return Response(
            cargo_handling.build_template_csv(),
            mimetype='text/csv',
            headers={'Content-Disposition': 'attachment; filename="RP02_cargo_handling_template.csv"',
                     'Cache-Control': 'no-store'}
        )
    return Response(
        cargo_handling.build_template_excel(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': 'attachment; filename="RP02_cargo_handling_template.xlsx"',
                 'Cache-Control': 'no-store'}
    )


@bp.route('/api/module/RP02/cargo-backdated/preview', methods=['POST'], strict_slashes=False)
@upload_required
def cargo_backdated_preview():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    rows, errors = cargo_handling.parse_upload(f)
    return jsonify({
        'total_rows': len(rows),
        'format_errors': errors,
        'sample': rows[:5]
    })


@bp.route('/api/module/RP02/cargo-backdated/apply', methods=['POST'], strict_slashes=False)
@upload_required
def cargo_backdated_apply():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    rows, errors = cargo_handling.parse_upload(f)
    if errors:
        return jsonify({'error': 'Fix format errors before applying',
                        'format_errors': errors}), 400
    mode = request.form.get('mode', 'replace')
    user = session.get('username') or session.get('user_id') or 'admin'
    if mode == 'append':
        inserted = cargo_handling.append_rows(rows, user)
    else:
        inserted = cargo_handling.replace_all(rows, user)
    return jsonify({'inserted': inserted, 'mode': mode})


@bp.route('/api/module/RP02/cargo-backdated/rows', strict_slashes=False)
@read_required
def cargo_backdated_rows():
    try:
        page = int(request.args.get('page', 1))
        size = int(request.args.get('size', 50))
    except (TypeError, ValueError):
        page, size = 1, 50
    try:
        filters = _json.loads(request.args.get('colfilters') or '[]')
        if not isinstance(filters, list):
            filters = []
    except (ValueError, TypeError):
        filters = []
    rows, total = cargo_handling.get_rows(page, size, filters)
    return jsonify({
        'data': rows,
        'last_page': max(1, (total + size - 1) // size),
        'total': total
    })


@bp.route('/api/module/RP02/cargo-backdated/row/update', methods=['POST'], strict_slashes=False)
@upload_required
def cargo_backdated_row_update():
    data = request.json or {}
    if not data.get('id'):
        return jsonify({'error': 'Missing id'}), 400
    res = cargo_handling.update_row(data['id'], data)
    if res.get('error'):
        return jsonify(res), 400
    return jsonify(res)


@bp.route('/api/module/RP02/cargo-backdated/row/delete', methods=['POST'], strict_slashes=False)
@login_required
def cargo_backdated_row_delete():
    if not get_perms().get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    data = request.json or {}
    if not data.get('id'):
        return jsonify({'error': 'Missing id'}), 400
    cargo_handling.delete_row(data['id'])
    return jsonify({'success': True})
