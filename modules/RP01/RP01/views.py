from flask import render_template, session, redirect, url_for
from functools import wraps
from . import bp
from .vessel_sof import views as _vessel_sof_views  # noqa: registers vessel-sof routes on bp
from .mbc_sof import views as _mbc_sof_views        # noqa: registers mbc-sof routes on bp
from .mbc_master import views as _mbc_master_views  # noqa: registers mbc-master routes on bp
from .mbc_tat    import views as _mbc_tat_views     # noqa: registers mbc-tat routes on bp
from .vessel_discharged import views as _vessel_discharged_views  # noqa: registers vessel-discharged routes on bp
from .custom_report    import views as _custom_report_views       # noqa: registers custom-report routes on bp
from .dashboard        import views as _dashboard_views           # noqa: registers dashboard routes on bp
from .daily_ops         import views as _daily_ops_views        # noqa: registers daily-ops routes on bp
from .shift_report      import views as _shift_report_views    # noqa: registers shift-report routes on bp
from .barge_report      import views as _barge_report_views    # noqa: registers barge-report routes on bp
from .mv_mbc_month_report import views as _mv_monthly_report_views  # noqa: registers mv-monthly-report routes on bp
from .live_dashboard    import views as _live_dashboard_views  # noqa: registers live-dashboard routes on bp
from .mbc_tracking      import views as _mbc_tracking_views    # noqa: registers mbc-tracking routes on bp
from .port_map          import views as _port_map_views        # noqa: registers port-map routes on bp
from .daily_barge_report import views as _mv_barge_report_views  # noqa: registers mv-barge-report routes on bp
from .cargo_report import views as _cargo_report_views
from .cargo_statistics import cargo_statistics as _cargo_statistics_views
from .historical_data import views as _historical_data_views  # noqa: registers historical-data routes on bp
from .Barge_Position_Report  import views as _barge_position_views

from .report_24_hours import views as _report_24_hours_views
from .daily_progress_report import views as _daily_progress_report_views
from .two_hour_report import views as _two_hour_report_views  # noqa: registers two-hour-report routes on bp
from .port_overview import views as _port_overview_views  # noqa: registers port-overview routes on bp
from .ai_chat import views as _ai_chat_views  # noqa: registers ai-chat routes on bp
from .shift_cargo_balance import views as _shift_cargo_balance_views  # noqa: registers shift-cargo-balance routes on bp

from database import get_user_permissions

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def _get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), 'RP01')

@bp.route('/module/RP01/')
@login_required
def index():
    perms = _get_perms()
    if not perms or not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('rp01.html', username=session.get('username'), permissions=perms)
