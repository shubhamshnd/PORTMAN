from flask import Blueprint

MODULE_INFO = {
    'code': 'RP02',
    'name': 'Finance Reports'
}

bp = Blueprint('RP02', __name__, template_folder='.')

from . import views
from . import cargo_views
