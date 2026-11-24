"""RP03 ETA learner: logs predictions every 15 min and retrains weekly at night.

Standalone scheduler installed as a Windows service via NSSM, like weather_service:

    nssm install PortmanEta "<python.exe>" "d:\\PMS\\DPPL\\PORTMAN\\eta_service.py"
    nssm set PortmanEta AppDirectory d:\\PMS\\DPPL\\PORTMAN

One-off:  python eta_service.py --once      log current predictions and exit
          python eta_service.py --retrain   retrain now (same checks) and exit
"""
import os
import sys
import time

os.environ.setdefault('OMP_NUM_THREADS', '2')   # leave the CPU to the web app; set before scikit-learn loads

import logging
from datetime import datetime

from modules.RP03 import eta_learn

POLL_SECONDS = 15 * 60
RETRAIN_HOURS = range(1, 5)     # 01:00-04:59, the quiet part of the ops day


def main():
    logging.basicConfig(level=logging.INFO, stream=sys.stdout,
                        format='%(asctime)s %(levelname)s %(message)s')
    if '--once' in sys.argv:
        logging.info('%d predictions logged', eta_learn.log_predictions())
        return
    if '--retrain' in sys.argv:
        promoted, detail = eta_learn.learn()
        logging.info('%s: %s', 'promoted' if promoted else 'kept current model', detail['reason'])
        return
    while True:
        try:
            logging.info('%d predictions logged', eta_learn.log_predictions())
            now = datetime.now()
            if now.hour in RETRAIN_HOURS and eta_learn.retrain_due(now):
                eta_learn.learn(now)
        except Exception:
            logging.exception('RP03 ETA cycle failed')
        time.sleep(POLL_SECONDS)


if __name__ == '__main__':
    main()
