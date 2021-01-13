# pyinstaller --onedir main.py

__author__ = 'Jade McKenzie', 'Tyler McCamley'

import SetOverdue as Go

DAYS_OVER = 60
MIN_TOTAL = 0.0

Go.syrinx_clear()
overdue = Go.get_overdue(DAYS_OVER, MIN_TOTAL)
Go.set_overdue(overdue[0], DAYS_OVER)
