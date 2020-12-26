import datetime

from .pdf import update_form_values, get_form_fields


def xldate_to_datetime(xldate):
    temp = datetime.datetime(1899, 12, 30)
    delta = datetime.timedelta(days=xldate)
    return temp + delta
