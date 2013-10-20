# coding: utf-8

from datetime import datetime, timedelta
class Style(object):
    ORIGIN = datetime(1899, 12, 30)

    def __init__(self, applyNumFmt, numFmtId):
        self.applyNumFmt = bool(int(applyNumFmt)) if applyNumFmt else False
        self.numFmtId = int(numFmtId) if numFmtId else numFmtId

    @classmethod
    def format(self, value, numFmtId):
        _format = self.get_formatter(numFmtId)
        return _format(value)

    @classmethod
    def get_formatter(self, numFmtId):
        if 12 <= numFmtId <= 17: return self.format_as_date
        if 18 <= numFmtId <= 21: return self.format_as_time
        if numFmtId == 22: return self.format_as_datetime
        return lambda x: x

    @classmethod
    def format_as_date(self, value):
        d = self.ORIGIN + timedelta(int(value))
        return '%04d/%02d/%02d' % (d.year, d.month, d.day)

    @classmethod
    def format_as_time(self, value):
        from math import floor
        totalsec = float(value) * 86400.0
        hour = floor(totalsec / 3600.0)
        rem = totalsec % 3600.0
        minute = floor(rem / 60.0)
        rem = rem % 60.0
        second = floor(rem)
        return '%02d:%02d:%02d' % (hour, minute, second)

    @classmethod
    def format_as_datetime(self, value):
        dt = self.ORIGIN + timedelta(float(value))
        return '%04d/%02d/%02d %02d:%02d:%02d' % (
            dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


