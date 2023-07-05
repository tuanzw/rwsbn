from datetime import datetime

date_fmts = "%d%m%y,%d%m%Y,%d%b%y,%d%b%Y,%d %b %y,%d %b %Y"


def ymd_date(dt: str) -> str | None:
    for fmt in date_fmts.split(","):
        try:
            return datetime.strftime(datetime.strptime(dt, fmt), "%y%m%d")
        except:
            pass
    return None
