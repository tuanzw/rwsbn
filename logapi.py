from logging import handlers, Formatter, getLogger, DEBUG
from datetime import datetime


filename = "./app.log"


def filer(self):
    subfix = datetime.now().strftime("%Y-%m-%d")
    return f"{filename}_{subfix}"


logger = getLogger("applog")

formatter = Formatter("[%(asctime)s] %(levelname)-8s| %(module)-10s| %(message)s")
handler = handlers.TimedRotatingFileHandler(
    filename=filename, when="w5", interval=1, backupCount=7, encoding="utf-8"
)
handler.rotation_filename = filer
handler.setFormatter(fmt=formatter)
logger.addHandler(handler)
logger.setLevel(DEBUG)
