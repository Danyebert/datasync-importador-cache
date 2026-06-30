from __future__ import annotations

import logging
from pathlib import Path
from datetime import datetime


def configurar_logger(base_dir: Path) -> logging.Logger:
    log_dir = base_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"import_cache_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger("datasync_import_cache")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.info("Log iniciado em %s", log_file)
    return logger
