import json
import sys
from datetime import datetime, timezone
from typing import Any


def log(level: str, event: str, **kwargs: Any):
    log_entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "level": level,
        "event": event,
        **kwargs
    }
    print(json.dumps(log_entry), file=sys.stderr, flush=True)


def info(event: str, **kwargs: Any):
    log("info", event, **kwargs)


def warn(event: str, **kwargs: Any):
    log("warn", event, **kwargs)


def error(event: str, **kwargs: Any):
    log("error", event, **kwargs)
