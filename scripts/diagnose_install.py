"""One-off diagnostic: why `pip` may not be recognized. Writes NDJSON to .cursor/debug.log."""
import json
import os
import subprocess
import sys
from pathlib import Path

LOG_PATH = Path(__file__).resolve().parents[1] / ".cursor" / "debug.log"
SESSION = "debug-session"
RUN_ID = "diagnose-install"

# #region agent log
payload = {
    "sessionId": SESSION,
    "runId": RUN_ID,
    "hypothesisId": "H1",
    "location": "scripts/diagnose_install.py",
    "message": "install env check",
    "data": {
        "sys_executable": sys.executable,
        "path_contains_scripts": "Scripts" in os.environ.get("Path", ""),
        "path_contains_pip": "pip" in os.environ.get("Path", "").lower(),
        "python_m_pip_ok": subprocess.run(
            [sys.executable, "-m", "pip", "--version"], capture_output=True
        ).returncode == 0,
    },
    "timestamp": __import__("time").time() * 1000,
}
LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
with open(LOG_PATH, "a", encoding="utf-8") as f:
    f.write(json.dumps(payload) + "\n")
# #endregion

print("sys.executable:", sys.executable)
print("python -m pip ok:", payload["data"]["python_m_pip_ok"])
print("log written to:", LOG_PATH)
