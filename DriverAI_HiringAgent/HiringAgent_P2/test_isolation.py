"""Send every path a test can write to into a throwaway directory.

Import this BEFORE `hiring_agent` (and before anything that imports it). config.py resolves
LOGS_DIR / OUTPUT_DIR / INPUT_DIR at import time and calls mkdir on them, so setting these
afterwards is too late - the live folders are created and written to regardless.

Without it the suites appended to the production P2_Logs/<year>/<month>/P2.log and created
real Candidate_List_Results_<timestamp>.xlsx files in P2_Final_Results/, where a file left
by a test run is indistinguishable from a client-facing export.

setdefault, not assignment: a caller who has deliberately pointed these somewhere (CI, a
scratch run) keeps their choice.
"""
import atexit
import os
import shutil
import tempfile
from pathlib import Path

ROOT = Path(tempfile.mkdtemp(prefix="p2_test_"))
atexit.register(shutil.rmtree, ROOT, True)

for _var, _sub in (
    ("HIRING_LOGS_DIR", "logs"),
    ("HIRING_PUBLISH_DIR", "publish"),
    ("HIRING_CLIENT_EXPORT_DIR", "client_export"),
    ("HIRING_OUTPUT_DIR", "output"),
    ("HIRING_INPUT_DIR", "input"),
):
    os.environ.setdefault(_var, str(ROOT / _sub))
