"""
Python P1 intake engine — replaces Power Automate for email triage.

Also re-exports run_intake / score_existing_excel from the original
hiring_agent/intake.py module so that hiring_agent/__init__.py continues
to work without modification during the migration period.
"""
from .classify import Classifier, ClassifyResult, Classification, BLOCKED, classifier
from .appref import mint_app_ref, extract_quoted_ref, is_valid_app_ref
from .graph_mail import GraphMailClient, Message, Attachment
from .route import Router, Route, RouteResult
from .worker import IntakeWorker, WorkerConfig, WorkerResult

# ── Backward-compatibility shim ───────────────────────────────────────────────
# The original hiring_agent/intake.py exports run_intake and score_existing_excel.
# Those are still used by hiring_agent/__init__.py and bot.py during the migration.
# We import them here via importlib so Python's package resolution finds the .py
# file rather than this directory (which shadows it normally).
import importlib.util as _ilu
import os as _os
from pathlib import Path as _Path

_legacy_path = _Path(__file__).resolve().parent.parent / "intake.py"
if _legacy_path.exists():
    _spec = _ilu.spec_from_file_location("_intake_legacy", _legacy_path)
    _legacy = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(_legacy)
    run_intake = _legacy.run_intake
    score_existing_excel = _legacy.score_existing_excel
else:
    # Legacy file removed — stubs to prevent import errors.
    def run_intake(*a, **kw):
        raise NotImplementedError("Legacy intake.py has been removed; use IntakeWorker.")
    def score_existing_excel(*a, **kw):
        raise NotImplementedError("Legacy intake.py has been removed.")

__all__ = [
    # Classifier
    "Classifier", "ClassifyResult", "Classification", "BLOCKED", "classifier",
    # APP-Ref
    "mint_app_ref", "extract_quoted_ref", "is_valid_app_ref",
    # Graph
    "GraphMailClient", "Message", "Attachment",
    # Routing
    "Router", "Route", "RouteResult",
    # Worker
    "IntakeWorker", "WorkerConfig", "WorkerResult",
    # Legacy compat
    "run_intake", "score_existing_excel",
]

