"""Database layer for the unified DriverAI intake + scoring pipeline."""
from .migrations import get_connection, migrate

__all__ = ["get_connection", "migrate"]
