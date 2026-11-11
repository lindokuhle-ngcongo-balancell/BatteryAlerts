"""
Load run-time configuration from environment variables or a .env file.

Priority order (highest to lowest):
  1. Explicit CLI flags (handled in cli.py — they override these values)
  2. Environment variables (set in shell or in a .env file)
  3. Hard-coded defaults below
"""

from __future__ import annotations

import os
from dataclasses import dataclass

try:
    from dotenv import load_dotenv

    load_dotenv(override=False)  # .env values don't override existing env vars
except ImportError:
    pass  # python-dotenv not installed; rely on os.environ only



@dataclass
class SnowflakeConfig:
    """Snowflake connection settings."""

    account: str
    username: str
    warehouse: str
    database: str
    schema: str
    role: str | None
    module_table: str
    cell_table: str
    authenticator: str



def get_snowflake_config() -> SnowflakeConfig:
    """Read Snowflake settings, supporting prefixed and legacy .env key names."""

    def setting(name: str, alias: str, default: str = "") -> str:
        return os.getenv(name) or os.getenv(alias) or default

    return SnowflakeConfig(
        account=setting("SNOWFLAKE_ACCOUNT", "account"),
        username=setting("SNOWFLAKE_USER", "user"),
        warehouse=setting("SNOWFLAKE_WAREHOUSE", "warehouse"),
        database=setting("SNOWFLAKE_DATABASE", "database"),
        schema=setting("SNOWFLAKE_SCHEMA", "schema"),
        role=setting("SNOWFLAKE_ROLE", "role") or None,
        module_table=os.getenv("SNOWFLAKE_MODULE_TABLE", "TELEMETRY_MODULE"),
        cell_table=os.getenv("SNOWFLAKE_CELL_TABLE", "TELEMETRY_CELL"),
        authenticator=setting("SNOWFLAKE_AUTHENTICATOR", "authenticator", "externalbrowser"),
    )
