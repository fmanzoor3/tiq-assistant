"""Configuration management for TIQ Assistant."""

from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict

from tiq_assistant.core.models import ActivityCode


class Settings(BaseSettings):
    """Application settings loaded from environment or .env file."""

    model_config = SettingsConfigDict(
        env_prefix="TIQ_",
        env_file=".env",
        env_file_encoding="utf-8",
    )

    # Paths
    data_dir: Path = Path.home() / ".tiq-assistant"
    database_name: str = "tiq.db"

    # User defaults
    consultant_id: str = "FMANZOOR"
    default_location: str = "ANKARA"
    default_activity_code: str = "GLST"
    meeting_activity_code: str = "TPLNT"

    # Matching
    min_match_confidence: float = 0.5
    skip_canceled_meetings: bool = True
    min_meeting_duration_minutes: int = 15

    # Export
    date_format: str = "%d.%m.%Y"

    @property
    def database_path(self) -> Path:
        return self.data_dir / self.database_name

    def ensure_data_dir(self) -> None:
        """Create data directory if it doesn't exist."""
        self.data_dir.mkdir(parents=True, exist_ok=True)


# Global settings instance
settings = Settings()
