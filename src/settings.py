"""Application settings loaded from environment variables."""

from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Azure AD / Microsoft Graph configuration."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
    )

    azure_tenant_id: str
    azure_client_id: str


@lru_cache
def get_settings() -> Settings:
    """Get cached settings instance."""
    return Settings()
