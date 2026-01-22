"""Custom exceptions for TIQ Assistant."""


class TIQAssistantError(Exception):
    """Base exception for all TIQ Assistant errors."""
    pass


class ValidationError(TIQAssistantError):
    """Data validation failure."""
    pass


class StorageError(TIQAssistantError):
    """Database/storage operation failure."""
    pass


class ParsingError(TIQAssistantError):
    """File parsing failure."""
    pass


class ExportError(TIQAssistantError):
    """Export operation failure."""
    pass


class ConfigurationError(TIQAssistantError):
    """Configuration/setup issue."""
    pass


class ProjectNotFoundError(TIQAssistantError):
    """Project not found."""
    pass


class TicketNotFoundError(TIQAssistantError):
    """Ticket not found."""
    pass
