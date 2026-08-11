class ServiceError(Exception):
    """Base class for errors that are safe to report to API clients."""


class ValidationError(ServiceError):
    """The client supplied an invalid payload."""


class AuthenticationError(ServiceError):
    """The local service token is missing or invalid."""

