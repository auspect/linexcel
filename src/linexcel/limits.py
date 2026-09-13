"""Validation shared by the public analysis and inspection entry points."""


def validate_limits(**limits: int | None) -> None:
    """None selects a default; zero is a real limit, never 'unlimited'."""
    for name, value in limits.items():
        if value is None:
            continue
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError(f"{name} must be a non-negative integer or None")
        if value < 0:
            raise ValueError(f"{name} must be a non-negative integer or None")


def limit_or_default(name: str, value: int | None, default: int) -> int:
    validate_limits(**{name: value})
    return default if value is None else value
