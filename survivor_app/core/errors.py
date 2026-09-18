class FileWriteError(Exception):
    """Raised when an output file cannot be written (e.g. it is open in Excel)."""

    def __init__(self, filename: str):
        super().__init__(f"[{filename}] cannot be written. It is probably open in another program.")
        self.filename = filename


class ConfigValidationError(Exception):
    """Raised when a Config fails validation and an operation refuses to proceed."""

    def __init__(self, errors: list[str]):
        super().__init__("; ".join(errors))
        self.errors = errors
