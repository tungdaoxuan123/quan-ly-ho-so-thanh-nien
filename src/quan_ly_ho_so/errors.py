"""Exceptions shared across the workbook-backed app modules."""


class WorkbookError(Exception):
    """Expected workbook operation failure shown to the user."""


class WorkbookBusyError(WorkbookError):
    pass


class StaleWorkbookError(WorkbookError):
    pass
