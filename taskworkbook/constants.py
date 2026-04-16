FIXED_SHEETS = ("Tasks", "CompletedTasks", "TaskHistory")
SHEET_HEADERS = {
    "Tasks": ("CreatedOn", "Task", "Description", "DayOfWeek", "Worker", "Completion"),
    "CompletedTasks": ("CreatedOn", "Worker", "Task", "Description", "DayOfWeek", "Audited"),
    "TaskHistory": ("CreatedOn", "Worker", "Task", "Description", "DayOfWeek"),
}
SHEET_FILES = {
    "Tasks": "Tasks.xlsx",
    "CompletedTasks": "CompletedTasks.xlsx",
    "TaskHistory": "TaskHistory.xlsx",
}

APP_BG = "#eef3f9"
ACCENT = "#2563eb"
TEXT = "#0f172a"
MUTED = "#64748b"
CHECKBOX_COLUMNS = {
    "Tasks": {"Completion"},
    "CompletedTasks": {"Audited"},
}
FILTERABLE_SHEETS = {"Tasks", "CompletedTasks"}
FILTER_ALL = "All"
