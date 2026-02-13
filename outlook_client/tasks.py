"""
Outlook COM client – task operations.
"""

from __future__ import annotations

import datetime as _dt

from outlook_client._helpers import (
    IMPORTANCE_REVERSE,
    TASK_STATUS_REVERSE,
    task_to_dict,
)


class TasksMixin:
    """Tasks methods – mixed into OutlookClient."""

    def _tasks_folder(self):
        """Return the default Tasks folder (olFolderTasks = 13)."""
        return self._ns.GetDefaultFolder(13)

    # ── Queries ──────────────────────────────────────────────────────

    def list_tasks(
        self, include_completed: bool = False, max_results: int = 25,
    ) -> list[dict]:
        """List tasks, optionally including completed ones."""
        tasks = self._tasks_folder()
        items = tasks.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results: list[dict] = []
        for i, task in enumerate(items):
            if i >= max_results:
                break
            results.append(task_to_dict(task))
        return results

    # ── Mutations ────────────────────────────────────────────────────

    def create_task(
        self,
        subject: str,
        due_date: str = "",
        body: str = "",
        importance: str = "normal",
        reminder: bool = False,
        categories: str = "",
    ) -> dict:
        """Create a new Outlook task."""
        task = self._app.CreateItem(3)  # olTaskItem
        task.Subject = subject
        if due_date:
            task.DueDate = _dt.datetime.fromisoformat(due_date).strftime("%m/%d/%Y")
        if body:
            task.Body = body
        task.Importance = IMPORTANCE_REVERSE.get(importance.lower(), 1)
        task.ReminderSet = reminder
        if categories:
            task.Categories = categories
        task.Save()
        return task_to_dict(task)

    def complete_task(self, subject_search: str) -> str:
        """Mark a task as complete."""
        tasks = self._tasks_folder()
        items = tasks.Items
        needle = subject_search.lower()
        for task in items:
            if needle in str(getattr(task, "Subject", "")).lower():
                task.Complete = True
                task.Status = 2  # olTaskComplete
                task.PercentComplete = 100
                task.Save()
                return f"Task completed: '{task.Subject}'"
        return f"No task found matching '{subject_search}'."

    def update_task(
        self,
        subject_search: str,
        new_subject: str | None = None,
        new_due_date: str | None = None,
        new_body: str | None = None,
        importance: str | None = None,
        percent_complete: int | None = None,
        status: str | None = None,
    ) -> dict | None:
        """Update a task found by subject search."""
        tasks = self._tasks_folder()
        items = tasks.Items
        needle = subject_search.lower()
        for task in items:
            if needle in str(getattr(task, "Subject", "")).lower():
                if new_subject is not None:
                    task.Subject = new_subject
                if new_due_date is not None:
                    task.DueDate = _dt.datetime.fromisoformat(new_due_date).strftime("%m/%d/%Y")
                if new_body is not None:
                    task.Body = new_body
                if importance is not None:
                    task.Importance = IMPORTANCE_REVERSE.get(importance.lower(), 1)
                if percent_complete is not None:
                    task.PercentComplete = percent_complete
                if status is not None:
                    task.Status = TASK_STATUS_REVERSE.get(status.lower(), 0)
                task.Save()
                return task_to_dict(task)
        return None

    def delete_task(self, subject_search: str) -> str:
        """Delete a task by subject search."""
        tasks = self._tasks_folder()
        items = tasks.Items
        needle = subject_search.lower()
        for task in items:
            if needle in str(getattr(task, "Subject", "")).lower():
                subj = task.Subject
                task.Delete()
                return f"Deleted task: '{subj}'"
        return f"No task found matching '{subject_search}'."
