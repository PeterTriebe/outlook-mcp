"""
MCP tools – Tasks.
"""

from __future__ import annotations

import json


def register_tasks_tools(mcp, get_client):
    """Register all task tools."""

    @mcp.tool()
    def list_tasks(include_completed: bool = False, max_results: int = 25) -> str:
        """List Outlook tasks.

        Args:
            include_completed: Include completed tasks (default False).
            max_results: Maximum number of tasks to return (default 25).
        """
        tasks = get_client().list_tasks(
            include_completed=include_completed, max_results=max_results,
        )
        if not tasks:
            return "No tasks found."
        return json.dumps(tasks, indent=2, ensure_ascii=False)

    @mcp.tool()
    def create_task(
        subject: str,
        due_date: str = "",
        body: str = "",
        importance: str = "normal",
        reminder: bool = False,
        categories: str = "",
    ) -> str:
        """Create a new task in Outlook.

        Args:
            subject: Task subject / title.
            due_date: Due date (YYYY-MM-DD). Leave empty for no due date.
            body: Optional description text.
            importance: Priority: low, normal (default), high.
            reminder: Set a reminder (default False).
            categories: Comma-separated category names.
        """
        task = get_client().create_task(
            subject=subject, due_date=due_date, body=body,
            importance=importance, reminder=reminder, categories=categories,
        )
        return f"Task created:\n{json.dumps(task, indent=2, ensure_ascii=False)}"

    @mcp.tool()
    def complete_task(subject_search: str) -> str:
        """Mark a task as complete.

        Args:
            subject_search: Partial subject text to find the task.
        """
        return get_client().complete_task(subject_search)

    @mcp.tool()
    def update_task(
        subject_search: str,
        new_subject: str = "",
        new_due_date: str = "",
        new_body: str = "",
        importance: str = "",
        percent_complete: int | None = None,
        status: str = "",
    ) -> str:
        """Update an existing task found by subject search.

        Args:
            subject_search: Partial subject text to find the task.
            new_subject: New subject (leave empty to keep current).
            new_due_date: New due date YYYY-MM-DD (leave empty to keep current).
            new_body: New description (leave empty to keep current).
            importance: New priority: low, normal, high (leave empty to keep current).
            percent_complete: Completion percentage 0-100 (omit to keep current).
            status: New status: not_started, in_progress, complete, waiting, deferred (leave empty to keep current).
        """
        task = get_client().update_task(
            subject_search=subject_search,
            new_subject=new_subject or None,
            new_due_date=new_due_date or None,
            new_body=new_body or None,
            importance=importance or None,
            percent_complete=percent_complete,
            status=status or None,
        )
        if not task:
            return f"No task found matching '{subject_search}'."
        return f"Task updated:\n{json.dumps(task, indent=2, ensure_ascii=False)}"

    @mcp.tool()
    def delete_task(subject_search: str) -> str:
        """Delete a task by searching its subject.

        Args:
            subject_search: Partial subject text to find the task to delete.
        """
        return get_client().delete_task(subject_search)
