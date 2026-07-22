import tkinter as tk
import unittest
from tkinter import ttk

from todo_app import ScheduleDialog, ScheduleEvent, Task, TaskDialog


class NoteReturnKeyTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.geometry("1x1+10000+10000")
        self.root.update()

    def tearDown(self):
        self.root.destroy()

    def assert_enter_inserts_newline(self, dialog):
        dialog.update()
        dialog.txt_note.delete("1.0", "end")
        dialog.txt_note.insert("1.0", "첫째 줄")
        dialog.txt_note.mark_set("insert", "end-1c")
        dialog.txt_note.focus_force()
        dialog.update()

        dialog.txt_note.event_generate("<Return>")
        self.root.update()

        self.assertEqual(dialog.winfo_exists(), 1, "메모에서 Enter를 눌러도 저장되면 안 됩니다")
        self.assertEqual(dialog.txt_note.get("1.0", "end-1c"), "첫째 줄\n")
        dialog.destroy()

    def test_task_note_enter_inserts_newline_without_saving(self):
        dialog = TaskDialog(self.root, task=Task(title="업무"))
        self.assert_enter_inserts_newline(dialog)

    def test_schedule_note_enter_inserts_newline_without_saving(self):
        event = ScheduleEvent(title="일정", start="2026-07-22")
        dialog = ScheduleDialog(self.root, event=event)
        self.assert_enter_inserts_newline(dialog)

    def test_task_title_enter_still_saves(self):
        task = Task(title="업무")
        dialog = TaskDialog(self.root, task=task)
        dialog.update()
        title_entry = next(
            widget for widget in dialog.winfo_children()
            if isinstance(widget, ttk.Entry)
        )
        title_entry.focus_force()
        dialog.update()

        title_entry.event_generate("<Return>")
        self.root.update()

        self.assertEqual(dialog.winfo_exists(), 0)
        self.assertIs(dialog.result, task)


if __name__ == "__main__":
    unittest.main()
