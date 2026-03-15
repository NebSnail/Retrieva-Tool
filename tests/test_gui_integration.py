import unittest
from unittest import mock

import tkinter as tk

from query_tool import launch_gui


def _can_start_tk() -> bool:
    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        return True
    except tk.TclError:
        return False


@unittest.skipUnless(_can_start_tk(), "Tk environment is not available")
class TestGuiIntegration(unittest.TestCase):
    def test_launch_gui_smoke(self):
        # Smoke test: GUI should initialize and exit without raising.
        with mock.patch.object(tk.Tk, "mainloop", lambda self: self.destroy()):
            launch_gui(default_excel=__import__("pathlib").Path("not_exists.xlsx"))


if __name__ == "__main__":
    unittest.main()
