"""Entry point: python -m echem_gui"""

from .app import EchemGUI

if __name__ == "__main__":
    app = EchemGUI()
    app.mainloop()
