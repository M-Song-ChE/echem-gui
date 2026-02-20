"""Entry point: python -m echem_core"""

from .app import EchemGUI

if __name__ == "__main__":
    app = EchemGUI()
    app.mainloop()
