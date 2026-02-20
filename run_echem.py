"""Launch the Electrochemistry Analysis GUI."""

from echem_core import EchemGUI

if __name__ == "__main__":
    app = EchemGUI()
    app.mainloop()