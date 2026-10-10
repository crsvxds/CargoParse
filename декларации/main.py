import tkinter as tk
try:
    from tkinterdnd2 import TkinterDnD
    root = TkinterDnD.Tk()
except ImportError:
    # Если tkinterdnd2 не установлен, используем обычный Tk
    root = tk.Tk()

from ui import CustomsApp

if __name__ == "__main__":
    app = CustomsApp(root)
    root.mainloop()