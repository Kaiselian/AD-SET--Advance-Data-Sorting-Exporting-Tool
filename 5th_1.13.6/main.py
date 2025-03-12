import tkinter as tk
import ttkbootstrap as tb
from views.main_view import MainView

def main():
    root = tb.Window(themename="darkly")  # Default theme
    root.title("(╯°□°）╯︵ ┻━┻ Advanced Data Search & Export Tool 1.13.6")
    root.geometry("1920x1080")
    root.state("zoomed")

    # Initialize the main view
    main_view = MainView(root)
    main_view.pack(fill=tk.BOTH, expand=True)

    root.mainloop()

if __name__ == "__main__":
    main()

    # (╯°□°）╯︵ ┻━┻ #
