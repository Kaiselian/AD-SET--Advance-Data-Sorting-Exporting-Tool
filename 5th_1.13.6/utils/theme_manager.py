import darkdetect

def get_system_theme():
    """Detects the system theme (light/dark)."""
    return "darkly" if darkdetect.isDark() else "journal"

def change_theme(root, selected_theme):
    """Changes the application theme."""
    root.style.theme_use(selected_theme)