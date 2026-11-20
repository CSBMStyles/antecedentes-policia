"""Alias de la CLI. Se conserva el nombre por compatibilidad.

Uso:
    python main_sin_ui.py --dry-run --si
    python main_sin_ui.py --perfil AUXILIAR --tanda 5 --si
"""

from main import main

if __name__ == "__main__":
    raise SystemExit(main())
