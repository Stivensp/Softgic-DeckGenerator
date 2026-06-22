"""
Softgic Deck Generator — Menu interactivo
Ejecutar: menu.bat  (doble clic)  o  python menu.py
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import traceback
from pathlib import Path

ROOT    = Path(__file__).resolve().parent
PYTHON  = sys.executable
CLEAR   = "cls" if os.name == "nt" else "clear"
VERSION = "1.0.0"
BOX_W   = 56   # ancho visible del header

# ── Terminal: UTF-8 + ANSI ──────────────────────────────────────────────────
if os.name == "nt":
    os.system("chcp 65001 > nul 2>&1")
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleMode(
            ctypes.windll.kernel32.GetStdHandle(-11), 7
        )
        _ANSI_OK = True
    except Exception:
        _ANSI_OK = False
else:
    _ANSI_OK = True

def _c(code: str) -> str:
    return code if _ANSI_OK else ""

RST   = _c("\033[0m")
BOLD  = _c("\033[1m")
BLUE  = _c("\033[38;2;0;174;239m")   # #00AEEF — Softgic accent
WHITE = _c("\033[97m")
GREY  = _c("\033[90m")
CYAN  = _c("\033[96m")
GREEN = _c("\033[92m")
RED   = _c("\033[91m")
YELL  = _c("\033[93m")

_ANSI_PAT = re.compile(r"\033\[[0-9;]*m")

def _vlen(s: str) -> int:
    """Longitud visible de un string (descarta codigos ANSI)."""
    return len(_ANSI_PAT.sub("", s))

# ── Box drawing ───────────────────────────────────────────────────────────────
def _box_top() -> None:
    print(f"  {BLUE}┌{'─' * (BOX_W - 2)}┐{RST}")

def _box_row(text: str = "") -> None:
    pad = BOX_W - 4 - _vlen(text)
    print(f"  {BLUE}│{RST}  {text}{' ' * max(0, pad)}  {BLUE}│{RST}")

def _box_bot() -> None:
    print(f"  {BLUE}└{'─' * (BOX_W - 2)}┘{RST}")

def _rule() -> None:
    print(f"  {GREY}{'─' * (BOX_W + 2)}{RST}")

# ── Primitivos de UI ──────────────────────────────────────────────────────────
def clear() -> None:
    os.system(CLEAR)

def blank() -> None:
    print()

def header(subtitle: str = "") -> None:
    sub = subtitle or "Genera presentaciones .pptx desde YAML / JSON"
    _box_top()
    _box_row(f"{BOLD}{WHITE}✦  SOFTGIC DECK GENERATOR{RST}  {GREY}v{VERSION}{RST}")
    _box_row(f"{GREY}{sub}{RST}")
    _box_bot()
    blank()

def item(key: str, label: str, note: str = "") -> None:
    n = f"  {GREY}({note}){RST}" if note else ""
    print(f"  {CYAN}{BOLD}[{key}]{RST}  {WHITE}{label}{RST}{n}")

def subdesc(text: str) -> None:
    print(f"         {GREY}{text}{RST}")

def pause(msg: str = "Presiona Enter para continuar...") -> None:
    input(f"\n  {GREY}{msg}{RST}")

def ask(prompt: str) -> str:
    return input(f"\n  {CYAN}›{RST} {prompt} ").strip()

def success(msg: str) -> None:
    print(f"\n  {GREEN}✓{RST}  {WHITE}{msg}{RST}")

def error(msg: str) -> None:
    print(f"\n  {RED}✗{RST}  {RED}{msg}{RST}")

def warn(msg: str) -> None:
    print(f"\n  {YELL}!{RST}  {YELL}{msg}{RST}")

# ── Utilidades ─────────────────────────────────────────────────────────────────
def find_yaml_files() -> list[Path]:
    files: list[Path] = []
    for folder in (ROOT / "examples", ROOT):
        if folder.exists():
            files += sorted(folder.glob("*.yaml")) + sorted(folder.glob("*.yml"))
    seen: set[Path] = set()
    result: list[Path] = []
    for f in files:
        r = f.resolve()
        if r not in seen:
            seen.add(r)
            result.append(f)
    return result

def count_outputs() -> int:
    d = ROOT / "output"
    return len(list(d.glob("*.pptx"))) if d.exists() else 0

def open_file(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]

def run(cmd: list[str]) -> int:
    return subprocess.run(cmd, cwd=str(ROOT)).returncode

# ── Flujos de seleccion ────────────────────────────────────────────────────────
def pick_input_file() -> Path | None:
    yaml_files = find_yaml_files()
    clear()
    header("Seleccionar archivo de entrada")

    if yaml_files:
        for i, f in enumerate(yaml_files, 1):
            item(str(i), str(f.relative_to(ROOT)))
        blank()
        item(str(len(yaml_files) + 1), "Escribir ruta manualmente")
        item("0", "Cancelar")

        choice = ask("Archivo:")
        if choice == "0":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(yaml_files):
            return yaml_files[int(choice) - 1].resolve()
        if choice != str(len(yaml_files) + 1):
            error("Opción no válida.")
            pause()
            return None

    ruta = ask("Ruta del archivo YAML/JSON:").strip('"')
    if not ruta:
        return None
    p = Path(ruta) if Path(ruta).is_absolute() else ROOT / ruta
    if not p.exists():
        error(f"No existe: '{p}'")
        pause()
        return None
    return p.resolve()


def pick_output_file(input_file: Path) -> Path | None:
    default = ROOT / "output" / (input_file.stem + ".pptx")
    clear()
    header("Archivo de salida")
    print(f"  {GREY}Por defecto:{RST}  {WHITE}output\\{default.name}{RST}")
    print(f"  {GREY}(deja vacío para usar ese nombre){RST}")

    ruta = ask("Nombre del .pptx:").strip('"')
    if not ruta:
        return default
    p = Path(ruta)
    if not p.suffix:
        p = p.with_suffix(".pptx")
    if not p.is_absolute():
        p = ROOT / "output" / p
    return p.resolve()


# ── Secciones del menu ─────────────────────────────────────────────────────────
def menu_generate() -> None:
    input_file = pick_input_file()
    if not input_file:
        return
    output_file = pick_output_file(input_file)
    if not output_file:
        return

    clear()
    header(f"Generando  ›  {input_file.name}")
    _rule()
    ok_flag = run([
        PYTHON, str(ROOT / "generate.py"),
        "--input",  str(input_file),
        "--output", str(output_file),
    ]) == 0
    _rule()

    if ok_flag and output_file.exists():
        size_kb = output_file.stat().st_size / 1024
        success(f"{output_file.name}  —  {size_kb:.1f} KB")
        resp = ask("Abrir ahora? [S/n]:").lower()
        if resp in ("", "s", "si", "y", "yes"):
            open_file(output_file)
    else:
        error("No se pudo generar el deck. Revisa los mensajes de arriba.")
    pause()


def menu_list_outputs() -> None:
    clear()
    header("Decks generados")
    output_dir = ROOT / "output"
    files = sorted(output_dir.glob("*.pptx")) if output_dir.exists() else []

    if not files:
        warn("No hay archivos en output/ todavía.")
        pause()
        return

    for i, f in enumerate(files, 1):
        item(str(i), f.name, f"{f.stat().st_size / 1024:.1f} KB")
    blank()
    item("0", "Volver")

    choice = ask("Abrir archivo:")
    if choice.isdigit() and 1 <= int(choice) <= len(files):
        open_file(files[int(choice) - 1])
    pause()


def menu_tests() -> None:
    clear()
    header("Ejecutar tests")
    item("1", "Todos los tests")
    item("2", "Con reporte de cobertura")
    item("3", "Solo tests unitarios")
    item("4", "Solo tests de integración")
    blank()
    item("0", "Volver")

    choice = ask("Opción:")
    cmds: dict[str, list[str]] = {
        "1": [PYTHON, "-m", "pytest", "tests/", "-v"],
        "2": [PYTHON, "-m", "pytest", "tests/", "--cov=src", "--cov-report=term-missing"],
        "3": [PYTHON, "-m", "pytest", "tests/unit/", "-v"],
        "4": [PYTHON, "-m", "pytest", "tests/integration/", "-v"],
    }
    if choice not in cmds:
        return
    clear()
    header("Tests")
    _rule()
    run(cmds[choice])
    _rule()
    pause()


def menu_diseno() -> None:
    while True:
        clear()
        header("Herramientas de diseño")

        item("1", "Abrir editor visual")
        subdesc("Genera config_visual.pptx y lo abre en PowerPoint")
        blank()
        item("2", "Aplicar cambios del editor")
        subdesc("Lee las posiciones editadas → actualiza config/generated/layout.yaml y config/generated/input_template.yaml")
        blank()
        item("3", "Preview de layouts")
        subdesc("Genera PREVIEW_layouts.pptx con todas las cajas etiquetadas")
        blank()
        item("4", "Inspeccionar template maestro")
        subdesc("Lista los layouts disponibles en assets/template.pptx")
        blank()
        item("5", "Regenerar assets")
        subdesc("Recrea template.pptx, logos y avatar en images/")
        blank()
        item("6", "Abrir editor de contenido web")
        subdesc("Formulario en el navegador para llenar titulos/bullets/precios y generar el deck (frontend/)")
        blank()
        item("0", "Volver")

        choice = ask("Opción:")

        if choice == "0":
            break

        elif choice == "1":
            clear()
            header("Editor visual")
            _rule()
            run([PYTHON, str(ROOT / "tools" / "open_visual_editor.py")])
            _rule()
            blank()
            print(f"  {WHITE}Cuando termines en PowerPoint:{RST}")
            print(f"  {GREY}  1. Guarda el archivo  (Ctrl+S){RST}")
            print(f"  {GREY}  2. Cierra PowerPoint{RST}")
            print(f"  {GREY}  3. Elige [2] para aplicar los cambios{RST}")
            pause()

        elif choice == "2":
            clear()
            header("Aplicar cambios del editor")
            _rule()
            run([PYTHON, str(ROOT / "tools" / "apply_visual_config.py")])
            _rule()
            resp = ask("Generar deck de prueba para ver los cambios? [S/n]:").lower()
            if resp in ("", "s", "si", "y", "yes"):
                inp = ROOT / "examples" / "propuesta_comercial.yaml"
                out = ROOT / "output" / "test_layout.pptx"
                _rule()
                run([PYTHON, str(ROOT / "generate.py"),
                     "--input", str(inp), "--output", str(out)])
                _rule()
                if out.exists():
                    open_file(out)
            pause()

        elif choice == "3":
            clear()
            header("Preview de layouts")
            _rule()
            run([PYTHON, str(ROOT / "tools" / "preview_layout.py")])
            _rule()
            out = ROOT / "output" / "PREVIEW_layouts.pptx"
            if out.exists():
                resp = ask("Abrir ahora? [S/n]:").lower()
                if resp in ("", "s", "si", "y", "yes"):
                    open_file(out)
            pause()

        elif choice == "4":
            clear()
            header("Inspeccionar template")
            _rule()
            run([PYTHON, str(ROOT / "tools" / "inspect_template.py")])
            _rule()
            pause()

        elif choice == "5":
            clear()
            header("Regenerar assets")
            _rule()
            run([PYTHON, str(ROOT / "tools" / "create_test_assets.py")])
            _rule()
            success("Assets regenerados.")
            pause()

        elif choice == "6":
            clear()
            header("Editor de contenido web")
            _rule()
            print(f"  {WHITE}Abriendo en http://127.0.0.1:5000{RST}")
            print(f"  {GREY}Formulario para llenar el contenido del deck y generarlo, sin tocar YAML a mano.{RST}")
            print(f"  {GREY}Presiona Ctrl+C para detener el servidor y volver al menu.{RST}")
            blank()
            import threading
            import webbrowser
            threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
            try:
                run([PYTHON, str(ROOT / "frontend" / "server.py")])
            except KeyboardInterrupt:
                pass
            _rule()
            pause()


def main() -> None:
    while True:
        clear()
        header()

        n = count_outputs()
        note = f"{n} generado{'s' if n != 1 else ''}" if n else ""

        item("1", "Generar un deck")
        item("2", "Ver decks generados", note)
        item("3", "Herramientas de diseño")
        item("4", "Ejecutar tests")
        blank()
        item("0", "Salir")

        choice = ask("Opción:")

        if choice == "1":
            menu_generate()
        elif choice == "2":
            menu_list_outputs()
        elif choice == "3":
            menu_diseno()
        elif choice == "4":
            menu_tests()
        elif choice == "0":
            clear()
            print(f"\n  {GREY}Hasta luego.{RST}\n")
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n  {GREY}Cancelado.{RST}\n")
    except Exception:
        print(f"\n  {RED}{'═' * 52}{RST}")
        print(f"  {RED}ERROR INESPERADO:{RST}")
        print(f"  {RED}{'═' * 52}{RST}")
        traceback.print_exc()
        input("\nPresiona Enter para cerrar...")
