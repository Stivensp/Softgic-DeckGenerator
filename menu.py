"""
Softgic Deck Generator — Menu interactivo
Ejecutar: menu.bat  (doble clic)  o  python menu.py
"""
from __future__ import annotations

import os
import subprocess
import sys
import traceback
from pathlib import Path

# Directorio raiz del proyecto (donde vive menu.py)
ROOT   = Path(__file__).resolve().parent
PYTHON = sys.executable
CLEAR  = "cls" if os.name == "nt" else "clear"


def clear() -> None:
    os.system(CLEAR)


def header() -> None:
    print("=" * 55)
    print("   SOFTGIC DECK GENERATOR")
    print("=" * 55)
    print()


def pause(msg: str = "Presiona Enter para continuar...") -> None:
    input(f"\n{msg}")


def find_yaml_files() -> list[Path]:
    files: list[Path] = []
    for folder in (ROOT / "examples", ROOT):
        if folder.exists():
            files += sorted(folder.glob("*.yaml")) + sorted(folder.glob("*.yml"))
    seen: set[Path] = set()
    unique: list[Path] = []
    for f in files:
        resolved = f.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(f)
    return unique


def pick_input_file() -> Path | None:
    yaml_files = find_yaml_files()

    print("  ARCHIVO DE ENTRADA\n")
    if yaml_files:
        print("  Archivos encontrados:")
        for i, f in enumerate(yaml_files, 1):
            print(f"    [{i}] {f.relative_to(ROOT)}")
        print(f"    [{len(yaml_files)+1}] Escribir ruta manualmente")
        print("    [0] Cancelar")
        print()
        choice = input("  Selecciona una opcion: ").strip()

        if choice == "0":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(yaml_files):
            return yaml_files[int(choice) - 1].resolve()
        if choice == str(len(yaml_files) + 1):
            pass
        else:
            print("\n  Opcion invalida.")
            pause()
            return None

    ruta = input("  Ruta del archivo YAML/JSON: ").strip().strip('"')
    if not ruta:
        return None
    p = Path(ruta)
    if not p.is_absolute():
        p = ROOT / p
    if not p.exists():
        print(f"\n  ERROR: No existe el archivo '{p}'")
        pause()
        return None
    return p.resolve()


def pick_output_file(input_file: Path) -> Path | None:
    default = ROOT / "output" / (input_file.stem + ".pptx")
    print("  ARCHIVO DE SALIDA\n")
    print(f"  Nombre por defecto: output\\{default.name}")
    print("  (deja en blanco para usar ese nombre)\n")
    ruta = input("  Nombre del .pptx: ").strip().strip('"')
    if not ruta:
        return default
    p = Path(ruta)
    if not p.suffix:
        p = p.with_suffix(".pptx")
    if not p.is_absolute():
        p = ROOT / "output" / p
    return p.resolve()


def run_generator(input_file: Path, output_file: Path) -> bool:
    print()
    print("-" * 55)
    cmd = [
        PYTHON,
        str(ROOT / "generate.py"),   # ruta absoluta siempre
        "--input",  str(input_file),
        "--output", str(output_file),
    ]
    result = subprocess.run(cmd, cwd=str(ROOT))
    print("-" * 55)
    return result.returncode == 0


def open_file(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]


def menu_generate() -> None:
    clear()
    header()
    input_file = pick_input_file()
    if not input_file:
        return

    clear()
    header()
    output_file = pick_output_file(input_file)
    if not output_file:
        return

    clear()
    header()
    print(f"  Generando deck...\n")
    ok = run_generator(input_file, output_file)

    if ok and output_file.exists():
        print(f"\n  Deck generado: {output_file}")
        resp = input("\n  Abrir el archivo ahora? [S/n]: ").strip().lower()
        if resp in ("", "s", "si", "y", "yes"):
            open_file(output_file)
    else:
        print("\n  El deck no se pudo generar. Revisa los mensajes de arriba.")

    pause()


def menu_list_outputs() -> None:
    clear()
    header()
    print("  DECKS GENERADOS\n")
    output_dir = ROOT / "output"
    files = sorted(output_dir.glob("*.pptx")) if output_dir.exists() else []

    if not files:
        print("  No hay archivos en output/ todavia.")
        pause()
        return

    for i, f in enumerate(files, 1):
        size_kb = f.stat().st_size / 1024
        print(f"  [{i}] {f.name:<40} {size_kb:>6.1f} KB")

    print("\n  [0] Volver\n")
    choice = input("  Selecciona un archivo para abrirlo: ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(files):
        open_file(files[int(choice) - 1])
    pause()


def menu_run_tests() -> None:
    clear()
    header()
    print("  EJECUTAR TESTS\n")
    print("  [1] Todos los tests")
    print("  [2] Tests con cobertura")
    print("  [3] Solo tests unitarios")
    print("  [4] Solo tests de integracion")
    print("  [0] Volver\n")
    choice = input("  Selecciona: ").strip()

    cmds: dict[str, list[str]] = {
        "1": [PYTHON, "-m", "pytest", "tests/", "-v"],
        "2": [PYTHON, "-m", "pytest", "tests/", "--cov=src", "--cov-report=term-missing"],
        "3": [PYTHON, "-m", "pytest", "tests/unit/", "-v"],
        "4": [PYTHON, "-m", "pytest", "tests/integration/", "-v"],
    }
    if choice not in cmds:
        return

    clear()
    header()
    subprocess.run(cmds[choice], cwd=str(ROOT))
    pause()


def menu_diseno() -> None:
    clear()
    header()
    print("  HERRAMIENTAS DE DISENO\n")
    print("  [1] Abrir editor visual de layout")
    print("      Abre PowerPoint con cajas movibles — arrastralas donde quieras")
    print()
    print("  [2] Aplicar cambios del editor visual")
    print("      Lee las posiciones que moviste y actualiza config/layout.yaml")
    print()
    print("  [3] Ver preview de layouts (referencia)")
    print("      Genera output/PREVIEW_layouts.pptx con cajas etiquetadas")
    print()
    print("  [4] Inspeccionar template maestro")
    print("      Muestra posiciones de shapes en assets/template.pptx")
    print()
    print("  [5] Regenerar assets (logo + avatar)")
    print()
    print("  [0] Volver\n")
    choice = input("  Selecciona: ").strip()

    if choice == "1":
        clear()
        header()
        print("  Generando editor visual...\n")
        subprocess.run([PYTHON, str(ROOT / "tools" / "open_visual_editor.py")], cwd=str(ROOT))
        print()
        print("  Cuando termines de mover las cajas en PowerPoint:")
        print("  1. Guarda el archivo (Ctrl+S)")
        print("  2. Cierra PowerPoint")
        print("  3. Vuelve aqui y elige la opcion [2] para aplicar los cambios")
        pause()

    elif choice == "2":
        clear()
        header()
        print("  Aplicando cambios del editor visual...\n")
        subprocess.run([PYTHON, str(ROOT / "tools" / "apply_visual_config.py")], cwd=str(ROOT))
        print()
        resp = input("  Quieres generar un deck de prueba ahora? [S/n]: ").strip().lower()
        if resp in ("", "s", "si", "y", "yes"):
            inp = ROOT / "examples" / "propuesta_comercial.yaml"
            out = ROOT / "output" / "test_layout.pptx"
            subprocess.run([PYTHON, str(ROOT / "generate.py"),
                            "--input", str(inp), "--output", str(out)], cwd=str(ROOT))
            if out.exists():
                open_file(out)
        pause()

    elif choice == "3":
        clear()
        header()
        print("  Generando preview...\n")
        subprocess.run([PYTHON, str(ROOT / "tools" / "preview_layout.py")], cwd=str(ROOT))
        out = ROOT / "output" / "PREVIEW_layouts.pptx"
        if out.exists():
            resp = input("\n  Abrir ahora? [S/n]: ").strip().lower()
            if resp in ("", "s", "si", "y", "yes"):
                open_file(out)
        pause()

    elif choice == "4":
        clear()
        header()
        subprocess.run([PYTHON, str(ROOT / "tools" / "inspect_template.py")], cwd=str(ROOT))
        pause()

    elif choice == "5":
        clear()
        header()
        print("  Regenerando assets...\n")
        subprocess.run([PYTHON, str(ROOT / "tools" / "create_test_assets.py")], cwd=str(ROOT))
        pause()


def main() -> None:
    while True:
        clear()
        header()
        print("  [1] Generar un deck")
        print("  [2] Ver decks generados")
        print("  [3] Herramientas de diseno")
        print("  [4] Ejecutar tests")
        print("  [0] Salir")
        print()
        choice = input("  Selecciona una opcion: ").strip()

        if choice == "1":
            menu_generate()
        elif choice == "2":
            menu_list_outputs()
        elif choice == "3":
            menu_diseno()
        elif choice == "4":
            menu_run_tests()
        elif choice == "0":
            clear()
            print("\n  Hasta luego.\n")
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  Cancelado.")
    except Exception:
        print("\n" + "=" * 55)
        print("  ERROR INESPERADO:")
        print("=" * 55)
        traceback.print_exc()
        input("\nPresiona Enter para cerrar...")
