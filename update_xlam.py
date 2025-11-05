import argparse
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Optional, Sequence

XLAM_FILE_FORMAT = 55  # Excel constant xlOpenXMLAddIn
DEFAULT_BUILD_DIRNAME = "build"


def get_default_addins_folder() -> Optional[str]:
    """Return the default Microsoft Excel Add-ins folder for the current user."""
    user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    candidate = Path(user_profile) / "AppData" / "Roaming" / "Microsoft" / "AddIns"
    return str(candidate) if candidate.exists() else None


def _read_text_with_fallback(path: Path) -> str:
    """Read text using common VBA encodings with graceful fallback."""
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    # As a last resort, replace undecodable bytes to avoid crashing the build.
    return data.decode("latin-1", errors="replace")


def _read_vb_name(source_path: Path) -> Optional[str]:
    """Extract the VB_Name attribute from a VBA source file."""
    try:
        for line in _read_text_with_fallback(source_path).splitlines():
            stripped = line.strip()
            if stripped.startswith("Attribute VB_Name"):
                _, value = stripped.split("=", 1)
                return value.strip().strip('"')
    except FileNotFoundError:
        return None
    return None


def collect_vba_sources(repo_dir: Path) -> Dict[str, Sequence[Path]]:
    """Collect VBA source files grouped by component type."""
    sources: Dict[str, Sequence[Path]] = {
        "standard": [],
        "classes": [],
        "forms": [],
        "sheets": [],
    }

    modules_dir = repo_dir / "Modules"
    if modules_dir.is_dir():
        sources["standard"] = sorted(modules_dir.glob("*.bas"), key=lambda p: p.name)

    class_dir = repo_dir / "Class Modules"
    if class_dir.is_dir():
        sources["classes"] = sorted(class_dir.glob("*.cls"), key=lambda p: p.name)

    forms_dir = repo_dir / "Forms"
    if forms_dir.is_dir():
        sources["forms"] = sorted(forms_dir.glob("*.frm"), key=lambda p: p.name)

    sheet_dir = repo_dir / "Microsoft Excel Objects"
    sheet_modules = []
    if sheet_dir.is_dir():
        for path in sorted(sheet_dir.glob("*.cls"), key=lambda p: p.name):
            if path.name.lower() == "thisworkbook.cls":
                continue
            sheet_modules.append(path)
    sources["sheets"] = sheet_modules

    return sources


def ensure_pywin32():
    """Import pywin32 components, raising a helpful error if missing."""
    try:
        import pythoncom  # type: ignore
        from win32com.client import DispatchEx  # type: ignore
        from pywintypes import com_error  # type: ignore
    except ImportError as exc:  # pragma: no cover - environment specific
        raise RuntimeError(
            "pywin32 is required to build the XLAM. Install with 'pip install pywin32'."
        ) from exc
    return pythoncom, DispatchEx, com_error


def _remove_component_if_exists(vbproject, name: str) -> None:
    """Remove a VBA component by name if it exists and is removable."""
    vbext_ct_document = 100  # Document modules such as ThisWorkbook/Sheet1
    try:
        component = vbproject.VBComponents(name)
    except Exception:  # pragma: no cover - depends on Excel runtime
        return
    if component.Type == vbext_ct_document:
        return
        vbproject.VBComponents.Remove(component)


def _sync_code_module(component, source_path: Path) -> None:
    """Replace the entire contents of a VBA component with the given source file."""
    code = _read_text_with_fallback(source_path)
    code_module = component.CodeModule
    existing = code_module.CountOfLines
    if existing:
        code_module.DeleteLines(1, existing)
    code_module.AddFromString(code)


def _deactivate_addin_if_loaded(target_path: Path) -> Dict[str, Iterable[str]]:
    """Temporarily disable the add-in in running Excel instances to release file locks."""
    pythoncom, _, com_error = ensure_pywin32()
    from win32com.client import Dispatch  # type: ignore

    pythoncom.CoInitialize()
    disabled_addins: set[str] = set()
    excel = None
    existing_instance = True
    try:
        try:
            pythoncom.GetActiveObject("Excel.Application")
        except com_error:
            existing_instance = False
        excel = Dispatch("Excel.Application")
        normalized_target = str(target_path.resolve()).lower()
        original_alerts = getattr(excel, "DisplayAlerts", True)
        excel.DisplayAlerts = False

        try:
            for collection_name in ("AddIns2", "AddIns"):
                try:
                    collection = getattr(excel, collection_name)
                except AttributeError:
                    continue
                for idx in range(collection.Count, 0, -1):
                    try:
                        addin = collection.Item(idx)
                        full_name = getattr(addin, "FullName", "")
                    except Exception:
                        continue
                    if not full_name:
                        continue
                    try:
                        addin_path = str(Path(full_name).resolve()).lower()
                    except (OSError, ValueError):
                        addin_path = full_name.lower()
                    if addin_path != normalized_target:
                        continue
                    try:
                        if bool(getattr(addin, "Installed", False)):
                            addin.Installed = False
                            disabled_addins.add(addin_path)
                    except Exception:
                        continue

            workbooks = getattr(excel, "Workbooks", None)
            if workbooks is not None:
                for idx in range(workbooks.Count, 0, -1):
                    try:
                        workbook = workbooks.Item(idx)
                        full_name = getattr(workbook, "FullName", "")
                    except Exception:
                        continue
                    if not full_name:
                        continue
                    try:
                        workbook_path = str(Path(full_name).resolve()).lower()
                    except (OSError, ValueError):
                        workbook_path = full_name.lower()
                    if workbook_path == normalized_target:
                        try:
                            workbook.Close(SaveChanges=False)
                        except Exception:
                            continue
        finally:
            excel.DisplayAlerts = original_alerts

        if not existing_instance:
            try:
                excel.Quit()
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()

    if disabled_addins:
        return {"disabled_addins": tuple(disabled_addins), "excel_was_running": existing_instance}
    return {"disabled_addins": (), "excel_was_running": existing_instance}


def _reactivate_addins(info: Dict[str, Iterable[str]]) -> None:
    """Restore add-ins that were temporarily disabled."""
    disabled = tuple(info.get("disabled_addins", ()))
    if not disabled:
        return

    pythoncom, _, com_error = ensure_pywin32()
    from win32com.client import Dispatch  # type: ignore

    pythoncom.CoInitialize()
    excel = None
    try:
        try:
            excel = Dispatch("Excel.Application")
        except com_error:
            return

        original_alerts = getattr(excel, "DisplayAlerts", True)
        excel.DisplayAlerts = False

        try:
            for collection_name in ("AddIns2", "AddIns"):
                try:
                    collection = getattr(excel, collection_name)
                except AttributeError:
                    continue
                for idx in range(collection.Count, 0, -1):
                    try:
                        addin = collection.Item(idx)
                        full_name = getattr(addin, "FullName", "")
                    except Exception:
                        continue
                    if not full_name:
                        continue
                    try:
                        addin_path = str(Path(full_name).resolve()).lower()
                    except (OSError, ValueError):
                        addin_path = full_name.lower()
                    if addin_path in disabled:
                        try:
                            addin.Installed = True
                        except Exception:
                            continue
        finally:
            excel.DisplayAlerts = original_alerts

        if not info.get("excel_was_running", True):
            try:
                excel.Quit()
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()


def build_xlam(
    repo_dir: Path,
    filename: str,
    build_dir: Optional[Path] = None,
    excel_visible: bool = False,
) -> Path:
    """Compile VBA sources into an XLAM file using Excel automation."""
    sources = collect_vba_sources(repo_dir)
    workbook_module = repo_dir / "Microsoft Excel Objects" / "ThisWorkbook.cls"
    if not any(sources.values()) and not workbook_module.exists():
        raise RuntimeError("No VBA source files were found to compile.")

    pythoncom, DispatchEx, com_error = ensure_pywin32()
    pythoncom.CoInitialize()  # Ensure COM is initialized for the current thread

    excel = DispatchEx("Excel.Application")
    excel.Visible = bool(excel_visible)
    excel.DisplayAlerts = False

    build_root = build_dir or (repo_dir / DEFAULT_BUILD_DIRNAME)
    build_root.mkdir(parents=True, exist_ok=True)
    output_path = build_root / filename

    try:
        workbook = excel.Workbooks.Add()
        workbook.IsAddin = True
        vbproject = workbook.VBProject

        # Replace workbook module code
        if workbook_module.exists():
            _sync_code_module(vbproject.VBComponents("ThisWorkbook"), workbook_module)

        # Replace sheet modules
        for sheet_path in sources["sheets"]:
            module_name = _read_vb_name(sheet_path)
            if not module_name:
                continue
            try:
                component = vbproject.VBComponents(module_name)
            except com_error:
                # Create a new worksheet if the code name does not exist
                worksheet = workbook.Worksheets.Add()
                component = vbproject.VBComponents(worksheet.CodeName)
                try:
                    component.Name = module_name
                except Exception:
                    # Some Excel versions do not allow renaming document module codenames
                    pass
            _sync_code_module(component, sheet_path)

        # Remove default modules before re-importing
        removable: Iterable = list(vbproject.VBComponents)
        for component in removable:
            if component.Type == 1:  # Standard module
                vbproject.VBComponents.Remove(component)
            elif component.Type == 2:  # Class module
                vbproject.VBComponents.Remove(component)
            elif component.Type == 3:  # UserForm
                vbproject.VBComponents.Remove(component)

        # Import class modules, standard modules, and forms
        for path in sources["classes"]:
            module_name = _read_vb_name(path)
            if module_name:
                _remove_component_if_exists(vbproject, module_name)
            vbproject.VBComponents.Import(str(path))

        for path in sources["standard"]:
            module_name = _read_vb_name(path)
            if module_name:
                _remove_component_if_exists(vbproject, module_name)
            vbproject.VBComponents.Import(str(path))

        for path in sources["forms"]:
            module_name = _read_vb_name(path)
            if module_name:
                _remove_component_if_exists(vbproject, module_name)
            vbproject.VBComponents.Import(str(path))

        workbook.SaveAs(str(output_path), FileFormat=XLAM_FILE_FORMAT)
        workbook.Close(SaveChanges=False)
    except com_error as exc:
        raise RuntimeError(
            "Excel automation failed. Ensure Excel is installed and 'Trust access to the "
            "VBA project object model' is enabled."
        ) from exc
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    return output_path


def backup_existing(target_path: Path) -> Optional[Path]:
    """Back up the existing file if present, returning the backup path."""
    if not target_path.exists():
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = target_path.with_suffix(target_path.suffix + f".{timestamp}.bak")
    shutil.copy2(target_path, backup_path)
    return backup_path


def update_xlam(
    target_dir: str,
    filename: str,
    build_dir: Optional[str] = None,
    excel_visible: bool = False,
) -> Path:
    """Build the XLAM from source and copy it into the target add-in folder."""
    repo_dir = Path(__file__).resolve().parent
    target_directory = Path(target_dir).expanduser()
    target_directory.mkdir(parents=True, exist_ok=True)

    build_directory = Path(build_dir).expanduser() if build_dir else None
    built_path = build_xlam(repo_dir, filename, build_directory, excel_visible)

    target_path = target_directory / filename
    unlock_info = _deactivate_addin_if_loaded(target_path)
    disabled_addins = tuple(unlock_info.get("disabled_addins", ()))
    if disabled_addins:
        print("Temporarily disabled running add-in to release file lock.")

    backup_path = backup_existing(target_path)
    reactivated = False
    try:
        shutil.copy2(built_path, target_path)
    except PermissionError as exc:
        raise RuntimeError(
            f"Could not overwrite '{target_path}' because it is currently in use. "
            "Close Excel or disable the add-in, then rerun the script."
        ) from exc
    finally:
        if disabled_addins:
            _reactivate_addins(unlock_info)
            reactivated = True

    if reactivated:
        print("Restored add-in activation state.")

    if backup_path:
        print(f"Existing add-in backed up to: {backup_path}")
    print(f"Built add-in: {built_path}")
    print(f"Deployed add-in to: {target_path}")
    return target_path


def parse_args(argv: Optional[Sequence[str]] = None):
    parser = argparse.ArgumentParser(
        description="Build and deploy an Excel add-in (.xlam) from VBA sources."
    )
    parser.add_argument(
        "--target",
        help="Destination directory for the add-in (default: Excel Add-ins folder)",
    )
    parser.add_argument(
        "--filename",
        default="Vantage.xlam",
        help="Output add-in filename (default: Vantage.xlam)",
    )
    parser.add_argument(
        "--build-dir",
        help="Directory where the compiled add-in should be written before deployment (default: ./build)",
    )
    parser.add_argument(
        "--excel-visible",
        action="store_true",
        help="Show the Excel instance while building (useful for debugging).",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)

    target_dir = args.target or get_default_addins_folder()
    if not target_dir:
        print(
            "Could not determine Excel Add-ins folder automatically. Specify it with --target.",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        update_xlam(target_dir, args.filename, args.build_dir, args.excel_visible)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
