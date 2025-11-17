import argparse
import csv
import os
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import winreg  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover - windows specific
    winreg = None  # type: ignore

XLAM_FILE_FORMAT = 55  # Excel constant xlOpenXMLAddIn
DEFAULT_BUILD_DIRNAME = "build"
AddinUnlockInfo = Dict[str, Any]
DEFAULT_CSPROJ_RELATIVE = Path("VantagePackageHolder") / "VantagePackageHolder.csproj"
ADDIN_PROG_ID = "VantagePackageHolder.Addin"
ADDIN_CLSID = "{F5DA47BA-19D6-46CD-ACB7-BC918636925E}"
ADDIN_FRIENDLY_NAME = "Vantage Package Holder Add-in"
DEFAULT_REPO_ROOT = Path(r"C:\Users\andrep54\Vantage Add-in")


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


def _strip_document_metadata(code: str) -> str:
    """Remove export metadata that breaks document-module compilation."""
    lines = []
    for original in code.splitlines():
        stripped = original.strip()
        if not stripped:
            lines.append("")
            continue
        if stripped.startswith("VERSION "):
            continue
        if stripped.startswith("BEGIN"):
            continue
        if stripped == "END":
            continue
        if stripped.startswith("Attribute "):
            continue
        lines.append(original)

    while lines and not lines[0].strip():
        lines.pop(0)

    cleaned = "\n".join(lines)
    if cleaned and not cleaned.endswith("\n"):
        cleaned += "\n"
    return cleaned


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
    fallback = source_path.stem
    return fallback if fallback else None


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


def resolve_repo_dir(custom_root: Optional[str]) -> Path:
    """Determine which folder contains the workbook sources."""
    if custom_root:
        candidate = Path(custom_root).expanduser()
        if not candidate.exists():
            raise RuntimeError(f"Specified repo root does not exist: {candidate}")
        return candidate

    if DEFAULT_REPO_ROOT.exists():
        return DEFAULT_REPO_ROOT

    return Path(__file__).resolve().parent


def _resolve_csproj_path(repo_dir: Path, csproj_hint: Optional[str]) -> Optional[Path]:
    """Return a resolved csproj path if it exists."""
    if csproj_hint:
        candidate = Path(csproj_hint)
        if not candidate.is_absolute():
            candidate = repo_dir / candidate
        if not candidate.exists():
            raise RuntimeError(f"Specified COM project does not exist: {candidate}")
        return candidate

    default_path = repo_dir / DEFAULT_CSPROJ_RELATIVE
    return default_path if default_path.exists() else None


def _read_assembly_name(csproj_path: Path) -> str:
    """Parse the target assembly name from a csproj file."""
    try:
        tree = ET.parse(csproj_path)
    except (ET.ParseError, OSError) as exc:
        raise RuntimeError(f"Failed to parse project file: {csproj_path}") from exc

    root = tree.getroot()
    namespace = {"msb": "http://schemas.microsoft.com/developer/msbuild/2003"}
    for node in root.findall(".//msb:AssemblyName", namespace):
        if node.text and node.text.strip():
            return node.text.strip()
    return csproj_path.stem


def _resolve_msbuild_command(msbuild_path: Optional[str] = None) -> List[str]:
    """Locate an MSBuild executable (or dotnet fallback) and return the command list."""
    if msbuild_path:
        full_path = Path(msbuild_path)
        if not full_path.exists():
            raise RuntimeError(f"msbuild executable not found: {msbuild_path}")
        return [str(full_path)]

    env_path = os.environ.get("MSBUILD_EXE_PATH")
    if env_path and Path(env_path).exists():
        return [env_path]

    which_msbuild = shutil.which("msbuild")
    if which_msbuild:
        return [which_msbuild]

    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    vswhere = (
        Path(program_files_x86)
        / "Microsoft Visual Studio"
        / "Installer"
        / "vswhere.exe"
    )
    if vswhere.exists():
        try:
            output = subprocess.check_output(
                [
                    str(vswhere),
                    "-latest",
                    "-requires",
                    "Microsoft.Component.MSBuild",
                    "-find",
                    "MSBuild\\**\\Bin\\MSBuild.exe",
                ],
                text=True,
            ).strip()
        except subprocess.CalledProcessError:
            output = ""
        for line in output.splitlines():
            candidate = line.strip()
            if candidate and Path(candidate).exists():
                return [candidate]

    dotnet = shutil.which("dotnet")
    if dotnet:
        return [dotnet, "msbuild"]

    raise RuntimeError(
        "Could not locate MSBuild. Install Visual Studio Build Tools or specify --msbuild."
    )


def _run_subprocess(command: Sequence[str], cwd: Optional[Path] = None) -> None:
    """Run a command and raise a RuntimeError with context if it fails."""
    try:
        subprocess.run(command, check=True, cwd=str(cwd) if cwd else None)
    except FileNotFoundError as exc:
        raise RuntimeError(f"Command not found: {command[0]}") from exc
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            f"Command failed (exit code {exc.returncode}): {' '.join(command)}"
        ) from exc


def build_com_addin(
    repo_dir: Path,
    csproj_hint: Optional[str],
    configuration: str,
    platform: str,
    msbuild_path: Optional[str],
) -> Optional[Path]:
    """Build the COM add-in project if present and return the output DLL path."""
    csproj_path = _resolve_csproj_path(repo_dir, csproj_hint)
    if csproj_path is None:
        return None

    command = _resolve_msbuild_command(msbuild_path)
    command = list(command) + [
        str(csproj_path),
        f"/p:Configuration={configuration}",
        f"/p:Platform={platform}",
        "/t:Build",
    ]
    print(f"Building COM add-in project: {' '.join(command)}")
    _run_subprocess(command, cwd=csproj_path.parent)

    assembly_name = _read_assembly_name(csproj_path)
    output_path = (
        csproj_path.parent / "bin" / configuration / f"{assembly_name}.dll"
    )
    if not output_path.exists():
        raise RuntimeError(
            f"MSBuild completed but the expected output was not found: {output_path}"
        )
    return output_path


def stage_com_binary(built_path: Path, keep: int = 5) -> Path:
    """Temporarily returns the original DLL path (staging disabled)."""
    return built_path


def _resolve_regasm_command(regasm_path: Optional[str]) -> List[str]:
    """Locate regasm.exe if registration is requested."""
    if regasm_path:
        candidate = Path(regasm_path)
        if not candidate.exists():
            raise RuntimeError(f"regasm executable not found: {regasm_path}")
        return [str(candidate)]

    which_regasm = shutil.which("regasm")
    if which_regasm:
        return [which_regasm]

    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    search_roots = [
        windir / "Microsoft.NET" / "Framework64",
        windir / "Microsoft.NET" / "Framework",
    ]
    for root in search_roots:
        for regasm in sorted(root.glob("v*/RegAsm.exe"), reverse=True):
            if regasm.exists():
                return [str(regasm)]

    raise RuntimeError(
        "regasm.exe not found. Provide its path with --regasm or install the .NET Framework developer tools."
    )


def register_com_addin(dll_path: Path, regasm_path: Optional[str]) -> None:
    """Register the compiled COM add-in using regasm."""
    command = _resolve_regasm_command(regasm_path)
    command = list(command) + [str(dll_path), "/codebase", "/nologo"]
    print(f"Registering COM add-in: {' '.join(command)}")
    _run_subprocess(command)


def register_com_addin_per_user(dll_path: Path, regasm_path: Optional[str]) -> None:
    """Register the COM add-in for the current user without requiring elevation."""
    if winreg is None:
        raise RuntimeError("winreg is unavailable on this platform; per-user registration is not supported.")

    reg_script = _generate_reg_script(dll_path, regasm_path)
    entries = _parse_reg_script(reg_script)
    _apply_reg_entries(entries, codebase=str(dll_path.as_uri()))
    _write_excel_addin_key(codebase=str(dll_path.as_uri()))
    print("Registered COM add-in for the current user.")


def _generate_reg_script(dll_path: Path, regasm_path: Optional[str]) -> str:
    command = _resolve_regasm_command(regasm_path)
    with tempfile.TemporaryDirectory() as tmpdir:
        reg_path = Path(tmpdir) / "addin.reg"
        cmd = list(command) + [
            str(dll_path),
            f"/regfile:{reg_path}",
            "/codebase",
            "/nologo",
        ]
        _run_subprocess(cmd)
        if reg_path.exists():
            return reg_path.read_text(encoding="utf-8", errors="ignore")
        default_reg = dll_path.with_suffix(".reg")
        if default_reg.exists():
            return default_reg.read_text(encoding="utf-8", errors="ignore")
        raise RuntimeError("regasm did not produce a registry file; cannot continue.")


def _parse_reg_script(script: str) -> List[Tuple[str, List[str]]]:
    entries: List[Tuple[str, List[str]]] = []
    current_key: Optional[str] = None
    current_values: List[str] = []
    for raw_line in script.splitlines():
        line = raw_line.strip()
        if not line or line.startswith(";") or line.upper().startswith("REGEDIT"):
            continue
        if line.startswith("[") and line.endswith("]"):
            if current_key is not None:
                entries.append((current_key, current_values))
            current_key = line[1:-1]
            current_values = []
        else:
            current_values.append(line)
    if current_key is not None:
        entries.append((current_key, current_values))
    return entries


def _map_reg_key(raw_key: str) -> Optional[Tuple[int, str]]:
    lowered = raw_key.lower()
    if lowered.startswith("hkey_classes_root\\"):
        sub = raw_key.split("\\", 1)[1]
        return winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{sub}"
    if lowered.startswith("hkey_local_machine\\software\\classes\\"):
        sub = raw_key.split("classes\\", 1)[1]
        return winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{sub}"
    if lowered.startswith("hkey_local_machine\\software\\microsoft\\office\\"):
        sub = raw_key.split("microsoft\\office\\", 1)[1]
        return winreg.HKEY_CURRENT_USER, f"Software\\Microsoft\\Office\\{sub}"
    if lowered.startswith("hkey_current_user\\"):
        sub = raw_key.split("\\", 1)[1]
        return winreg.HKEY_CURRENT_USER, sub
    return None


def _unescape_reg_string(token: str) -> str:
    value = token[1:-1]
    value = value.replace("\\\\", "\\").replace("\\\"", "\"")
    return value


def _apply_reg_entries(entries: List[Tuple[str, List[str]]], codebase: str) -> None:
    for raw_key, raw_values in entries:
        mapped = _map_reg_key(raw_key)
        if not mapped:
            continue
        root, subkey = mapped
        with winreg.CreateKeyEx(root, subkey, 0, winreg.KEY_WRITE) as handle:
            for raw_value in raw_values:
                if raw_value.startswith("@="):
                    name = ""
                    data = raw_value[2:]
                else:
                    if "=" not in raw_value:
                        continue
                    name_part, data = raw_value.split("=", 1)
                    name = name_part.strip('"')

                if data.startswith('"') and data.endswith('"'):
                    value = _unescape_reg_string(data)
                    if name.lower() == "codebase":
                        value = codebase
                    winreg.SetValueEx(handle, name, 0, winreg.REG_SZ, value)
                elif data.lower().startswith("dword:"):
                    winreg.SetValueEx(handle, name, 0, winreg.REG_DWORD, int(data[6:], 16))
                else:
                    raise RuntimeError(f"Unsupported registry value format: {raw_value}")


def _write_excel_addin_key(codebase: str) -> None:
    friendly = ADDIN_FRIENDLY_NAME
    target = r"Software\Microsoft\Office\Excel\Addins\{}".format(ADDIN_PROG_ID)
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, target, 0, winreg.KEY_WRITE) as handle:
        winreg.SetValueEx(handle, "FriendlyName", 0, winreg.REG_SZ, friendly)
        winreg.SetValueEx(handle, "Description", 0, winreg.REG_SZ, friendly)
        winreg.SetValueEx(handle, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(handle, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(handle, "CLSID", 0, winreg.REG_SZ, ADDIN_CLSID)


def _list_excel_processes() -> List[int]:
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        return []

    pids: List[int] = []
    for row in csv.reader(result.stdout.splitlines()):
        if not row:
            continue
        name = row[0].strip().strip('"')
        if name.upper() != "EXCEL.EXE":
            continue
        try:
            pids.append(int(row[1]))
        except (IndexError, ValueError):
            continue
    return pids


def _ensure_excel_closed(timeout: float = 10.0) -> None:
    pids = _list_excel_processes()
    if not pids:
        return

    print("Closing Excel to release COM add-in DLL...")
    for pid in pids:
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    deadline = time.time() + timeout
    while time.time() < deadline:
        if not _list_excel_processes():
            print("Excel closed.")
            return
        time.sleep(0.5)

    raise RuntimeError("Excel is still running. Close all Excel instances and rerun the command.")


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
    raw_code = _read_text_with_fallback(source_path)
    code = _strip_document_metadata(raw_code)
    code_module = component.CodeModule
    existing = code_module.CountOfLines
    if existing:
        code_module.DeleteLines(1, existing)
    code_module.AddFromString(code)


def _deactivate_addin_if_loaded(target_path: Path) -> AddinUnlockInfo:
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


def _reactivate_addins(info: AddinUnlockInfo) -> None:
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


def backup_existing(target_path: Path, backup_dir: Optional[Path] = None) -> Optional[Path]:
    """Back up the existing file if present, returning the backup path."""
    if not target_path.exists():
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if backup_dir:
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_filename = f"{target_path.stem}_{timestamp}{target_path.suffix}"
        backup_path = backup_dir / backup_filename
    else:
        backup_path = target_path.with_suffix(target_path.suffix + f".{timestamp}.bak")
    shutil.copy2(target_path, backup_path)
    return backup_path


def update_xlam(
    target_dir: str,
    filename: str,
    build_dir: Optional[str] = None,
    excel_visible: bool = False,
    repo_root: Optional[str] = None,
    csproj_path: Optional[str] = None,
    com_configuration: str = "Debug",
    com_platform: str = "AnyCPU",
    msbuild_path: Optional[str] = None,
    skip_com_build: bool = False,
    register_com: bool = False,
    register_com_per_user: bool = False,
    close_excel: bool = False,
    regasm_path: Optional[str] = None,
) -> Path:
    """Build the COM add-in (optional), create the XLAM, and deploy both."""
    if register_com and register_com_per_user:
        raise RuntimeError("Choose either --register-com or --register-com-per-user, not both.")
    if register_com_per_user and winreg is None:
        raise RuntimeError("winreg is unavailable; per-user registration is not supported on this platform.")
    if (register_com or register_com_per_user) and skip_com_build:
        raise RuntimeError(
            "Registration requires building the COM project. Remove --skip-com-build or the registration flag."
        )

    if close_excel:
        _ensure_excel_closed()

    repo_dir = resolve_repo_dir(repo_root)
    com_output: Optional[Path] = None
    if not skip_com_build:
        com_output = build_com_addin(
            repo_dir,
            csproj_path,
            com_configuration,
            com_platform,
            msbuild_path,
        )
        if com_output:
            com_output = stage_com_binary(com_output)
            print(f"COM add-in staged at: {com_output}")
        else:
            print("No COM add-in project detected. Skipping COM build.")

        if register_com:
            if com_output is None:
                raise RuntimeError("Cannot register COM add-in because no DLL was built.")
            register_com_addin(com_output, regasm_path)
        elif register_com_per_user:
            if com_output is None:
                raise RuntimeError("Cannot register COM add-in because no DLL was built.")
            register_com_addin_per_user(com_output, regasm_path)

    target_directory = Path(target_dir).expanduser()
    target_directory.mkdir(parents=True, exist_ok=True)

    build_directory = Path(build_dir).expanduser() if build_dir else None
    built_path = build_xlam(repo_dir, filename, build_directory, excel_visible)

    target_path = target_directory / filename
    unlock_info = _deactivate_addin_if_loaded(target_path)
    disabled_addins = tuple(unlock_info.get("disabled_addins", ()))
    if disabled_addins:
        print("Temporarily disabled running add-in to release file lock.")

    backup_dir = built_path.parent / "backups"
    backup_path = backup_existing(target_path, backup_dir)
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
    parser.add_argument(
        "--repo-root",
        help=(
            "Root directory for the VBA/C# sources "
            "(default: C:\\Users\\andrep54\\Vantage Add-in if it exists, otherwise the script directory)."
        ),
    )
    parser.add_argument(
        "--csproj",
        help=(
            "Path to the COM add-in .csproj to build first "
            "(default: ./VantagePackageHolder/VantagePackageHolder.csproj if it exists)."
        ),
    )
    parser.add_argument(
        "--skip-com-build",
        action="store_true",
        help="Skip building the COM add-in even if a project file is present.",
    )
    parser.add_argument(
        "--configuration",
        default="Debug",
        help="MSBuild Configuration for the COM add-in build (default: Debug).",
    )
    parser.add_argument(
        "--platform",
        default="AnyCPU",
        help="MSBuild Platform for the COM add-in build (default: AnyCPU).",
    )
    parser.add_argument(
        "--msbuild",
        help="Full path to msbuild.exe (default: auto-detect).",
    )
    parser.add_argument(
        "--register-com",
        action="store_true",
        help="Register the compiled COM add-in using regasm after a successful build.",
    )
    parser.add_argument(
        "--register-com-per-user",
        action="store_true",
        help="Register the COM add-in only for the current user (no admin rights required).",
    )
    parser.add_argument(
        "--close-excel",
        action="store_true",
        help="Terminate running Excel instances before building to release locked DLLs.",
    )
    parser.add_argument(
        "--regasm",
        help="Full path to regasm.exe (default: auto-detect).",
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
        update_xlam(
            target_dir,
            args.filename,
            args.build_dir,
            args.excel_visible,
            repo_root=args.repo_root,
            csproj_path=args.csproj,
            com_configuration=args.configuration,
            com_platform=args.platform,
            msbuild_path=args.msbuild,
            skip_com_build=args.skip_com_build,
            register_com=args.register_com,
            register_com_per_user=args.register_com_per_user,
            close_excel=args.close_excel,
            regasm_path=args.regasm,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
