"""VBA module operations and macro execution.

Security model
--------------
This entire tool is opt-in:
  - Env gate: THEPEXCEL_MCP_ENABLE_VBA=1 must be set, or every action raises.
  - AccessVBOM pre-flight: checked per call via winreg before any COM touch.
    HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Security → AccessVBOM must == 1.
    Missing key → also raises with exact enable instructions.

VBA COM chain
-------------
    app.VBE                            → VBE object
    app.VBE.VBProjects(wb.Name)        → VBProject for the workbook
    vbproject.VBComponents             → collection of modules
    vbcomponent.CodeModule             → source code accessor

Module type constants (vba.vbext_ct_*)
    1 = vbext_ct_StdModule (standard module)
    2 = vbext_ct_ClassModule
    3 = vbext_ct_MSForm (UserForm)
    100 = vbext_ct_Document (Sheet/ThisWorkbook)

write_module creates standard modules only (type 1).
Running code in unsaved/new workbooks works fine; saving to .xlsm is the
user's responsibility (COM cannot coerce a .xlsx to .xlsm automatically).
"""

from __future__ import annotations

import os

try:
    import winreg  # Windows-only; absent on non-Windows test envs
except ImportError:
    winreg = None  # type: ignore[assignment]

from fastmcp.exceptions import ToolError

from ..session import ExcelSession

_session = ExcelSession()

# Module type code → readable label
_MODULE_TYPE = {
    1: "standard",
    2: "class",
    3: "form",
    100: "document",
}

_VBEXT_CT_STD_MODULE = 1  # only type that can be freely created/deleted


def _env_gate() -> None:
    """Raise ToolError if THEPEXCEL_MCP_ENABLE_VBA is not set to '1'."""
    if os.environ.get("THEPEXCEL_MCP_ENABLE_VBA") != "1":
        raise ToolError(
            "VBA tool disabled. Set THEPEXCEL_MCP_ENABLE_VBA=1 to enable. "
            "Example: set THEPEXCEL_MCP_ENABLE_VBA=1 in your environment, "
            "then restart the MCP server."
        )


def _access_vbom_preflight() -> None:
    """Raise ToolError if Excel's 'Trust access to VBA project object model' is not enabled.

    Reads HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Security\\AccessVBOM.
    Value must be 1 (DWORD).

    Enable path: Excel → File → Options → Trust Center → Trust Center Settings →
    Macro Settings → check "Trust access to the VBA project object model".
    """
    try:
        if winreg is None:
            pass  # non-Windows — fall through to error
        else:
            key_path = r"SOFTWARE\Microsoft\Office\16.0\Excel\Security"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                if value == 1:
                    return
    except FileNotFoundError:
        pass  # key/value absent → not enabled
    except Exception:
        pass  # any other error → treat as disabled (safe default)

    raise ToolError(
        "Excel VBA project model access is not enabled. "
        "To enable: Excel → File → Options → Trust Center → "
        "Trust Center Settings → Macro Settings → "
        'check "Trust access to the VBA project object model" → OK → restart Excel. '
        "Note: This is a user trust decision — the MCP server will never modify "
        "your registry or trust settings."
    )


def vba_action(
    action: str,
    workbook: str | None = None,
    module_name: str | None = None,
    code: str | None = None,
    proc_name: str | None = None,
    args: list | None = None,
) -> dict:
    """Dispatch a VBA action.

    Actions
    -------
    list_modules
        List all VBA modules: name, type (standard/class/form/document), line count.
    get_module
        Return full source code of a module. Requires ``module_name``.
    write_module
        Create or replace a standard module with ``code``.
        Requires ``module_name`` and ``code``.
        If the module exists, ALL existing lines are replaced.
        If it does not exist, a new standard module is created.
        Only standard modules (type=1) can be created/replaced this way.
        Note: running macros in unsaved/new workbooks works; to SAVE a workbook
        with VBA you must use a macro-enabled format (.xlsm/.xlsb).
    delete_module
        Delete a standard module. Requires ``module_name``.
        Document modules (Sheet*, ThisWorkbook) cannot be deleted.
    run
        Execute a Sub or Function via Application.Run.
        Requires ``proc_name`` (e.g. ``"Module1.MyProc"`` or just ``"MyProc"``).
        ``args`` is an optional list of positional arguments passed to the macro.
        Functions return their value; Subs return null.
        COM variant limits apply: args must be JSON-serialisable scalars
        (string, int, float, bool, None). Objects and arrays are not supported.
        The macro must be in the active workbook or a loaded add-in.
    """
    _env_gate()
    _access_vbom_preflight()

    if action == "list_modules":
        return _session.run_com(_list_modules, workbook)
    if action == "get_module":
        if not module_name:
            raise ToolError("action='get_module' requires 'module_name'.")
        return _session.run_com(_get_module, workbook, module_name)
    if action == "write_module":
        if not module_name:
            raise ToolError("action='write_module' requires 'module_name'.")
        if code is None:
            raise ToolError("action='write_module' requires 'code'.")
        return _session.run_com(_write_module, workbook, module_name, code)
    if action == "delete_module":
        if not module_name:
            raise ToolError("action='delete_module' requires 'module_name'.")
        return _session.run_com(_delete_module, workbook, module_name)
    if action == "run":
        if not proc_name:
            raise ToolError("action='run' requires 'proc_name'.")
        return _session.run_com(_run, workbook, proc_name, args or [])
    raise ToolError(
        f"Unknown action '{action}'. Valid: list_modules, get_module, "
        "write_module, delete_module, run."
    )


# ── COM helpers (run inside the worker thread) ────────────────────────────────

def _get_vbproject(workbook: str | None):
    """Return the VBProject COM object for the given workbook."""
    app = _session.get_app()
    wb = _session.get_workbook(workbook)
    try:
        vbe = app.VBE
        # VBProjects is keyed by project name, not workbook name directly
        for i in range(1, vbe.VBProjects.Count + 1):
            proj = vbe.VBProjects.Item(i)
            if proj.FileName.lower().endswith(
                (wb.FullName.lower(), wb.Name.lower())
            ) or proj.Name == wb.Name.replace(".xlsm", "").replace(".xlsb", ""):
                return proj
        # Fallback: return first project that matches any open workbook by filename
        for i in range(1, vbe.VBProjects.Count + 1):
            proj = vbe.VBProjects.Item(i)
            try:
                if proj.FileName == wb.FullName:
                    return proj
            except Exception:
                pass
        # Last resort: try direct collection access by workbook name key
        try:
            return vbe.VBProjects(wb.Name)
        except Exception:
            pass
        raise ToolError(
            f"VBProject for '{wb.Name}' not found. "
            "Ensure the workbook is open and AccessVBOM is enabled."
        )
    except ToolError:
        raise
    except Exception as e:
        raise _session.wrap(e, "Cannot access VBE")


def _list_modules(workbook: str | None) -> dict:
    vbp = _get_vbproject(workbook)
    modules = []
    for i in range(1, vbp.VBComponents.Count + 1):
        comp = vbp.VBComponents.Item(i)
        line_count = 0
        try:
            line_count = comp.CodeModule.CountOfLines
        except Exception:
            pass
        modules.append({
            "name": comp.Name,
            "type": _MODULE_TYPE.get(comp.Type, str(comp.Type)),
            "line_count": line_count,
        })
    return {"modules": modules, "count": len(modules)}


def _get_module(workbook: str | None, module_name: str) -> dict:
    vbp = _get_vbproject(workbook)
    comp = _find_component(vbp, module_name)
    cm = comp.CodeModule
    count = cm.CountOfLines
    code = cm.Lines(1, count) if count > 0 else ""
    return {
        "module": module_name,
        "type": _MODULE_TYPE.get(comp.Type, str(comp.Type)),
        "line_count": count,
        "code": code,
    }


def _write_module(workbook: str | None, module_name: str, code: str) -> dict:
    vbp = _get_vbproject(workbook)
    # Find or create
    comp = None
    for i in range(1, vbp.VBComponents.Count + 1):
        c = vbp.VBComponents.Item(i)
        if c.Name.lower() == module_name.lower():
            if c.Type != _VBEXT_CT_STD_MODULE:
                raise ToolError(
                    f"Module '{module_name}' exists but is type "
                    f"{_MODULE_TYPE.get(c.Type, c.Type)} — only standard modules "
                    "can be written. Delete or choose a different name."
                )
            comp = c
            break
    if comp is None:
        comp = vbp.VBComponents.Add(_VBEXT_CT_STD_MODULE)
        comp.Name = module_name
    # Replace all lines
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code)
    return {
        "written": module_name,
        "line_count": comp.CodeModule.CountOfLines,
    }


def _delete_module(workbook: str | None, module_name: str) -> dict:
    vbp = _get_vbproject(workbook)
    comp = _find_component(vbp, module_name)
    if comp.Type != _VBEXT_CT_STD_MODULE:
        raise ToolError(
            f"Cannot delete module '{module_name}' (type="
            f"{_MODULE_TYPE.get(comp.Type, comp.Type)}). "
            "Only standard modules can be deleted. "
            "Document modules (Sheet*, ThisWorkbook) are permanent."
        )
    vbp.VBComponents.Remove(comp)
    return {"deleted": module_name}


def _run(workbook: str | None, proc_name: str, args: list) -> dict:
    app = _session.get_app()
    # Ensure the correct workbook is active context (Application.Run scopes to active wb)
    if workbook:
        wb = _session.get_workbook(workbook)
        wb.Activate()
    try:
        if args:
            result = app.Run(proc_name, *args)
        else:
            result = app.Run(proc_name)
        return {"proc": proc_name, "result": result}
    except Exception as e:
        raise _session.wrap(e, f"Run '{proc_name}' failed")


def _find_component(vbp, module_name: str):
    """Return a VBComponent or raise ToolError with available module names."""
    for i in range(1, vbp.VBComponents.Count + 1):
        c = vbp.VBComponents.Item(i)
        if c.Name.lower() == module_name.lower():
            return c
    available = [vbp.VBComponents.Item(i).Name for i in range(1, vbp.VBComponents.Count + 1)]
    raise ToolError(
        f"Module '{module_name}' not found. Available modules: {available}"
    )
