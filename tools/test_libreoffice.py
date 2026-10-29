"""
Smoke-test the VBA modules without Excel, using LibreOffice in VBA-compatibility mode.

    python tools/test_libreoffice.py [output_folder] [locale]     # locale default en-US
    python tools/test_libreoffice.py out_es es-ES                 # Spanish: comma decimals
    python tools/reference_check.py [output_folder]/InvoiceBackup_RESULT.xlsx

Copies the sample workbook to a folder, injects the .bas modules, runs
GenerateEntityTabs twice (the second run must replace, not duplicate) and both
carrier exports, then saves the result. MsgBox calls are redirected to a text
file so nothing blocks, and the messages are printed at the end.

LibreOffice is not Excel. This catches syntax and logic errors; it does not
replace a run in Excel. Known differences: LibreOffice ignores FileFormat:=51
and writes the exported carrier files in the old .xls format (Excel writes
real .xlsx), and some formatting (header font colour) is not applied.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import uno  # noqa: F401  (LibreOffice python bridge)
from com.sun.star.beans import PropertyValue

ROOT = Path(__file__).resolve().parents[1]
MODULES = ["modHelpers", "modEntityTabs", "modCarrierExport"]
MACROS = [
    "modEntityTabs.GenerateEntityTabs",
    "modEntityTabs.GenerateEntityTabs",      # run twice: second run must replace, not duplicate
    "modCarrierExport.ExportPendingInvoices",
    "modCarrierExport.ExportRejectedInvoices",
]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def vba_for_libreoffice(path: Path) -> str:
    lines = path.read_text(encoding="utf-8").splitlines()
    lines = [l for l in lines if not l.startswith("Attribute VB_")]
    code = "\n".join(lines).replace("MsgBox ", "TestLog ")
    return "Option VBASupport 1\n" + code


def log_module(log_path: Path) -> str:
    """MsgBox replacement: append each message to a text file the test can read."""
    return f'''Option VBASupport 1
Public Sub TestLog(ByVal msg As String, Optional ByVal style As Variant, Optional ByVal title As Variant)
    Dim f As Integer
    f = FreeFile
    Open "{log_path}" For Append As #f
    Print #f, msg
    Close #f
End Sub
'''


def main(out_dir: Path, locale: str = "en-US") -> bool:
    out_dir.mkdir(parents=True, exist_ok=True)
    work = out_dir / "InvoiceBackup_TEST.xlsx"
    shutil.copy(ROOT / "InvoiceBackup_TEMPLATE.xlsx", work)

    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         "--accept=socket,host=127.0.0.1,port=2002;urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            sys.exit("could not connect to LibreOffice")
        # run with the requested locale (e.g. es-ES uses a comma as decimal separator,
        # which is how a Spanish Excel behaves - this is where a formula bug was found)
        cfg = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        l10n = cfg.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess",
                                               (prop("nodepath", "/org.openoffice.Setup/L10N"),))
        l10n.setPropertyValue("ooSetupSystemLocale", locale)
        l10n.commitChanges()
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(work.as_uri(), "_blank", 0,
                                           (prop("Hidden", False), prop("MacroExecutionMode", 4)))

        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        libs.VBACompatibilityMode = True
        lib = libs.getByName("Standard")
        log_path = out_dir / "macro_messages.txt"
        log_path.unlink(missing_ok=True)
        if lib.hasByName("modTestLog"):
            lib.removeByName("modTestLog")
        lib.insertByName("modTestLog", log_module(log_path))
        for m in MODULES:
            code = vba_for_libreoffice(ROOT / "src" / f"{m}.bas")
            if lib.hasByName(m):
                lib.replaceByName(m, code)
            else:
                lib.insertByName(m, code)

        sp = doc.getScriptProvider()
        for macro in MACROS:
            t0 = time.time()
            script = sp.getScript(f"vnd.sun.star.script:Standard.{macro}?language=Basic&location=document")
            script.invoke((), (), ())
            print(f"ran {macro} in {time.time() - t0:.1f}s")

        messages = log_path.read_text() if log_path.exists() else ""
        print(f"--- macro messages ({locale}) ---")
        print(messages)

        # every status formula in CHECK must use a dot decimal (formulas are stored in English)
        check = doc.Sheets.getByName("CHECK")
        bad = []
        for r in range(5, 200):
            f = check.getCellByPosition(4, r).getFormula()
            if "ABS(" in f.upper() and "<0.005" not in f:
                bad.append(f)
        ok = "Stopped" not in messages and not bad
        print("CHECK formulas:", "OK" if not bad else f"BAD {bad[:3]}")
        result = out_dir / "InvoiceBackup_RESULT.xlsx"
        doc.storeToURL(result.as_uri(), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        print(f"saved {result}")
        return ok
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp(prefix="ibs_test_"))
    loc = sys.argv[2] if len(sys.argv) > 2 else "en-US"
    passed = main(target, loc)
    print(f"output folder: {target}")
    print("TEST", "PASSED" if passed else "FAILED")
    sys.exit(0 if passed else 1)
