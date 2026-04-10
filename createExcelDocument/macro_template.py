"""
Build orders_template.xlsm on Windows using Excel COM (one-time automation).

Requires: Excel installed, pywin32, and temporary AccessVBOM registry access
("Trust access to the VBA project object model") — enabled only for the build.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

if sys.platform == "win32":
    import winreg
else:
    winreg = None  # type: ignore[assignment]

CLIPBOARD_INI_NAME = "excel_clipboard_launch.ini"

# Workbook_SheetFollowHyperlink: Copy Path uses # in-sheet links; file URI in column 28 (AB).
# Tracking URLs for each row live in hidden columns 29…43 (AC…AQ). "View Link List" collects them,
# writes a temp UTF-8 .txt plus optional .ctx.tsv (company, order, … from row headers),
# and runs VIEWER hidden (window style 0) from ini (Python + tkinter grid).
# Reads UTF-8 ini (PY=, SCRIPT=, VIEWER=) from AA1 / excel_clipboard_launch.ini.
THISWORKBOOK_VBA = r'''Option Explicit

Private Const COL_TRACK_URI_START As Long = 29
Private Const COL_TRACK_URI_END As Long = 43

Private Function ReadUtf8File(ByVal path As String) As String
    Dim stm As Object
    On Error GoTo CleanFail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8File = stm.ReadText
    stm.Close
    Exit Function
CleanFail:
    ReadUtf8File = ""
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
End Function

Private Sub WriteUtf8File(ByVal path As String, ByVal content As String)
    Dim stm As Object
    On Error GoTo CleanFail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile path, 2
    stm.Close
    Exit Sub
CleanFail:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
End Sub

Private Function CollectTrackingUrlsForRow(ByVal Sh As Object, ByVal rowNum As Long) As String
    Dim c As Long
    Dim v As Variant
    Dim s As String
    Dim body As String
    body = ""
    For c = COL_TRACK_URI_START To COL_TRACK_URI_END
        v = Sh.Cells(rowNum, c).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            s = Trim(CStr(v))
            If Len(s) > 0 Then
                body = body & s & vbLf
            End If
        End If
    Next c
    CollectTrackingUrlsForRow = body
End Function

Private Function HeaderColumn(ByVal Sh As Object, ByVal want As String) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim h As String
    On Error Resume Next
    lastCol = Sh.Cells(1, Sh.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    For c = 1 To lastCol
        h = Trim(CStr(Sh.Cells(1, c).Value))
        If StrComp(h, want, vbTextCompare) = 0 Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
    HeaderColumn = 0
End Function

Private Function CtxLine(ByVal key As String, ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then
        CtxLine = ""
        Exit Function
    End If
    If IsEmpty(v) Then
        CtxLine = ""
        Exit Function
    End If
    s = Trim(CStr(v))
    If Len(s) = 0 Then
        CtxLine = ""
        Exit Function
    End If
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(9), " ")
    CtxLine = key & Chr(9) & s & vbLf
End Function

Private Sub WriteTrackingContextTsv(ByVal Sh As Object, ByVal rowNum As Long, ByVal path As String)
    Dim c As Long
    Dim body As String
    body = ""
    c = HeaderColumn(Sh, "Company")
    If c > 0 Then body = body & CtxLine("company", Sh.Cells(rowNum, c).Value)
    c = HeaderColumn(Sh, "Order Number")
    If c > 0 Then body = body & CtxLine("order_number", Sh.Cells(rowNum, c).Value)
    c = HeaderColumn(Sh, "Category")
    If c > 0 Then body = body & CtxLine("category", Sh.Cells(rowNum, c).Value)
    c = HeaderColumn(Sh, "Purchase Date")
    If c > 0 Then body = body & CtxLine("purchase_datetime", Sh.Cells(rowNum, c).Value)
    c = HeaderColumn(Sh, "Email")
    If c > 0 Then body = body & CtxLine("email", Sh.Cells(rowNum, c).Value)
    If Len(body) > 0 Then
        Call WriteUtf8File(path, body)
    End If
End Sub

Private Sub LaunchTrackingLinkViewerForRow(ByVal Sh As Object, ByVal rowNum As Long)
    Dim body As String
    Dim tempPath As String
    Dim fso As Object
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim viewer As String
    Dim cmd As String
    Dim shell As Object

    body = CollectTrackingUrlsForRow(Sh, rowNum)
    If Len(Trim(body)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2) & "\email_sorter_tracking_r" & rowNum & "_t" & CLng(Timer * 10000) & ".txt"
    Call WriteUtf8File(tempPath, body)
    Call WriteTrackingContextTsv(Sh, rowNum, Replace(tempPath, ".txt", ".ctx.tsv"))

    iniPath = Trim(CStr(Sh.Range("AA1").Value))
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then
        If Len(ThisWorkbook.Path) > 0 Then
            iniPath = ThisWorkbook.Path & Application.PathSeparator & "excel_clipboard_launch.ini"
        End If
    End If
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then Exit Sub

    allText = ReadUtf8File(iniPath)
    py = IniValue(allText, "PY")
    viewer = IniValue(allText, "VIEWER")
    If Len(py) = 0 Or Len(viewer) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(viewer) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & viewer & Chr(34) & " " & Chr(34) & tempPath & Chr(34)
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
End Sub

Private Function IniValue(ByVal allText As String, ByVal key As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim prefix As String
    prefix = UCase(key) & "="
    lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If UCase(Left(line, Len(prefix))) = prefix Then
            IniValue = Trim(Mid(line, Len(prefix) + 1))
            Exit Function
        End If
    Next i
    IniValue = ""
End Function

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Const COL_FILE_URI As Long = 28
    Dim header As String
    Dim uri As String
    Dim py As String
    Dim scriptPath As String
    Dim cmd As String
    Dim shell As Object
    Dim iniPath As String
    Dim allText As String
    Dim fso As Object

    On Error GoTo CleanFail

    header = CStr(Sh.Cells(1, Target.Range.Column).Value)

    If StrComp(header, "View Link List", vbTextCompare) = 0 Then
        Call LaunchTrackingLinkViewerForRow(Sh, Target.Range.Row)
        Exit Sub
    End If

    If StrComp(header, "Copy Path", vbTextCompare) <> 0 Then Exit Sub

    uri = CStr(Sh.Cells(Target.Range.Row, COL_FILE_URI).Value)
    uri = Trim(uri)
    If Len(uri) = 0 Then GoTo CleanFail
    If Left(LCase(uri), 5) <> "file:" Then GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")
    iniPath = Trim(CStr(Sh.Range("AA1").Value))
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then
        If Len(ThisWorkbook.Path) > 0 Then
            iniPath = ThisWorkbook.Path & Application.PathSeparator & "excel_clipboard_launch.ini"
        End If
    End If
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then GoTo CleanFail

    allText = ReadUtf8File(iniPath)
    If Len(allText) = 0 Then GoTo CleanFail

    py = IniValue(allText, "PY")
    scriptPath = IniValue(allText, "SCRIPT")
    If Len(py) = 0 Or Len(scriptPath) = 0 Then GoTo CleanFail

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & scriptPath & Chr(34) & " " & Chr(34) & Replace(uri, Chr(34), Chr(34) & Chr(34)) & Chr(34)

    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
    Exit Sub

CleanFail:
End Sub
'''


def _excel_security_key_paths() -> list[str]:
    return [
        rf"Software\Microsoft\Office\{ver}\Excel\Security"
        for ver in ("16.0", "15.0", "14.0", "12.0")
    ]


def _open_excel_security_key(write: bool = False):
    if sys.platform != "win32" or winreg is None:
        return None
    access = winreg.KEY_READ | (winreg.KEY_SET_VALUE if write else 0)
    for subkey in _excel_security_key_paths():
        try:
            return winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey, 0, access)
        except OSError:
            continue
    return None


def _read_access_vbom() -> int | None:
    k = _open_excel_security_key(write=False)
    if k is None:
        return None
    try:
        with k:
            val, _ = winreg.QueryValueEx(k, "AccessVBOM")
            return int(val)
    except OSError:
        return None


def _write_access_vbom(value: int) -> None:
    k = _open_excel_security_key(write=True)
    if k is None:
        raise RuntimeError("Could not open Excel Security registry key.")
    with k:
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, int(value))


def write_clipboard_launch_ini(
    dest_file: Path,
    py_exe: str,
    script_path: Path,
    *,
    viewer_script: Path | None = None,
) -> Path:
    """Write PY=/SCRIPT=/VIEWER= (UTF-8) to dest_file (e.g. python_files/excel_clipboard_launch.ini)."""
    dest_file = dest_file.resolve()
    dest_file.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"PY={py_exe}\n", f"SCRIPT={script_path.resolve()}\n"]
    if viewer_script is not None:
        lines.append(f"VIEWER={viewer_script.resolve()}\n")
    dest_file.write_text("".join(lines), encoding="utf-8")
    return dest_file


def build_macro_template_file(dest: Path) -> bool:
    """
    Create dest (.xlsm) with ThisWorkbook VBA using Excel automation.
    Temporarily sets AccessVBOM=1 if needed, then restores previous value.
    """
    if sys.platform != "win32":
        return False

    try:
        import win32com.client
    except ImportError:
        print("macro_template: pywin32 not installed; cannot auto-build Excel template.")
        return False

    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.is_file():
        dest.unlink()

    prev_vbom = _read_access_vbom()
    vbom_changed = False
    try:
        if prev_vbom != 1:
            _write_access_vbom(1)
            vbom_changed = True

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Add()
            cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
            n = cm.CountOfLines
            if n > 0:
                cm.DeleteLines(1, n)
            cm.AddFromString(THISWORKBOOK_VBA)
            xl_open_xml_macro = 52
            wb.SaveAs(str(dest), FileFormat=xl_open_xml_macro)
            wb.Close(SaveChanges=False)
            wb = None
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            excel.Quit()

        return dest.is_file()
    except Exception as e:
        print(f"macro_template: Excel automation failed ({type(e).__name__}: {e}).")
        return False
    finally:
        if vbom_changed:
            try:
                if prev_vbom is None:
                    k = _open_excel_security_key(write=True)
                    if k is not None:
                        with k:
                            try:
                                winreg.DeleteValue(k, "AccessVBOM")
                            except OSError:
                                _write_access_vbom(0)
                else:
                    _write_access_vbom(prev_vbom)
            except OSError as ex:
                print(f"macro_template: could not restore AccessVBOM registry ({ex}).")


def ensure_macro_template(dest: Path) -> bool:
    """If dest is missing, try to build it."""
    dest = dest.resolve()
    if dest.is_file():
        return True
    print(f"macro_template: creating '{dest}' via Excel (first-time setup)...")
    return build_macro_template_file(dest)


if __name__ == "__main__":
    load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)
    py_files = Path(__file__).resolve().parent.parent
    target = Path(os.getenv("EXCEL_TEMPLATE_PATH", str(py_files / "orders_template.xlsm"))).expanduser().resolve()
    ok = ensure_macro_template(target)
    sys.exit(0 if ok else 1)
