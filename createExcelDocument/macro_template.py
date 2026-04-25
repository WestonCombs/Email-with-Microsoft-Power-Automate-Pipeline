"""
Build orders_template.xlsm on Windows using Excel COM (one-time automation).

Requires: Excel installed, pywin32, and temporary AccessVBOM registry access
("Trust access to the VBA project object model") — enabled only for the build.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from shared.settings_store import apply_runtime_settings_from_json

if sys.platform == "win32":
    import winreg
else:
    winreg = None  # type: ignore[assignment]

CLIPBOARD_INI_NAME = "excel_clipboard_launch.ini"

# Workbook_SheetFollowHyperlink: Open File Location uses # in-sheet links; file URI in column 29 (AC).
# Tracking URLs: 30…44. Tracking numbers: 45…59. Link-cross-check flags: 60…74 (1 = also found on tracking URL).
# Reads UTF-8 ini (PY=, SCRIPT=, VIEWER=, TRACKING_NUMBERS_VIEWER=, TRACKING_STATUS_VIEWER=) from AA1 / excel_clipboard_launch.ini.
THISWORKBOOK_VBA = r'''Option Explicit

Private Const COL_TRACK_URI_START As Long = 30
Private Const COL_TRACK_URI_END As Long = 44
Private Const COL_TRACK_NUM_START As Long = 45
Private Const COL_TRACK_NUM_END As Long = 59
Private Const COL_TRACK_CONF_START As Long = 60
Private Const COL_TRACK_CONF_END As Long = 74
Private Const DEFAULT_HEADER_ROW As Long = 2

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

Private Function CollectTrackingNumbersForRow(ByVal Sh As Object, ByVal rowNum As Long) As String
    Dim c As Long
    Dim v As Variant
    Dim s As String
    Dim body As String
    body = ""
    For c = COL_TRACK_NUM_START To COL_TRACK_NUM_END
        v = Sh.Cells(rowNum, c).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            s = Trim(CStr(v))
            If Len(s) > 0 Then
                body = body & s & vbLf
            End If
        End If
    Next c
    CollectTrackingNumbersForRow = body
End Function

Private Function CollectTrackingNumbersAndConfirmForRow(ByVal Sh As Object, ByVal rowNum As Long) As String
    Dim c As Long
    Dim slot As Long
    Dim v As Variant
    Dim fv As Variant
    Dim s As String
    Dim flag As String
    Dim body As String
    body = ""
    For c = COL_TRACK_NUM_START To COL_TRACK_NUM_END
        slot = c - COL_TRACK_NUM_START
        v = Sh.Cells(rowNum, c).Value
        If IsError(v) Or IsEmpty(v) Then GoTo NextTC
        s = Trim(CStr(v))
        If Len(s) = 0 Then GoTo NextTC
        flag = "0"
        fv = Sh.Cells(rowNum, COL_TRACK_CONF_START + slot).Value
        If Not IsError(fv) And Not IsEmpty(fv) Then
            If Trim(CStr(fv)) = "1" Then flag = "1"
        End If
        body = body & s & Chr(9) & flag & vbLf
NextTC:
    Next c
    CollectTrackingNumbersAndConfirmForRow = body
End Function

Private Function HeaderRow(ByVal Sh As Object) As Long
    If Trim(CStr(Sh.Cells(DEFAULT_HEADER_ROW, 1).Value)) = "Category" Then
        HeaderRow = DEFAULT_HEADER_ROW
    Else
        HeaderRow = 1
    End If
End Function

Private Function HeaderColumn(ByVal Sh As Object, ByVal want As String) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim h As String
    Dim rowNum As Long
    rowNum = HeaderRow(Sh)
    On Error Resume Next
    lastCol = Sh.Cells(rowNum, Sh.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    For c = 1 To lastCol
        h = Trim(CStr(Sh.Cells(rowNum, c).Value))
        If StrComp(h, want, vbTextCompare) = 0 Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
    HeaderColumn = 0
End Function

Private Function HeaderColumnAny(ByVal Sh As Object, ParamArray wants() As Variant) As Long
    Dim i As Long
    Dim c As Long
    For i = LBound(wants) To UBound(wants)
        c = HeaderColumn(Sh, CStr(wants(i)))
        If c > 0 Then
            HeaderColumnAny = c
            Exit Function
        End If
    Next i
    HeaderColumnAny = 0
End Function

Private Function TrimmedCellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TrimmedCellText = ""
        Exit Function
    End If
    TrimmedCellText = Trim(CStr(v))
End Function

Private Function ContextValueForHeaders(ByVal Sh As Object, ByVal rowNum As Long, ParamArray headers() As Variant) As String
    Dim c As Long
    Dim i As Long
    Dim orderCol As Long
    Dim targetOrder As String
    Dim r As Long
    Dim lastData As Long
    Dim s As String

    c = 0
    For i = LBound(headers) To UBound(headers)
        c = HeaderColumn(Sh, CStr(headers(i)))
        If c > 0 Then Exit For
    Next i
    If c = 0 Then
        ContextValueForHeaders = ""
        Exit Function
    End If

    s = TrimmedCellText(Sh.Cells(rowNum, c).Value)
    If Len(s) > 0 Then
        ContextValueForHeaders = s
        Exit Function
    End If

    orderCol = HeaderColumn(Sh, "Order Number")
    If orderCol = 0 Then
        ContextValueForHeaders = ""
        Exit Function
    End If

    targetOrder = TrimmedCellText(Sh.Cells(rowNum, orderCol).Value)
    If Len(targetOrder) = 0 Then
        ContextValueForHeaders = ""
        Exit Function
    End If

    r = rowNum - 1
    Do While r >= 2
        If TrimmedCellText(Sh.Cells(r, orderCol).Value) <> targetOrder Then Exit Do
        s = TrimmedCellText(Sh.Cells(r, c).Value)
        If Len(s) > 0 Then
            ContextValueForHeaders = s
            Exit Function
        End If
        r = r - 1
    Loop

    On Error Resume Next
    lastData = Sh.Cells(Sh.Rows.Count, orderCol).End(xlUp).Row
    On Error GoTo 0
    If lastData < rowNum Then lastData = rowNum

    r = rowNum + 1
    Do While r <= lastData
        If TrimmedCellText(Sh.Cells(r, orderCol).Value) <> targetOrder Then Exit Do
        s = TrimmedCellText(Sh.Cells(r, c).Value)
        If Len(s) > 0 Then
            ContextValueForHeaders = s
            Exit Function
        End If
        r = r + 1
    Loop

    ContextValueForHeaders = ""
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
    Dim body As String
    Dim companyText As String
    body = ""
    companyText = ContextValueForHeaders(Sh, rowNum, "Company", "Retailer", "Store", "Merchant", "Vendor")
    If Len(companyText) > 0 Then body = body & CtxLine("company", companyText)
    body = body & CtxLine("order_number", ContextValueForHeaders(Sh, rowNum, "Order Number"))
    body = body & CtxLine("category", ContextValueForHeaders(Sh, rowNum, "Category"))
    body = body & CtxLine("purchase_datetime", ContextValueForHeaders(Sh, rowNum, "Purchase Date"))
    body = body & CtxLine("email", ContextValueForHeaders(Sh, rowNum, "Email"))
    body = body & CtxLine("workbook_path", ThisWorkbook.FullName)
    body = body & CtxLine("sheet_name", Sh.Name)
    body = body & CtxLine("row_number", rowNum)
    Dim tns As String
    Dim flat As String
    tns = CollectTrackingNumbersForRow(Sh, rowNum)
    If Len(Trim(tns)) > 0 Then
        flat = Replace(Replace(Trim(tns), vbCr, ""), vbLf, ", ")
        body = body & CtxLine("tracking_numbers", flat)
    End If
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

Private Sub LaunchTrackingNumbersViewerForRow(ByVal Sh As Object, ByVal rowNum As Long)
    Dim body As String
    Dim tempPath As String
    Dim fso As Object
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim viewer As String
    Dim cmd As String
    Dim shell As Object

    body = CollectTrackingNumbersAndConfirmForRow(Sh, rowNum)
    If Len(Trim(body)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2) & "\email_sorter_trknums_r" & rowNum & "_t" & CLng(Timer * 10000) & ".txt"
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
    viewer = IniValue(allText, "TRACKING_NUMBERS_VIEWER")
    If Len(py) = 0 Or Len(viewer) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(viewer) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & viewer & Chr(34) & " " & Chr(34) & tempPath & Chr(34) & " web"
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
End Sub

Private Function CollectTrackingNumbersOrderBlockForRow(ByVal Sh As Object, ByVal rowNum As Long) As String
    Dim orderCol As Long
    Dim v As Variant
    Dim targetOrder As String
    Dim vv As Variant
    Dim startR As Long
    Dim endR As Long
    Dim lastData As Long
    Dim r2 As Long
    Dim body As String
    orderCol = HeaderColumn(Sh, "Order Number")
    If orderCol = 0 Then
        CollectTrackingNumbersOrderBlockForRow = ""
        Exit Function
    End If
    v = Sh.Cells(rowNum, orderCol).Value
    If IsError(v) Or IsEmpty(v) Then
        targetOrder = ""
    Else
        targetOrder = Trim(CStr(v))
    End If
    startR = rowNum
    Do While startR > 2
        vv = Sh.Cells(startR - 1, orderCol).Value
        If IsError(vv) Or IsEmpty(vv) Then Exit Do
        If Trim(CStr(vv)) <> targetOrder Then Exit Do
        startR = startR - 1
    Loop
    On Error Resume Next
    lastData = Sh.Cells(Sh.Rows.Count, orderCol).End(xlUp).Row
    On Error GoTo 0
    If lastData < 2 Then lastData = rowNum
    endR = rowNum
    Do While endR < lastData
        vv = Sh.Cells(endR + 1, orderCol).Value
        If IsError(vv) Or IsEmpty(vv) Then Exit Do
        If Trim(CStr(vv)) <> targetOrder Then Exit Do
        endR = endR + 1
    Loop
    body = ""
    For r2 = startR To endR
        body = body & CollectTrackingNumbersAndConfirmForRow(Sh, r2)
    Next r2
    CollectTrackingNumbersOrderBlockForRow = body
End Function

Private Sub LaunchTrackingNumbersOrderViewerForRow(ByVal Sh As Object, ByVal rowNum As Long)
    Dim body As String
    Dim tempPath As String
    Dim fso As Object
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim viewer As String
    Dim cmd As String
    Dim shell As Object

    body = CollectTrackingNumbersOrderBlockForRow(Sh, rowNum)
    If Len(Trim(body)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2) & "\email_sorter_trkord_r" & rowNum & "_t" & CLng(Timer * 10000) & ".txt"
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
    viewer = IniValue(allText, "TRACKING_NUMBERS_VIEWER")
    If Len(py) = 0 Or Len(viewer) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(viewer) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & viewer & Chr(34) & " " & Chr(34) & tempPath & Chr(34) & " order"
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
End Sub

Private Sub LaunchTrackingStatusViewerForRow(ByVal Sh As Object, ByVal rowNum As Long)
    Dim body As String
    Dim tempPath As String
    Dim fso As Object
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim viewer As String
    Dim cmd As String
    Dim shell As Object

    ' Same tracking set as "View Tracking Numbers (All For Order)" (order block, all rows).
    body = CollectTrackingNumbersOrderBlockForRow(Sh, rowNum)
    If Len(Trim(body)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.GetSpecialFolder(2) & "\email_sorter_trkstat_r" & rowNum & "_t" & CLng(Timer * 10000) & ".txt"
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
    viewer = IniValue(allText, "TRACKING_STATUS_VIEWER")
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

Private Sub Workbook_Open()
    On Error Resume Next
    Call LaunchPodWorkflowWatcher
End Sub

Private Sub LaunchGiftInvoiceLinkWorkflow(ByVal Sh As Object, ByVal rowNum As Long)
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim linkScript As String
    Dim fso As Object
    Dim cmd As String
    Dim shell As Object

    On Error GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")
    iniPath = Trim(CStr(Sh.Range("AA1").Value))
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then
        If Len(ThisWorkbook.Path) > 0 Then
            iniPath = ThisWorkbook.Path & Application.PathSeparator & "excel_clipboard_launch.ini"
        End If
    End If
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then Exit Sub

    allText = ReadUtf8File(iniPath)
    If Len(allText) = 0 Then Exit Sub

    py = IniValue(allText, "PY")
    linkScript = IniValue(allText, "GIFTCARD_LINK")
    If Len(py) = 0 Or Len(linkScript) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(linkScript) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & linkScript & Chr(34) & " " & Chr(34) & ThisWorkbook.FullName & Chr(34) & " " & CStr(rowNum)
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 1, False
    Exit Sub

CleanFail:
End Sub

Private Sub LaunchPodWorkflowWatcher()
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim podScript As String
    Dim fso As Object
    Dim cmd As String
    Dim shell As Object

    On Error GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")
    iniPath = Trim(CStr(ThisWorkbook.Worksheets("Orders").Range("AA1").Value))
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then
        If Len(ThisWorkbook.Path) > 0 Then
            iniPath = ThisWorkbook.Path & Application.PathSeparator & "excel_clipboard_launch.ini"
        End If
    End If
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then Exit Sub

    allText = ReadUtf8File(iniPath)
    If Len(allText) = 0 Then Exit Sub

    py = IniValue(allText, "PY")
    podScript = IniValue(allText, "POD_WORKFLOW")
    If Len(py) = 0 Or Len(podScript) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(podScript) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & podScript & Chr(34) & " watch " & Chr(34) & ThisWorkbook.FullName & Chr(34)
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
    Exit Sub

CleanFail:
End Sub

Private Sub LaunchRemainingPodViewer(ByVal rowNum As Long)
    Dim iniPath As String
    Dim allText As String
    Dim py As String
    Dim podScript As String
    Dim fso As Object
    Dim cmd As String
    Dim shell As Object

    On Error GoTo CleanFail

    Set fso = CreateObject("Scripting.FileSystemObject")
    iniPath = Trim(CStr(ThisWorkbook.Worksheets("Orders").Range("AA1").Value))
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then
        If Len(ThisWorkbook.Path) > 0 Then
            iniPath = ThisWorkbook.Path & Application.PathSeparator & "excel_clipboard_launch.ini"
        End If
    End If
    If Len(iniPath) = 0 Or Not fso.FileExists(iniPath) Then Exit Sub

    allText = ReadUtf8File(iniPath)
    If Len(allText) = 0 Then Exit Sub

    py = IniValue(allText, "PY")
    podScript = IniValue(allText, "POD_WORKFLOW")
    If Len(py) = 0 Or Len(podScript) = 0 Then Exit Sub
    If Not fso.FileExists(py) Then Exit Sub
    If Not fso.FileExists(podScript) Then Exit Sub

    cmd = Chr(34) & py & Chr(34) & " " & Chr(34) & podScript & Chr(34) & " launch-remaining " & Chr(34) & ThisWorkbook.FullName & Chr(34) & " " & CStr(rowNum)
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 0, False
    Exit Sub

CleanFail:
End Sub

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Const COL_FILE_URI As Long = 29
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

    header = Trim(CStr(Sh.Cells(HeaderRow(Sh), Target.Range.Column).Value))

    If StrComp(header, "View Tracking Links", vbTextCompare) = 0 _
        Or StrComp(header, "View tracking links", vbTextCompare) = 0 _
        Or StrComp(header, "View Link List", vbTextCompare) = 0 Then
        Call LaunchTrackingLinkViewerForRow(Sh, Target.Range.Row)
        Exit Sub
    End If

    If StrComp(header, "View Tracking Numbers", vbTextCompare) = 0 _
        Or StrComp(header, "View tracking numbers (web)", vbTextCompare) = 0 Then
        Call LaunchTrackingNumbersViewerForRow(Sh, Target.Range.Row)
        Exit Sub
    End If

    If StrComp(header, "View Tracking Numbers (All For Order)", vbTextCompare) = 0 Then
        Call LaunchTrackingNumbersOrderViewerForRow(Sh, Target.Range.Row)
        Exit Sub
    End If

    If StrComp(header, "Shipping Status", vbTextCompare) = 0 _
        Or StrComp(header, "Shipping summary", vbTextCompare) = 0 _
        Or StrComp(header, "View shipping status", vbTextCompare) = 0 _
        Or StrComp(header, "View Shipping Status", vbTextCompare) = 0 Then
        Dim catCol As Long
        Dim catValue As String
        catCol = HeaderColumn(Sh, "Category")
        catValue = ""
        If catCol > 0 Then
            catValue = TrimmedCellText(Sh.Cells(Target.Range.Row, catCol).Value)
        End If
        If StrComp(catValue, "Automation Hub", vbTextCompare) = 0 Then
            Call LaunchRemainingPodViewer(Target.Range.Row)
        Else
            Call LaunchTrackingStatusViewerForRow(Sh, Target.Range.Row)
        End If
        Exit Sub
    End If

    If StrComp(header, "Invoice Link", vbTextCompare) = 0 _
        Or StrComp(header, "Invoice link", vbTextCompare) = 0 Then
        Call LaunchGiftInvoiceLinkWorkflow(Sh, Target.Range.Row)
        Exit Sub
    End If

    If StrComp(header, "Open File Location", vbTextCompare) <> 0 _
        And StrComp(header, "Copy Path", vbTextCompare) <> 0 Then Exit Sub

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
    giftcard_link_script: Path | None = None,
    tracking_numbers_viewer_script: Path | None = None,
    tracking_status_viewer_script: Path | None = None,
    pod_workflow_script: Path | None = None,
) -> Path:
    """Write the Excel launcher INI consumed by VBA helpers (UTF-8)."""
    dest_file = dest_file.resolve()
    dest_file.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"PY={py_exe}\n", f"SCRIPT={script_path.resolve()}\n"]
    if viewer_script is not None:
        lines.append(f"VIEWER={viewer_script.resolve()}\n")
    if giftcard_link_script is not None:
        lines.append(f"GIFTCARD_LINK={giftcard_link_script.resolve()}\n")
    if tracking_numbers_viewer_script is not None:
        lines.append(f"TRACKING_NUMBERS_VIEWER={tracking_numbers_viewer_script.resolve()}\n")
    if tracking_status_viewer_script is not None:
        lines.append(f"TRACKING_STATUS_VIEWER={tracking_status_viewer_script.resolve()}\n")
    if pod_workflow_script is not None:
        lines.append(f"POD_WORKFLOW={pod_workflow_script.resolve()}\n")
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
        import pythoncom
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

        # Excel COM requires per-thread CoInitialize (e.g. worker threads, some hosts).
        co_inited = False
        try:
            pythoncom.CoInitialize()
            co_inited = True
            excel = None
            try:
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
            finally:
                if excel is not None:
                    try:
                        excel.Quit()
                    except Exception:
                        pass
        finally:
            if co_inited:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

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


def refresh_macro_template(dest: Path) -> bool:
    """
    Rebuild the macro template in place from the current VBA source.

    The refresh is performed via a temporary file first so an existing working template
    is not lost if Excel COM fails during regeneration.
    """
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    temp_dest = dest.with_name(dest.stem + ".__codex_refresh__.xlsm")
    try:
        if temp_dest.exists():
            temp_dest.unlink()
    except OSError:
        pass

    if not build_macro_template_file(temp_dest):
        try:
            if temp_dest.exists():
                temp_dest.unlink()
        except OSError:
            pass
        return False

    backup = dest.with_name(dest.stem + ".__codex_backup__.xlsm")
    try:
        if backup.exists():
            backup.unlink()
    except OSError:
        pass

    try:
        if dest.exists():
            dest.replace(backup)
        temp_dest.replace(dest)
        try:
            if backup.exists():
                backup.unlink()
        except OSError:
            pass
        return True
    except OSError as exc:
        print(f"macro_template: could not replace template in place ({exc}).")
        try:
            if temp_dest.exists():
                temp_dest.unlink()
        except OSError:
            pass
        try:
            if backup.exists() and not dest.exists():
                backup.replace(dest)
        except OSError:
            pass
        return False


if __name__ == "__main__":
    _PYTHON_FILES_MAIN = Path(__file__).resolve().parent.parent
    if str(_PYTHON_FILES_MAIN) not in sys.path:
        sys.path.insert(0, str(_PYTHON_FILES_MAIN))

    apply_runtime_settings_from_json()
    from shared.project_paths import ensure_base_dir_in_environ

    default_tpl = ensure_base_dir_in_environ() / "email_contents" / "orders_template.xlsm"
    target = Path(os.getenv("EXCEL_TEMPLATE_PATH", str(default_tpl))).expanduser().resolve()
    ok = ensure_macro_template(target)
    sys.exit(0 if ok else 1)
