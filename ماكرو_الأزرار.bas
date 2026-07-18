Attribute VB_Name = "modPrint3D"
' ============================================================
'  Macro package for the 3D Printing Cost Calculator
'  Buttons: NewProject / SaveToLog / CreateInvoicePDF
'
'  NOTE: This code contains NO Arabic characters, so the VBA
'  editor never corrupts it (some Windows code pages turn pasted
'  Arabic into "????"). Arabic sheet names and messages are built
'  from Unicode code points via ChrW, so it works on any machine.
' ============================================================
Option Explicit

' Build an Arabic string from space-separated Unicode code points
Private Function U(ByVal codes As String) As String
    Dim p() As String, i As Long, s As String
    p = Split(Trim(codes), " ")
    For i = LBound(p) To UBound(p)
        s = s & ChrW(CLng(p(i)))
    Next i
    U = s
End Function

' Sheet-name helpers (language-neutral)
Private Function ShCalc() As Worksheet
    Set ShCalc = ThisWorkbook.Worksheets(U("1575 1604 1581 1575 1587 1576 1577"))
End Function
Private Function ShLog() As Worksheet
    Set ShLog = ThisWorkbook.Worksheets(U("1587 1580 1604 32 1575 1604 1605 1588 1575 1585 1610 1593"))
End Function
Private Function ShInv() As Worksheet
    Set ShInv = ThisWorkbook.Worksheets(U("1601 1575 1578 1608 1585 1577"))
End Function

' ---- Save current project to the log sheet ----
Sub SaveToLog()
    Dim wsC As Worksheet, wsL As Worksheet
    Set wsC = ShCalc()
    Set wsL = ShLog()

    Dim r As Long
    r = 3
    Do While Trim(CStr(wsL.Cells(r, 1).Value)) <> ""
        r = r + 1
    Loop

    Dim pieces As Double
    pieces = wsC.Range("B7").Value
    If pieces < 1 Then pieces = 1

    wsL.Cells(r, 1).Value = wsC.Range("D6").Value
    wsL.Cells(r, 2).Value = wsC.Range("B6").Value
    wsL.Cells(r, 3).Value = wsC.Range("D5").Value
    wsL.Cells(r, 4).Value = wsC.Range("B5").Value
    wsL.Cells(r, 5).Value = wsC.Range("B7").Value
    wsL.Cells(r, 6).Value = wsC.Range("C21").Value
    wsL.Cells(r, 7).Value = wsC.Range("B24").Value
    wsL.Cells(r, 8).Value = wsC.Range("E21").Value
    wsL.Cells(r, 9).Value = wsC.Range("D33").Value * pieces
    wsL.Cells(r, 10).Value = wsC.Range("D34").Value
    wsL.Cells(r, 11).Value = wsC.Range("D36").Value * pieces
    wsL.Cells(r, 12).Value = wsC.Range("D38").Value * pieces
    wsL.Cells(r, 13).Value = wsC.Range("D39").Value * pieces
    wsL.Cells(r, 14).Value = wsC.Range("D41").Value

    MsgBox U("1578 1605 32 1581 1601 1592 32 1575 1604 1605 1588 1585 1608 1593 32 1601 1610 32 1575 1604 1587 1580 1604 32 40 1575 1604 1589 1601 32") _
        & r & U("41 46"), vbInformation, U("1581 1601 1592")
End Sub

' ---- Prepare a new project (clears inputs only, keeps the log) ----
Sub NewProject()
    Dim wsC As Worksheet
    Set wsC = ShCalc()

    If MsgBox(U("1605 1587 1581 32 1576 1610 1575 1606 1575 1578 32 1575 1604 1605 1588 1585 1608 1593 32 1575 1604 1581 1575 1604 1610 1567 32 40 1575 1604 1587 1580 1604 32 1604 1606 32 1610 1578 1571 1579 1585 41"), _
              vbYesNo + vbQuestion, U("1605 1588 1585 1608 1593 32 1580 1583 1610 1583")) = vbNo Then Exit Sub

    wsC.Range("B5,D5,B6,D6").ClearContents
    wsC.Range("B11:C20").ClearContents
    wsC.Range("B24,D24,B25").ClearContents
    wsC.Range("B28,D28,B29").ClearContents
    wsC.Range("B7").Value = 1
    wsC.Range("D7").Value = 1

    MsgBox U("1580 1575 1607 1586 32 1604 1573 1583 1582 1575 1604 32 1605 1588 1585 1608 1593 32 1580 1583 1610 1583 46"), _
           vbInformation, U("1605 1588 1585 1608 1593 32 1580 1583 1610 1583")
End Sub

' ---- Export the invoice sheet to PDF ----
Sub CreateInvoicePDF()
    Dim wsI As Worksheet
    Set wsI = ShInv()

    Dim invNo As String, fpath As String, base As String
    invNo = Trim(CStr(wsI.Range("C5").Value))
    If invNo = "" Then invNo = "INV"
    base = ThisWorkbook.Path
    If base = "" Then base = Application.DefaultFilePath
    fpath = base & Application.PathSeparator & "Invoice_" & invNo & ".pdf"

    On Error GoTo errH
    wsI.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, _
        Quality:=xlQualityStandard, OpenAfterPublish:=True
    MsgBox U("1578 1605 32 1573 1606 1588 1575 1569 32 1575 1604 1601 1575 1578 1608 1585 1577 58") _
        & vbCrLf & fpath, vbInformation, U("1601 1575 1578 1608 1585 1577 32 80 68 70")
    Exit Sub
errH:
    MsgBox U("1578 1593 1584 1617 1585 32 1573 1606 1588 1575 1569 32 1575 1604 1601 1575 1578 1608 1585 1577 58 32") _
        & Err.Description, vbExclamation, U("1582 1591 1571")
End Sub
