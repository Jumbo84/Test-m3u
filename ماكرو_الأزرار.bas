Attribute VB_Name = "modPrint3D"
' ============================================================
'  ماكرو حاسبة الطباعة ثلاثية الأبعاد
'  ثلاثة أزرار: مشروع جديد / حفظ للسجل / فاتورة PDF
'  استورد هذا الملف عبر VBA Editor (Alt+F11) ← File ← Import File
'  أو انسخ الأكواد والصقها في وحدة نمطية جديدة.
' ============================================================
Option Explicit

' ---- حفظ المشروع الحالي إلى «سجل المشاريع» ----
Sub SaveToLog()
    Dim wsC As Worksheet, wsL As Worksheet
    Set wsC = ThisWorkbook.Worksheets("الحاسبة")
    Set wsL = ThisWorkbook.Worksheets("سجل المشاريع")

    Dim r As Long
    r = 3
    Do While Trim(CStr(wsL.Cells(r, 1).Value)) <> ""
        r = r + 1
    Loop

    Dim pieces As Double
    pieces = wsC.Range("B7").Value
    If pieces < 1 Then pieces = 1

    wsL.Cells(r, 1).Value = wsC.Range("D6").Value          ' التاريخ
    wsL.Cells(r, 2).Value = wsC.Range("B6").Value          ' رقم الفاتورة
    wsL.Cells(r, 3).Value = wsC.Range("D5").Value          ' العميل
    wsL.Cells(r, 4).Value = wsC.Range("B5").Value          ' اسم المشروع
    wsL.Cells(r, 5).Value = wsC.Range("B7").Value          ' عدد القطع
    wsL.Cells(r, 6).Value = wsC.Range("C21").Value         ' وزن الفيلامنت (إجمالي)
    wsL.Cells(r, 7).Value = wsC.Range("B24").Value         ' ساعات الطباعة
    wsL.Cells(r, 8).Value = wsC.Range("E21").Value         ' تكلفة المواد (إجمالي)
    wsL.Cells(r, 9).Value = wsC.Range("D33").Value * pieces ' الكهرباء (إجمالي)
    wsL.Cells(r, 10).Value = wsC.Range("D34").Value        ' تكلفة التصميم الكلية
    wsL.Cells(r, 11).Value = wsC.Range("D36").Value * pieces ' تغليف وشحن (إجمالي)
    wsL.Cells(r, 12).Value = wsC.Range("D38").Value * pieces ' الإهلاك (إجمالي)
    wsL.Cells(r, 13).Value = wsC.Range("D39").Value * pieces ' الربح (إجمالي)
    wsL.Cells(r, 14).Value = wsC.Range("D41").Value        ' سعر البيع الكلي

    MsgBox "تم حفظ المشروع في السجل (الصف " & r & ").", vbInformation, "حفظ"
End Sub

' ---- تجهيز مشروع جديد (يمسح المدخلات فقط، ولا يمس السجل) ----
Sub NewProject()
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets("الحاسبة")

    If MsgBox("مسح بيانات المشروع الحالي؟ (السجل لن يتأثر)", _
              vbYesNo + vbQuestion, "مشروع جديد") = vbNo Then Exit Sub

    wsC.Range("B5,D5,B6,D6").ClearContents        ' الاسم/العميل/الفاتورة/التاريخ
    wsC.Range("B11:C20").ClearContents            ' جدول الألوان
    wsC.Range("B24,D24,B25").ClearContents        ' الأوقات
    wsC.Range("B28,D28,B29").ClearContents        ' المصاريف
    wsC.Range("B7").Value = 1                     ' عدد القطع
    wsC.Range("D7").Value = 1                     ' عدد النسخ المتوقعة

    MsgBox "جاهز لإدخال مشروع جديد.", vbInformation, "مشروع جديد"
End Sub

' ---- إنشاء فاتورة PDF من ورقة «فاتورة» ----
Sub CreateInvoicePDF()
    Dim wsI As Worksheet
    Set wsI = ThisWorkbook.Worksheets("فاتورة")

    Dim invNo As String, fpath As String
    invNo = Trim(CStr(wsI.Range("C5").Value))
    If invNo = "" Then invNo = "INV"

    Dim base As String
    base = ThisWorkbook.Path
    If base = "" Then base = Application.DefaultFilePath  ' لو لم يُحفظ الملف بعد
    fpath = base & Application.PathSeparator & "Invoice_" & invNo & ".pdf"

    On Error GoTo errH
    wsI.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, _
        Quality:=xlQualityStandard, OpenAfterPublish:=True
    MsgBox "تم إنشاء الفاتورة:" & vbCrLf & fpath, vbInformation, "فاتورة PDF"
    Exit Sub
errH:
    MsgBox "تعذّر إنشاء الفاتورة: " & Err.Description, vbExclamation, "خطأ"
End Sub
