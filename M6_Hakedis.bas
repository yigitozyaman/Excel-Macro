Attribute VB_Name = "M6_Hakedis"
Option Explicit

Sub CreateMeasurmentPage()

Dim ws, ws02, ws03 As Worksheet
Dim lr, lr02, lr03 As Long
Dim i, j As Integer
Dim sayfavar As Boolean
Dim wsName, pozAdi As String

sayfavar = False

wsName = InputBox("Ýmalat Pozunu Girininiz: (Örneðin: A)")

Set ws = ActiveWorkbook.ActiveSheet

For Each ws03 In ActiveWorkbook.Sheets
    If ws03.name = wsName Then
        ActiveWorkbook.Sheets(wsName).Copy After:=ActiveWorkbook.Worksheets(Worksheets.Count)
        sayfavar = True
        Set ws02 = ActiveWorkbook.ActiveSheet
        Exit For
    End If
Next ws03

If sayfavar = False Then
    'ActiveWorkbook.Sheets("Metraj").Copy After:=ActiveWorkbook.Worksheets(Worksheets.Count)
    Call CopyExcelPage
    Set ws02 = ActiveWorkbook.ActiveSheet
    ws02.name = (wsName)
End If

With ws
    If Application.CountA(.Cells) > 0 Then
        lr = .Cells.Find(What:=wsName, After:=.Range("A1"), LookAt:=xlWhole, _
        LookIn:=xlFormulas, searchOrder:=xlByRows, Searchdirection:=xlNext, _
        MatchCase:=False).Row
        
        lr02 = .Cells.Find(What:=wsName, After:=.Cells(lr, 1), LookAt:=xlWhole, _
        LookIn:=xlValues, searchOrder:=xlByRows, Searchdirection:=xlNext, _
        MatchCase:=False).Row
    End If
End With

If sayfavar = False Then

lr03 = ws02.Cells.Find(What:="ÝÞÝN POZU VE TANIMI", After:=ws02.Range("A1"), LookAt:=xlWhole, LookIn:=xlFormulas, searchOrder:=xlByRows, Searchdirection:=xlNext, _
        MatchCase:=False).Row
        
ws02.Cells(lr03, 2).Value = ws.Cells(lr, 2).Value
ws02.Cells(lr03, 2).WrapText = False

    For i = (lr02 - lr - 1) To 1 Step -1
    
        lr03 = ws02.Cells.Find(What:="AÇIKLAMALAR", After:=ws02.Range("A1"), LookAt:=xlWhole, LookIn:=xlFormulas, searchOrder:=xlByRows, Searchdirection:=xlNext, _
        MatchCase:=False).Row

        ws02.Cells(lr03 + 2, 1).Value = ws.Cells(lr + i, 1).Offset(0, 1).Value
        ws02.Cells(lr03 + 2, 1).WrapText = False
        
        ws.Cells(lr + i, 1).Offset(0, 3).Formula = "='" & wsName & "'!" & ws02.Range("I16").Address(False, False)
        
        If i <> 1 Then
            ws02.Rows(11 & ":" & 17).Copy
            ws02.Rows(11).Resize(5).EntireRow.Insert Shift:=xlRows
        End If

    Next i
End If

End Sub

Sub CopyExcelPage()

Application.ScreenUpdating = False

Dim sourceWB, targetWB As Workbook
Dim sheetName As String
Dim sourceFilePath As String

Set targetWB = ActiveWorkbook
sourceFilePath = "C:\Users\Kuryap 2023\Desktop\Kuryap Dosyalar\T1-Excel Taslaklar\Metraj Sayfasý Taslak"
sheetName = "Metraj"

Set sourceWB = Workbooks.Open(sourceFilePath)
sourceWB.Sheets(sheetName).Copy After:=targetWB.Sheets(targetWB.Sheets.Count)

sourceWB.Close SaveChanges:=False

Application.ScreenUpdating = True

End Sub

Sub CreateOfferPage()

Dim ws As Worksheet
Set ws = ActiveWorkbook.ActiveSheet

Dim range01 As Range
Set range01 = Application.InputBox(Prompt:="Teklif Ýstenecek Ýmalatlarý Seçiniz.", Type:=8)

Dim rangeCount As Long
rangeCount = range01.Count

Dim rowNumber() As Long

ReDim rowNumber(1 To rangeCount)

End Sub

Sub FATURALINK()

Dim fso, folder, file As Object
Dim ws As Worksheet

Set ws = ActiveSheet

Dim path01, filename As String
path01 = ActiveWorkbook.Path

Set fso = CreateObject("Scripting.FilesystemObject")
path01 = path01 & "\Hakedis Faturalar"  'KLASÖR ADI DEÐÝÞTÝÐÝ ZAMAN DEÐÝÞTÝRÝLEMESÝ GEREKEN ALAN.
Set folder = fso.GetFolder(path01)

Dim i, j As Integer
i = 0

Dim valueRange() As Variant
Dim cllName As Variant

ReDim valueRange(1 To ActiveCell.Row)

For j = 1 To ActiveCell.Row - 1
    valueRange(j) = Cells(ActiveCell.Row - j, ActiveCell.Column).Value
Next j

Dim isWrite As Boolean

For Each file In folder.Files

isWrite = False
    For Each cllName In valueRange
    
    filename = Left(file.name, Len(file.name) - 4)
    If CStr(cllName) = filename Then

        isWrite = True

    End If

    Next cllName

    If LCase(Right(file.name, 3)) = "pdf" And isWrite = False Then
        
        filename = Left(file.name, Len(file.name) - 4)
        'ActiveCell.Offset(i, 0).EntireRow.Insert
        ActiveCell.Offset(i, 0).Value = filename
        ws.Hyperlinks.Add Anchor:=ActiveCell.Offset(i, 0), Address:=file.Path
        i = i + 1
        
    End If

Next file

Set fso = Nothing
Set folder = Nothing
Set file = Nothing
Set ws = Nothing
End Sub


