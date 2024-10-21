Attribute VB_Name = "M5_Table"
Option Explicit

Sub ShortCutsTable()
Application.OnKey "^+{P}", "M5_Table.OrganiseTableRow"
End Sub

Sub OrganiseTableRow()

Dim table01 As ListObject
Dim range01 As Range
Dim rangeValue As String

Dim tableName As String

tableName = InputBox("Tablo ismini giriniz:", , , 8)
Set table01 = ActiveSheet.ListObjects(tableName)
Set range01 = table01.ListColumns("POZ").DataBodyRange

rangeValue = range01.Cells(1, 1).Value

table01.Range.Interior.ColorIndex = xlNone

Dim i As Integer
Dim j As Integer

For i = 2 To table01.ListRows.Count
    
    If range01.Cells(i, 1).Value = rangeValue Then
    
        range01.Cells(i, 1).Interior.ColorIndex = range01.Cells(i - 1, 1).Interior.ColorIndex
        
        For j = 1 To table01.ListColumns.Count
                range01.Cells(i, j).Interior.ColorIndex = range01.Cells(i - 1, j).Interior.ColorIndex
        Next j
    
    Else
        rangeValue = range01.Cells(i, 1).Value
        
        Select Case range01.Cells(i - 1, 1).Interior.ColorIndex
        
            Case xlNone:
                For j = 1 To table01.ListColumns.Count
                    range01.Cells(i, j).Interior.ColorIndex = 24
                Next j
            
            Case 24:
                For j = 1 To table01.ListColumns.Count
                    range01.Cells(i, j).Interior.ColorIndex = xlNone
                Next j
        End Select
    End If
Next i
    
End Sub

Sub DonatiTablosu()

Dim donatiTablo As ListObject
Dim range01 As Range
Dim baslik As Variant

baslik = Array("A�IKLAMA", "POZ", "ADET", "�AP", "UZUNLUK (cm)", "BENZER", "A�IRLIK (kg)")

Set range01 = ActiveCell.Resize(2, UBound(baslik) + 1)

Set donatiTablo = ActiveSheet.ListObjects.Add(xlSrcRange, range01, xlYes)
donatiTablo.name = "Donat�Metraj"

Dim i As Integer
For i = LBound(baslik) To UBound(baslik)
    range01.Cells(0, i + 1).Value = baslik(i)
Next i

donatiTablo.ListColumns("A�IRLIK (kg)").DataBodyRange.Formula = "=+((22/7)*([@�AP]/1000)^2/4*7850)*[@ADET]*[@[UZUNLUK (cm)]]/100*[@BENZER]"
donatiTablo.ListColumns("A�IRLIK (kg)").DataBodyRange.NumberFormat = "#,##0.00 ""kg"""
donatiTablo.ListColumns("�AP").DataBodyRange.NumberFormat = " ""�"" 0"

With donatiTablo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
End With

Dim Cll As Range
For Each Cll In donatiTablo.HeaderRowRange
    Cll.EntireColumn.AutoFit
Next Cll

End Sub

Sub bfaTable()

Dim bfaTable As ListObject
Dim range01 As Range
Dim headers01 As Variant

headers01 = Array("POZ", "POZ A�IKLAMASI", "POZ B�R�M�", "YAPILAN ��", "ALT KALEMLER", "�� T�P�", "B�R�M", "M�KTAR", "B�R�M F�YAT", "TUTAR")

Set range01 = ActiveCell.Resize(20, UBound(headers01) + 1)

Set bfaTable = ActiveSheet.ListObjects.Add(xlSrcRange, range01, xlYes)
bfaTable.name = "BFA"

Dim i As Integer
For i = LBound(headers01) To UBound(headers01)
    range01.Cells(0, i + 1).Value = headers01(i)
Next i

With bfaTable.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
End With

Dim Cll As Range
For Each Cll In bfaTable.HeaderRowRange
    Cll.EntireColumn.AutoFit
Next Cll

Dim workType, unit As Variant
workType = Array("MALZEME", "ISCILIK", "NAKLIYE", "SARFIYAT")
unit = Array("adet", "mt", "m2", "m3", "ton", "kg", "set", "yuzde", "saat", "gun", "ay")

With bfaTable.ListColumns("�� T�P�").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , Join(workType, ",")
    .InCellDropdown = True
    .ErrorMessage = "L�tfen uygun i� tipini se�iniz. (Alt + alt ok tu�unu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("B�R�M").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , Join(unit, ",")
    .InCellDropdown = True
    .ErrorMessage = "L�tfen uygun birimi se�iniz. (Alt + alt ok tu�unu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("POZ B�R�M�").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , UCase(Join(unit, ","))
    .InCellDropdown = True
    .ErrorMessage = "L�tfen uygun birimi se�iniz. (Alt + alt ok tu�unu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("TUTAR").DataBodyRange
    .Formula = "=+[@M�KTAR]*[@B�R�M F�YAT]"
    .NumberFormat = "#,##0.00"
End With

End Sub



