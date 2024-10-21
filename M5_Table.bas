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

baslik = Array("AÇIKLAMA", "POZ", "ADET", "ÇAP", "UZUNLUK (cm)", "BENZER", "AÐIRLIK (kg)")

Set range01 = ActiveCell.Resize(2, UBound(baslik) + 1)

Set donatiTablo = ActiveSheet.ListObjects.Add(xlSrcRange, range01, xlYes)
donatiTablo.name = "DonatýMetraj"

Dim i As Integer
For i = LBound(baslik) To UBound(baslik)
    range01.Cells(0, i + 1).Value = baslik(i)
Next i

donatiTablo.ListColumns("AÐIRLIK (kg)").DataBodyRange.Formula = "=+((22/7)*([@ÇAP]/1000)^2/4*7850)*[@ADET]*[@[UZUNLUK (cm)]]/100*[@BENZER]"
donatiTablo.ListColumns("AÐIRLIK (kg)").DataBodyRange.NumberFormat = "#,##0.00 ""kg"""
donatiTablo.ListColumns("ÇAP").DataBodyRange.NumberFormat = " ""Ø"" 0"

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

headers01 = Array("POZ", "POZ AÇIKLAMASI", "POZ BÝRÝMÝ", "YAPILAN ÝÞ", "ALT KALEMLER", "ÝÞ TÝPÝ", "BÝRÝM", "MÝKTAR", "BÝRÝM FÝYAT", "TUTAR")

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

With bfaTable.ListColumns("ÝÞ TÝPÝ").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , Join(workType, ",")
    .InCellDropdown = True
    .ErrorMessage = "Lütfen uygun iþ tipini seçiniz. (Alt + alt ok tuþunu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("BÝRÝM").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , Join(unit, ",")
    .InCellDropdown = True
    .ErrorMessage = "Lütfen uygun birimi seçiniz. (Alt + alt ok tuþunu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("POZ BÝRÝMÝ").DataBodyRange.Validation
    .Delete
    .Add xlValidateList, xlValidAlertStop, , UCase(Join(unit, ","))
    .InCellDropdown = True
    .ErrorMessage = "Lütfen uygun birimi seçiniz. (Alt + alt ok tuþunu kullanabilirsiniz.)"
End With

With bfaTable.ListColumns("TUTAR").DataBodyRange
    .Formula = "=+[@MÝKTAR]*[@BÝRÝM FÝYAT]"
    .NumberFormat = "#,##0.00"
End With

End Sub



