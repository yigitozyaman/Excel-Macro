Attribute VB_Name = "M2_OrganiseSheet"
Option Explicit

Sub MergeExcelFiles()
  Dim fnameList, fnameCurFile As Variant
  Dim countFiles, countSheets As Integer
  Dim wksCurSheet As Worksheet
  Dim wbkCurBook, wbkSrcBook As Workbook

  fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)

  If (vbBoolean <> VarType(fnameList)) Then

    If (UBound(fnameList) > 0) Then
      countFiles = 0
      countSheets = 0

      Application.ScreenUpdating = False
      Application.Calculation = xlCalculationManual

      Set wbkCurBook = ActiveWorkbook

      For Each fnameCurFile In fnameList
          countFiles = countFiles + 1

          Set wbkSrcBook = Workbooks.Open(filename:=fnameCurFile)

          For Each wksCurSheet In wbkSrcBook.Sheets
                Application.DisplayAlerts = False
              countSheets = countSheets + 1
              wksCurSheet.Copy After:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
          Next

          wbkSrcBook.Close SaveChanges:=False

      Next

      Application.ScreenUpdating = True
      Application.Calculation = xlCalculationAutomatic

      MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
    End If

  Else
      MsgBox "No files selected", Title:="Merge Excel files"
  End If
End Sub

Sub Sayfa_Ad_Sirala()
'Sayfa isimlerini alfabetik olarak sýralar.

Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult

   iAnswer = MsgBox("Sayfa isimlerini A dan Z ye mi sýralansýn?" & Chr(10) _
   & "Hayýr'a týklarsanýz sayfa isimleri Z'den A'ya sýralanýr", _
   vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
 For i = 1 To Sheets.Count
 For j = 1 To Sheets.Count - 1

 If iAnswer = vbYes Then
 If UCase$(Sheets(j).name) > UCase$(Sheets(j + 1).name) Then
   Sheets(j).Move After:=Sheets(j + 1)
 End If

 ElseIf iAnswer = vbNo Then
 If UCase$(Sheets(j).name) < UCase$(Sheets(j + 1).name) Then
 Sheets(j).Move After:=Sheets(j + 1)
 End If
 End If
 Next j
 Next i
End Sub

Sub MoveSheetEnd()
'Aktif sayfayý sona taþýr. ctrl+shift+E kýsayolu ile kullanýlýr.

Dim ws As Worksheet
Dim Indx As Integer

Set ws = ActiveSheet
Indx = ws.Index

ws.Move After:=Sheets(Sheets.Count)

Application.OnKey "^+{E}", "M3_OrganiseSheet.MoveSheetEnd"

Worksheets(Indx).Activate

End Sub

Sub ChangeSheetName()

Dim ws As Worksheet
Dim i As Integer
i = 100

Application.DisplayAlerts = False
For Each ws In ActiveWindow.SelectedSheets

    ws.name = ws.Range("M2").Value '& " " & i
    'i = i + 1
    
Next ws
Application.DisplayAlerts = True

End Sub

Sub DeleteSheetByName()

Dim ws As Worksheet
Application.DisplayAlerts = False

For Each ws In ActiveWindow.SelectedSheets

    If InStr(1, ws.name, "Ýlave", vbTextCompare) = 0 Then
    
        ws.Delete
        
    End If
        
Next ws

Application.DisplayAlerts = True

End Sub

Sub SelectSheetByColour()

Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
'Debug.Print sh.Tab.ColorIndex

    If sh.Tab.ColorIndex = 6 Then
            'sh.Name = Left(sh.Name, Len(sh.Name) - 1)
            'sh.Name = Worksheets(sh.Index + 1).Name + "-" + sh.Name
            sh.Select False
    End If
    
Next sh

End Sub

Sub ChangeSheetName2()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    If Left(ws.name, 1) = " " Then
    
       ws.name = Right(ws.name, Len(ws.name) - 1)
    End If
Next ws

End Sub

Sub DocSheetsClear()

Dim fnameList, fnameCurFile As Variant
Dim openedWb As Workbook
Dim activeSheets As Worksheet
Dim name1 As Variant

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
Title:="Choose Excel files to delete its sheets", MultiSelect:=True)

For Each fnameCurFile In fnameList

    Set openedWb = Workbooks.Open(filename:=fnameCurFile)
    
        For Each activeSheets In openedWb.Sheets
    
            If InStr(1, activeSheets.name, "Ýlave", vbTextCompare) Then
                'MsgBox ("Bulundu")
                
            Else
                activeSheets.Delete
            End If
        Next
        
        openedWb.Close SaveChanges:=True
           
Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub

Sub SetPageNo()

Dim sheets1 As Worksheet
Dim pageNo As String
Dim targetCell, findArea As Range
Dim number1, number2 As Integer
number2 = 2

pageNo = "Sayfa No"

For Each sheets1 In ActiveWorkbook.Worksheets

If sheets1.Visible = True Then
Set targetCell = sheets1.Cells.Find(What:=pageNo)

If Not targetCell Is Nothing Then
targetCell.Offset(0, 1).Value = number2
number2 = number2 + 1
Set findArea = Range(targetCell.Offset(1, 0), sheets1.Cells(Rows.Count, targetCell.Column))

10:
Set targetCell = findArea.Find(What:=pageNo)
If Not targetCell Is Nothing Then
targetCell.Offset(0, 1).Value = number2
number2 = number2 + 1
Set findArea = Range(targetCell.Offset(1, 0), sheets1.Cells(Rows.Count, targetCell.Column))
GoTo 10:
End If
End If
End If
Next sheets1


End Sub
