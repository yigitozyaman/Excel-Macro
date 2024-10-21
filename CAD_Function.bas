Attribute VB_Name = "M4_CADfunction"
Option Explicit

Sub ShortCutsCAD()
    Application.OnKey "^+{Ç}", "M4_CADFunction.GetPolylineLength"
    Application.OnKey "^+{Ö}", "M4_CADFunction.GetPolylineArea"
End Sub

Sub GetPolylineLength()
    Dim acadApp As AcadApplication
    Set acadApp = GetObject(, "AutoCAD.Application")
    'acadapp.ActiveDocument
    Dim activeDoc As AcadDocument
    Set activeDoc = acadApp.ActiveDocument
    
    '----------------
    Dim totalsum As VbMsgBoxResult
    totalsum = MsgBox("Kümülatifi görmek istiyor musunuz?", vbYesNo)
    '----------------

    Dim selObj As Object
    
    On Error Resume Next
    activeDoc.SelectionSets.Item("mySelectionSets").Delete
    Set selObj = activeDoc.SelectionSets.Add("mySelectionSets")
    selObj.SelectOnScreen
    On Error GoTo 0
 
 Dim obj As Object
 
 For Each obj In selObj
    If obj.ObjectName = "AcDbPolyline" Then
        Dim poly As AcadLWPolyline
        Set poly = obj
        Dim length, tLen As Double
        length = poly.length
        'MsgBox "uzunluk: " & length
        Dim newColor As Object
        Set newColor = poly.TrueColor
        newColor.ColorIndex = 3
        poly.TrueColor = newColor
        
        If totalsum = vbNo Then
            Selection.Value = Round(length / 100, 2)  'cm uzunluðunu m olarak verir.
            Selection.Offset(1, 0).Select
        End If
        
        tLen = length + tLen
    End If
    Next
    'MsgBox "uzunluk: " & tLen
    
If totalsum = vbYes Then
    Selection.Value = Round(tLen / 100, 2) 'cm uzunluðunu m olarak verir.
End If

    selObj.Delete
   
End Sub

Sub GetPolylineArea()
    Dim acadApp As AcadApplication
    Set acadApp = GetObject(, "AutoCAD.Application")
    'acadapp.ActiveDocument
    Dim activeDoc As AcadDocument
    Set activeDoc = acadApp.ActiveDocument
    
    '----------------
    Dim totalsum As VbMsgBoxResult
    totalsum = MsgBox("Kümülatifi görmek istiyor musunuz?", vbYesNo)
    '----------------
    
    Dim selObj As Object
    
    On Error Resume Next
    activeDoc.SelectionSets.Item("mySelectionSets").Delete
    Set selObj = activeDoc.SelectionSets.Add("mySelectionSets")
    selObj.SelectOnScreen
    On Error GoTo 0
    
 Dim obj As Object
 
 For Each obj In selObj
    If obj.ObjectName = "AcDbPolyline" Then
        Dim poly As AcadLWPolyline
        Set poly = obj
        Dim area, tArea As Double
        area = poly.area
        'MsgBox "Alan: " & area
        Dim newColor As Object
        Set newColor = poly.TrueColor
        newColor.ColorIndex = 1
        poly.TrueColor = newColor
        
         If totalsum = vbNo Then
            Selection.Value = Round(area / 10000, 2) 'cm2 uzunluðunu m2 olarak verir.
            Selection.Offset(1, 0).Select
        End If
        
        tArea = area + tArea
    End If
    Next
    
    'MsgBox "Alan: " & tArea
If totalsum = vbYes Then
    Selection.Value = Round(tArea / 10000, 2) 'cm2 uzunluðunu m2 olarak verir.
End If

    selObj.Delete
   
End Sub

Sub GetTextValue()

Dim acadApp As AcadApplication
Set acadApp = GetObject(, "AutoCAD.Application")
'acadapp.ActiveDocument
Dim activeDoc As AcadDocument
Set activeDoc = acadApp.ActiveDocument

    Dim selObj As Object
    Dim filterType(0) As Integer
    Dim filterValue(0) As Variant
    
    filterType(0) = 0
    filterValue(0) = "MText"

    On Error Resume Next
    activeDoc.SelectionSets.Item("mySelectionSets").Delete
    Set selObj = activeDoc.SelectionSets.Add("mySelectionSets")
    On Error GoTo 0
    
    selObj.SelectOnScreen filterType, filterValue
    
Dim obj As Object
Dim i As Integer

For Each obj In selObj
    'If obj.ObjectName = "ZcadMText" Then
        Dim acadText As AcadMText
        Set acadText = obj
'        Dim newColor As Object
'        Set newColor = acadText.TrueColor
'        newColor.ColorIndex = 1
'        acadText.TrueColor = newColor    'End If

    Dim textValue As String
    textValue = obj.TextString
    textValue = Replace(textValue, "\A1;", "")
    textValue = Replace(textValue, "\pxqr;", "")
    textValue = Replace(textValue, "\pxql;", "")
    textValue = Replace(textValue, "\pxqc;", "")
    
    ActiveCell.Offset(i, 0).Value = textValue
    i = i + 1

Next obj

End Sub

Sub GetHatchArea()
    Dim acadApp As AcadApplication
    Set acadApp = GetObject(, "AutoCAD.Application")
    'acadapp.ActiveDocument
    Dim activeDoc As AcadDocument
    Set activeDoc = acadApp.ActiveDocument
    
    '----------------
    Dim totalsum As VbMsgBoxResult
    totalsum = MsgBox("Kümülatifi görmek istiyor musunuz?", vbYesNo)
    '----------------
    
    Dim selObj As Object
    
    On Error Resume Next
    activeDoc.SelectionSets.Item("mySelectionSets").Delete
    Set selObj = activeDoc.SelectionSets.Add("mySelectionSets")
    selObj.SelectOnScreen
    On Error GoTo 0
    
 Dim obj As Object
 
 For Each obj In selObj
    If obj.ObjectName = "AcDbHatch" Then
        Dim hatch1 As AcadHatch
        Set hatch1 = obj
        Dim area, tArea As Double
        area = hatch1.area
        'MsgBox "Alan: " & area
        Dim newColor As Object
        Set newColor = hatch1.TrueColor
        newColor.ColorIndex = 1
        hatch1.TrueColor = newColor
        
         If totalsum = vbNo Then
            Selection.Value = area / 10000 'cm2 uzunluðunu m2 olarak verir.
            Selection.Offset(1, 0).Select
        End If
        
        tArea = area + tArea
    End If
    Next
    
    'MsgBox "Alan: " & tArea
If totalsum = vbYes Then
    Selection.Value = tArea / 10000 'cm2 uzunluðunu m2 olarak verir.
End If

    selObj.Delete
   
End Sub

Sub GetTextObject()

Dim acadApp As AcadApplication
Dim acadDoc As AcadDocument

On Error Resume Next
Set acadApp = GetObject(, "AutoCAD.Application")
Set acadDoc = acadApp.ActiveDocument

If acadApp Is Nothing Then
    Call MsgBox("AutoCAD dosyasýný baþlatýn.", True)
End If

Dim selectionSet As AcadSelectionSet
Dim textDxfCodes(1) As Integer
Dim textDxfValues(1) As Variant

textDxfCodes(0) = 0: textDxfValues(0) = SelectionObjectTypeName
textDxfCodes(1) = 410: textDxfValues(1) = SelectionSpace

Set selectionSet = acadDoc.SelectionSets.Item("mySelSet")
If selectionSet Is Nothing Then
    Set selectionSet = acadDoc.SelectionSets.Add("mySelSet")
End If
selectionSet.Clear
Err.Clear
 

End Sub
