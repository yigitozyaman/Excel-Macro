Attribute VB_Name = "M3_CADfunction"
Option Explicit

Sub ShortCutsCAD()
    Application.OnKey "^+{Ç}", "M3_CADFunction.GetPolylineLength"
    Application.OnKey "^+{Ö}", "M3_CADFunction.GetPolylineArea"
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

