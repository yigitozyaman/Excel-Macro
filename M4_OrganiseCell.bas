Attribute VB_Name = "M4_OrganiseCell"
Option Explicit
Sub CellFormatting()
    Dim Cll As Range
    Set Cll = ActiveCell

    Select Case Cll.NumberFormat
        Case "#,##0.00":
            Cll.NumberFormat = "#,##0.00" & " " & ChrW(&H20BA)
        Case "#,##0.00 " & ChrW(&H20BA):
            Cll.NumberFormat = "#,##0.00" & " " & ChrW(&H20AC)
        Case "#,##0.00 " & ChrW(&H20AC):
            Cll.NumberFormat = "#,##0.00 ""$ """
        Case "#,##0.00 ""$ """:
            Cll.NumberFormat = "#,##0.00 ""LV"""
        Case "#,##0.00 ""LV""":
            Cll.NumberFormat = "#,##0.00"
        Case Else:
            Cll.NumberFormat = "#,##0.00"
    End Select
End Sub

Sub ShortCutOrganiseCell()
    Application.OnKey "^+{Ð}", "M4_OrganiseCell.CellFormatting"
End Sub

Function GetExchangeRate(ExchangeUnit As String) As Double

Dim xmlDoc As MSXML2.DOMDocument60
Dim xmlNode As MSXML2.IXMLDOMNode

Set xmlDoc = New MSXML2.DOMDocument60
xmlDoc.async = False
xmlDoc.Load "https://www.tcmb.gov.tr/kurlar/today.xml"

If ExchangeUnit = "eur" Or ExchangeUnit = "EUR" Or ExchangeUnit = "Eur" Then
    For Each xmlNode In xmlDoc.getElementsByTagName("Currency")
        If xmlNode.Attributes.getNamedItem("CurrencyCode").Text = "EUR" Then
            GetExchangeRate = Val(xmlNode.SelectSingleNode("BanknoteSelling").Text)
        End If
    Next xmlNode
    
ElseIf ExchangeUnit = "usd" Or ExchangeUnit = "USD" Or ExchangeUnit = "Usd" Then
    For Each xmlNode In xmlDoc.getElementsByTagName("Currency")
        If xmlNode.Attributes.getNamedItem("CurrencyCode").Text = "USD" Then
            GetExchangeRate = Val(xmlNode.SelectSingleNode("BanknoteSelling").Text)
        End If
    Next xmlNode
Else
    MsgBox ("Para birimini 'eur' yada 'usd' olarak yazýnýz. Büyük harfe dikkat ediniz.")
End If

Set xmlDoc = Nothing

End Function

'Main Function

Function SpellNumber(ByVal MyNumber)

Dim Dollars, Cents, Temp

Dim DecimalPlace, Count

ReDim Place(9) As String

Place(2) = " Bin "

Place(3) = " Milyon "

Place(4) = " Milyar "

Place(5) = " Trilyon "

' String representation of amount.

MyNumber = Trim(Str(MyNumber))

' Position of decimal place 0 if none.

DecimalPlace = InStr(MyNumber, ".")

' Convert cents and set MyNumber to dollar amount.

If DecimalPlace > 0 Then

Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
"00", 2))

MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

End If

Count = 1

Do While MyNumber <> ""

Temp = GetHundreds(Right(MyNumber, 3))

If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars

If Len(MyNumber) > 3 Then

MyNumber = Left(MyNumber, Len(MyNumber) - 3)

Else

MyNumber = ""

End If

Count = Count + 1

Loop

Select Case Dollars

Case ""

Dollars = "No Dollars"

Case "One"

Dollars = "One Dollar"

Case Else

Dollars = Dollars & " Türk Lirasý"

End Select

Select Case Cents

Case ""

Cents = " and No Cents"

Case "One"

Cents = " and One Cent"

Case Else

Cents = " ve " & Cents & " Kuruþ"

End Select

SpellNumber = Dollars & Cents

End Function


' Converts a number from 100-999 into text

Function GetHundreds(ByVal MyNumber)

Dim Result As String

If Val(MyNumber) = 0 Then Exit Function

MyNumber = Right("000" & MyNumber, 3)

' Convert the hundreds place.

If Mid(MyNumber, 1, 1) <> "0" Then

Result = GetDigit(Mid(MyNumber, 1, 1)) & " Yüz "

End If

' Convert the tens and ones place.

If Mid(MyNumber, 2, 1) <> "0" Then

Result = Result & GetTens(Mid(MyNumber, 2))

Else

Result = Result & GetDigit(Mid(MyNumber, 3))

End If

GetHundreds = Result

End Function


' Converts a number from 10 to 99 into text.


Function GetTens(TensText)

Dim Result As String

Result = "" ' Null out the temporary function value.

If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19...

Select Case Val(TensText)

Case 10: Result = "On"

Case 11: Result = "OnBir"

Case 12: Result = "OnÝki"

Case 13: Result = "OnÜç"

Case 14: Result = "OnDört"

Case 15: Result = "OnBeþ"

Case 16: Result = "OnAltý"

Case 17: Result = "OnYedi"

Case 18: Result = "OnSekiz"

Case 19: Result = "OnDokuz"

Case Else

End Select

Else ' If value between 20-99...

Select Case Val(Left(TensText, 1))

Case 2: Result = "Yirmi "

Case 3: Result = "Otuz "

Case 4: Result = "Kýrk "

Case 5: Result = "Elli "

Case 6: Result = "Altmýþ "

Case 7: Result = "Yetmiþ "

Case 8: Result = "Seksen "

Case 9: Result = "Doksan "

Case Else

End Select

Result = Result & GetDigit _
(Right(TensText, 1)) ' Retrieve ones place.

End If

GetTens = Result

End Function


' Converts a number from 1 to 9 into text.

Function GetDigit(Digit)

Select Case Val(Digit)

Case 1: GetDigit = "Bir"

Case 2: GetDigit = "Ýki"

Case 3: GetDigit = "Üç"

Case 4: GetDigit = "Dört"

Case 5: GetDigit = "Beþ"

Case 6: GetDigit = "Altý"

Case 7: GetDigit = "Yedi"

Case 8: GetDigit = "Sekiz"

Case 9: GetDigit = "Dokuz"

Case Else: GetDigit = ""

End Select

End Function

