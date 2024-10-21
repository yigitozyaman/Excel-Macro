Attribute VB_Name = "M1_CommandBar"
Option Explicit

Sub addInMenu()
Dim cBar As CommandBar
Dim mainMenu, subMenu, subMenu2 As CommandBarControl
Dim macro1 As CommandBarControl

Set cBar = Application.CommandBars(1)
cBar.Reset

Set mainMenu = cBar.Controls.Add(msoControlPopup, , , , True)
With mainMenu
    .Caption = "Makrolar"
    .BeginGroup = True
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Sayfa Numaras� D�zenle"
    .OnAction = "SetPageNo"
    .TooltipText = "Sayfa numaralar�n� d�zenler"
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Metraj Sayfas� Olu�tur"
    .OnAction = "AddQuantitySheet"
    .TooltipText = "Genel icmaldeki poz numaralar�na g�re metraj sayfas� a�ar"
End With


Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Polyline Uzunlu�u Al"
    .OnAction = "GetPolylineLength"
    .TooltipText = "Aktif olan ZwCAD dosyas�ndan se�ilen polylinelar�n toplam uzunlu�unu verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Polyline Alan� Al"
    .OnAction = "GetPolylineArea"
    .TooltipText = "Aktif olan ZwCAD dosyas�ndan se�ilen polylinelar�n toplam alan�n� verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Hatch Alan� Al"
    .OnAction = "GetHatchArea"
    .TooltipText = "Aktif olan ZwCAD dosyas�ndan se�ilen hatchlerin toplam alan�n� verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Text Al"
    .OnAction = "GetTextValue"
    .TooltipText = "Aktif olan AutoCAD dosyas�ndan se�ilen Text'in de�erini verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Donat� Tablosu Olu�tur"
    .OnAction = "DonatiTablosu"
    .TooltipText = "Donat� metraj� ��karmak i�in haz�r tablo format� olu�turur."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "BFA Tablosu Olu�tur"
    .OnAction = "bfaTable"
    .TooltipText = "Birim Fiyat Analizi yapmak i�in haz�r tablo format� olu�turur."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Poz Numaras� Ver"
    .OnAction = "SetPoz"
    .TooltipText = "BFA tablosuna poz numaras� verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Temel Metraj� Hesapla"
    .OnAction = "TemelMetrajHesapla"
    .TooltipText = "1 adet tekil temel i�in metraj hesab� yapar."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Temel Donat� Formu"
    .OnAction = "MetrajFormuKullan"
    .TooltipText = "Betonarme temeller i�in donat� hesab� yapar."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Soket Donat� Formu"
    .OnAction = "SoektDonatiFormKullan"
    .TooltipText = "Betonarme Soketler i�in donat� hesab� yapar."
End With

End Sub



