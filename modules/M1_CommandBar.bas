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
    .Caption = "Sayfa Numarasý Düzenle"
    .OnAction = "SetPageNo"
    .TooltipText = "Sayfa numaralarýný düzenler"
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Metraj Sayfasý Oluþtur"
    .OnAction = "AddQuantitySheet"
    .TooltipText = "Genel icmaldeki poz numaralarýna göre metraj sayfasý açar"
End With


Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Polyline Uzunluðu Al"
    .OnAction = "GetPolylineLength"
    .TooltipText = "Aktif olan ZwCAD dosyasýndan seçilen polylinelarýn toplam uzunluðunu verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Polyline Alaný Al"
    .OnAction = "GetPolylineArea"
    .TooltipText = "Aktif olan ZwCAD dosyasýndan seçilen polylinelarýn toplam alanýný verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Hatch Alaný Al"
    .OnAction = "GetHatchArea"
    .TooltipText = "Aktif olan ZwCAD dosyasýndan seçilen hatchlerin toplam alanýný verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Text Al"
    .OnAction = "GetTextValue"
    .TooltipText = "Aktif olan AutoCAD dosyasýndan seçilen Text'in deðerini verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Donatý Tablosu Oluþtur"
    .OnAction = "DonatiTablosu"
    .TooltipText = "Donatý metrajý çýkarmak için hazýr tablo formatý oluþturur."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "BFA Tablosu Oluþtur"
    .OnAction = "bfaTable"
    .TooltipText = "Birim Fiyat Analizi yapmak için hazýr tablo formatý oluþturur."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Poz Numarasý Ver"
    .OnAction = "SetPoz"
    .TooltipText = "BFA tablosuna poz numarasý verir."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Temel Metrajý Hesapla"
    .OnAction = "TemelMetrajHesapla"
    .TooltipText = "1 adet tekil temel için metraj hesabý yapar."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Temel Donatý Formu"
    .OnAction = "MetrajFormuKullan"
    .TooltipText = "Betonarme temeller için donatý hesabý yapar."
End With

Set macro1 = mainMenu.Controls.Add(msoControlButton, , , , True)
With macro1
    .Caption = "Soket Donatý Formu"
    .OnAction = "SoektDonatiFormKullan"
    .TooltipText = "Betonarme Soketler için donatý hesabý yapar."
End With

End Sub



