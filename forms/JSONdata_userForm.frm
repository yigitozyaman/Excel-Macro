VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JSONdata_userForm 
   Caption         =   "Data Details"
   ClientHeight    =   3708
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3588
   OleObjectBlob   =   "JSONdata_userForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JSONdata_userForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
 sub_data.SaveDataToJson
End Sub
