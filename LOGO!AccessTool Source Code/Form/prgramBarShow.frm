VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} prgramBarShow 
   Caption         =   "UserForm1"
   ClientHeight    =   1296
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   3780
   OleObjectBlob   =   "prgramBarShow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "prgramBarShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    prgramBarShow.Caption = STR(TREND_SYN)
End Sub
