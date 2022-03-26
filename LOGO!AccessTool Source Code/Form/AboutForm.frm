VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutForm 
   Caption         =   "About LOGO! Access Tool"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "AboutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()
  'MsgBox "Click down"
End Sub
Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 1) And (Shift = 2) Then
    Label1.Caption = STR(ABOUT_LABEL_VERSION_PREFIX) + " " + STR_VERSION + STR_BUILD
   ' MsgBox "left shift down"
    End If


End Sub
Private Sub Label2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If (Button = 1) And (Shift = 2) Then
    Label1.Caption = STR(ABOUT_LABEL_VERSION_PREFIX) + " " + STR_VERSION
    'MsgBox "left shift down"
    End If


End Sub

'Private Sub Label3_Click()
'    Shell "explorer.exe http://www.baidu.com/"
'End Sub

Private Sub OK_Click()
    Unload AboutForm
End Sub

Private Sub UserForm_Initialize()
    OK.Caption = STR(BTN_OK)
    Label1.Caption = STR(ABOUT_LABEL_VERSION_PREFIX) + " " + STR_VERSION
    Label2.Caption = STR(ABOUT_LABEL_COPYRIGHT)
    'Label3.Caption = STR(ABOUT_LABEL_ONLINEHELP)
    AboutForm.Caption = STR(ABOUT_FORMNAME)
End Sub
