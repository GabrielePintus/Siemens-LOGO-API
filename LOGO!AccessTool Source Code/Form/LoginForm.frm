VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "LogIn Panel"
   ClientHeight    =   2950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ip

Public Sub ConnectionHandler(connectionState, code)
    Select Case connectionState
    Case STATE_CONNECTING
        ButtonLogIn.Enabled = False
        InputIP.Enabled = False
        InputPWD.Enabled = False
        CheckBox1.Enabled = False
        ButtonLogIn.Caption = STR(BTN_LOGIN_CONNECTING)
    Case STATE_CONNECTED
        ButtonLogIn.Enabled = True
        InputIP.Enabled = True
        InputPWD.Enabled = True
        CheckBox1.Enabled = False
        ButtonLogIn.Caption = STR(BTN_LOGIN_CONNECTED)
    Case Else ' Including STATE_NOT_CONNECTED
        ButtonLogIn.Enabled = True
        InputIP.Enabled = True
        InputPWD.Enabled = True
         CheckBox1.Enabled = True
        ButtonLogIn.Caption = STR(BTN_LOGIN_NORMAL)
        'm_StopFlag = 0
        setStopFlag (0)
    End Select
End Sub


Public Function GetRunStopMenuButton()
 
    Set TargetBar = Application.CommandBars("Worksheet Menu Bar")
    ' Avoid loaded twice
    Dim menuItem As Object
    
    'Set GetRunStopMenuButton = TargetBar.Controls.item(2)
   For Each menuItem In TargetBar.Controls
   If menuItem.Type = msoControlButton And (menuItem.Caption = STR(MENU_STOP) Or menuItem.Caption = STR(MENU_START)) Then
            ' Found the Menu_Group
   Set GetRunStopMenuButton = menuItem
    Exit Function
  End If
  Next
    
    'Set GetRunStopMenuButton = Nothing ' Return Nothing to indicate Error
End Function
Private Sub ButtonLogIn_Click()
    Dim ipValue
    Dim pwdValue
    Dim runStopMenu
   
    Set runStopMenu = GetRunStopMenuButton()
    If CheckBox1.value <> True Then
        setStopFlag (1)
         runStopMenu.Caption = STR(MENU_START)
          runStopMenu.FaceId = 156
    Else
        runStopMenu.Caption = STR(MENU_STOP)
        runStopMenu.FaceId = 228
    End If
 
   
    
    ipValue = FormatIP(InputIP.value)
    pwdValue = InputPWD.value
    
    ' Validate the IP address
    If ipValue = "" Then
        MsgBox STR(MSG_LOGIN_FAIL_INVALID_IP)
    ElseIf pwdValue = "" Then
        MsgBox STR(MSG_LOGIN_FAIL_INVALID_PASSWORD)
        'Unload LoginForm
    Else
       SetStopFlagChange (1) 'used to help make react more quick.
       SynchronizeToFile
       SetHttpFlag (HTTP_FLAG)
       LogIn ipValue, pwdValue, "LoginFormOnLogInFail", "LoginFormOnLogInSucess", ""
    End If
End Sub

Private Sub TextBox1_Change()

End Sub



Private Sub CheckBox1_Click()

End Sub

Private Sub InputPWD_Change()

End Sub

Private Sub UserForm_Activate()
    LoginFormOnActivate
End Sub

Private Sub UserForm_Terminate()
    LoginFormOnTerminate
End Sub

Private Sub UserForm_Initialize()
    LoginForm.Caption = STR(LOGIN_FORMNAME)
    Label1.Caption = STR(LOGIN_LABEL_IP)
    
    ButtonLogIn.Caption = STR(BTN_LOGIN_NORMAL)
    CheckBox1.Caption = STR(RUN_WHEN_LOGIN)
    CheckBox1.value = True
    
    InputIP.value = GetUrl
End Sub
