Attribute VB_Name = "LoginFormScript"
Option Private Module
Dim m_ip
Dim m_pwd
Dim m_saved As Boolean

Public Sub LoginFormOnLogInSucess(arg0, status)
    LoginForm.Hide
End Sub

Public Sub LoginFormOnLogInFail(arg0, status)
    MsgBox STR(MSG_LOGIN_FAIL_GENERAL)
End Sub

Public Sub LoginFormConnectionHandler(connectionState, code)
    LoginForm.ConnectionHandler connectionState, code
End Sub

Public Sub LoginFormOnTerminate()
    m_ip = LoginForm.InputIP.value
    m_pwd = LoginForm.InputPWD.value

    m_saved = True
End Sub

Public Sub LoginFormOnActivate()
    If m_saved Then
        m_saved = False ' It can only be used for first time after terminate
        
        LoginForm.InputIP.value = m_ip
        LoginForm.InputPWD.value = m_pwd
    End If
    
    AddPropertyListener PROPERTY_ID_CONNECTION, "LoginFormConnectionHandler"
End Sub
