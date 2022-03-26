Attribute VB_Name = "STRModule"
Option Private Module
'add the sign format related 13 functions
Public Const STR_VERSION As String = "2.1.0"
Public Const STR_BUILD As String = ".01.200804"
Dim g_text ' The variant that is used to store all texts
Dim g_text_initialized As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Automatically Generated Code starts from here


Public Const BTN_OK As Integer = 0
Public Const ABOUT_LABEL_VERSION_PREFIX As Integer = 1
Public Const ABOUT_LABEL_COPYRIGHT As Integer = 2
Public Const ABOUT_LABEL_ONLINEHELP As Integer = 3
Public Const ABOUT_FORMNAME As Integer = 4
Public Const LOGIN_FORMNAME As Integer = 5
Public Const LOGIN_LABEL_IP As Integer = 6
Public Const LOGIN_LABEL_PSW As Integer = 7
Public Const BTN_LOGIN_NORMAL As Integer = 8
Public Const BTN_LOGIN_CONNECTING As Integer = 9
Public Const BTN_LOGIN_CONNECTED As Integer = 10
Public Const MENU_TOOL_GROUP As Integer = 11
Public Const MENU_ABOUT As Integer = 12
Public Const MSG_CONN_BROKEN As Integer = 13
Public Const MSG_LOGIN_FAIL_INVALID_IP As Integer = 14
Public Const MSG_LOGIN_FAIL_GENERAL As Integer = 15

' new added
Public Const RUN_WHEN_LOGIN As Integer = 16
Public Const MENU_START As Integer = 17
Public Const MENU_STOP As Integer = 18
  'config
Public Const MENU_CONFIG As Integer = 19
Public Const CONFIG_NAME As Integer = 20
Public Const CONFIG_LABEL_INTERVALTIME As Integer = 21
Public Const CONFIG_LABEL_TRENDLENGTH As Integer = 22
Public Const TREND_SYN As Integer = 23
Public Const MENU_ABOUT_TOOLTIP As Integer = 24
Public Const ENCRYT_CHECKBOAX As Integer = 25
Public Const MSG_LOGIN_FAIL_INVALID_PASSWORD = 26
Public Const CONFIG_LABEL_OK As Integer = BTN_OK


Public Sub languageTest()
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Sheets
        Debug.Print sht.name
    Next sht
    
   Set allText = GetTexts(2)

For i = 0 To allText.Count - 1
    S = allText.items()(i)
    Debug.Print allText.keys()(i) & " " & allText(allText.keys()(i))
    Debug.Print S
Next i
End Sub

  Public Function GetTexts(lanID)
    ThisWorkbook.Worksheets("ac_tool_language").Activate
    Set GetTexts = CreateObject("Scripting.Dictionary")
    Dim index As Integer
    index = 2
    Dim key As String
    Dim value As String
    Do While True
        key = ActiveSheet.Cells(index, 1)
        If key = "" Then
            Exit Do
        Else
            value = ActiveSheet.Cells(index, lanID + 2)
            GetTexts.item(index - 2) = value
            index = index + 1
        End If
    Loop
    'GetTexts = Array( _
        Array("OK", "LOGO! Access Tool", "" + ChrW(169) + " Siemens AG 2017", "Online Help", _
            "About LOGO! Access Tool", "Login Panel", "IP Address", "Password", "Log In", _
            "Log In " + ChrW(8230) + "", "Log Out", "LOGO!AccessTool", "About", "Connection Broken", _
            "Invalid IP Address", "LogIn Fail", _
            "Run when log in", "Start", "Stop", _
            "Configure", "Configure Panel", "Synchronization Period", "History Data Number", "History Data Synchoronizing") _
    , _
        Array("OK", "LOGO! Access Tool", "" + ChrW(169) + " Siemens AG 2017", "online hilfe", _
            "" + ChrW(252) + "ber LOGO! Access Tool", "Panel anmeldung", "IP adresse", "passwort", "Anmelden", _
            "Anmelden " + ChrW(8230) + "", "Log Out", "LOGO!AccessTool", "" + ChrW(252) + "ber", "verbindung unterbrochen", _
            "ung" + ChrW(252) + "ltige IP adresse", "login scheitern", _
            "Bei Anmeldung ausf¨¹hren", "Starten", "Stoppen", _
            "Konfigurieren", "Panel konfigurieren", "Synchronisierungszeitraum", "Nummer der Verlaufsdaten", "Synchronisierung der Verlaufsdaten") _
    )
End Function



' Automatically Generated Code End here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Detail information Refer to (https://msdn.microsoft.com/en-us/library/ms912047(WinEmbedded.10).aspx)
Public Function STR(id)
    If Not g_text_initialized Then
        Dim lanID As Integer ' 0 by default (English)
        Select Case GetLanguage()
            Case 3079 ' German Austria
                lanID = 1
            Case 1031 ' German Germany
                lanID = 1
            Case 5127 ' German Liechtenstein
                lanID = 1
            Case 4103 ' German Luxembourg
                lanID = 1
            Case 2055 ' German Switzerland
                lanID = 1
            Case 1041 ' Japanese Japan
                lanID = 2
            Case Else
                lanID = 0
          End Select
        'Initialize g_text before using
        Set g_text = GetTexts(lanID)
        g_text_initialized = True
    End If
    
    STR = g_text(id)
End Function


Public Sub Test()

    MsgBox STR(ABOUT_LABEL_VERSION_PREFIX)
End Sub


