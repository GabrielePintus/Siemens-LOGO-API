VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

 'Public m_Interval As Integer 'from the UI
 'Public m_DirNum As Integer
'Sub SetUrl(ip)
Private Sub Class_Initialize()
    m_Interval = 1
    m_DirNum = 5
    SetUrl ("192.168.0.3")
End Sub

Public Sub changeInterval(Interval As Integer)

End Sub

Public Sub changeInerNum(Interval As Integer)

End Sub

