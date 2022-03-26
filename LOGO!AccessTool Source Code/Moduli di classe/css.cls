VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "css"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents sht As Worksheet
Attribute sht.VB_VarHelpID = -1
Private Sub sht_Change(ByVal Target As Range)
    On Error GoTo ren
    Set myc.sht = Nothing
    Set myc = Nothing
    'MsgBox "发生变化的区域为:" & Target.Address, , "ChangeDemo"
    Cells(m_CallColumn, 1) = 2
    Cells(m_CallRow, 2) = 8
ren:
End Sub
