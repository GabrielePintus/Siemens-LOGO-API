VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "Configrue Panel"
   ClientHeight    =   2424
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4065
   OleObjectBlob   =   "ConfigForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_temp_interval As Integer
Private m_temp_his_num As Integer
Private Sub OK_Click()
  
 
End Sub
Private Sub ComboBox1_Change()
            On Error GoTo recovery
            If ComboBox1.value = Empty Then
                temp_value = 0
            Else
                 temp_value = CInt(ComboBox1.value)
            End If
            m_temp_interval = temp_value
            Debug.Print m_temp_interval
            Exit Sub
recovery:
            ComboBox1.value = m_temp_interval
            Exit Sub
End Sub

Private Sub ComboBox2_Change()
            On Error GoTo recovery
            If ComboBox2.value = Empty Then
                temp_value = 0
            Else
                 temp_value = CInt(ComboBox2.value)
            End If
            m_temp_his_num = temp_value
            Debug.Print m_temp_his_num
            Exit Sub
recovery:
            ComboBox2.value = m_temp_his_num
            Exit Sub
End Sub

Private Sub Label3_Click()

      m_OldInterval = m_Interval
      m_Interval = ComboBox1.value
      If m_DirNum <> ComboBox2.value Then
      m_OldDirNum = m_DirNum
      m_DirNum = ComboBox2.value
      
      End If
       SynchronizeToFile
            Unload ConfigForm
End Sub

Private Sub TextBox1_Change()
   
End Sub

Private Sub Trendlenth_Click()

End Sub

'1, 10, 30, 60 . The initial value is from ToolConfigManager
Private Sub UserForm_Initialize()
    ConfigForm.Caption = STR(CONFIG_NAME)
    Label1.Caption = STR(CONFIG_LABEL_INTERVALTIME)
    Label2.Caption = STR(CONFIG_LABEL_TRENDLENGTH)
    Label3.Caption = STR(CONFIG_LABEL_OK)
    
     Dim arr1(3) As Integer
     arr1(0) = 1
     arr1(1) = 10
     arr1(2) = 30
     arr1(3) = 60
    ComboBox1.List = arr1
    ComboBox1.value = m_Interval

     For i = 0 To 3
        If arr1(i) = m_Interval Then
         ComboBox1.ListIndex = i
        End If
     Next
     
       Dim arr2(99) As Integer
    For item = 0 To 99 '逐个赋值
        arr2(item) = item + 1
    Next
    
    ComboBox2.List = arr2
    ComboBox2.value = m_DirNum
    For item = 0 To 99 '逐个赋值
     If item + 1 = m_DirNum Then
            ComboBox2.ListIndex = item
        End If
    Next
    
End Sub
