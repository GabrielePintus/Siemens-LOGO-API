VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private objApp                  As Object 'ָ��ǰ���ĵ�����ThisDocument��ThisWorkBook
Private uForm                   As Object '����������
Private lbl1                    As Object '��ʾ��ǩ���� MSForms.Label
Private lbl2                    As Object '��ʾ���� MSForms.Label
Private FormName                As String

'������
Private Const GWL_STYLE         As Long = (-16)
Private Const WS_CAPTION        As Long = &HC00000
Private Const BarLength         As Long = 300            '����������

#If Win64 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub Class_Initialize()
    '��Office����ִ������������õ�BUG����ʹ��VBA���������ɾ�����ٴ���ͬ������ᱨ�����ļ�/·�����ʴ���
    '΢���Ѿ�ȷ���˸�BUG�Ĵ��� http://support.microsoft.com/kb/244238/zh-cn
    '��Ҫÿ��������ʱ�򣬴���һ�������ƵĴ���
    t = Timer
    ms = t - Int(t) '�������
    FormName = "FORM" & Format(Now, "ddhhmmss") & Replace(ms, ".", "")
End Sub

'����������
Public Sub ShowBar()
    CreateProgressBar
End Sub

'���ٽ�����
Public Sub DestroyBar()
    Unload uForm
    RemoveModual FormName
    Set uForm = Nothing
    Set objApp = Nothing
End Sub

'���ý���������
Public Sub ChangeProcessBarValue(value As Double, Optional message As String = "")
On Error Resume Next

    lbl1.Width = Int(value * BarLength) '��ʾ������
    lbl2.Caption = IIf(message = "", Format(value, "�Ѿ���� 0.00%"), message)
    DoEvents    'ת�ÿ���Ȩ������ϵͳ
    
End Sub

'��������
Public Sub SleepBar(ms As Long)
    Sleep ms
End Sub

'��������������
Private Sub CreateProgressBar()

    Dim UsForm  As Object
    
    If InStr(1, Application.name, "Word") > 0 Then
        Set objApp = ThisDocument
    ElseIf InStr(1, Application.name, "Excel") > 0 Then
        Set objApp = ThisWorkbook
    End If

    '����һ�����塣�����ж����С�
    RemoveModual FormName
    
    Set UsForm = objApp.VBProject.VBComponents.Add(3) 'vbext_ct_MSForm
    With UsForm
        '���ڸô��廹δ���У��൱�ڴ������״̬
        '���ڸô�������ԣ���Ҫ��Properties���Է���
        .Properties("Caption") = "����"
        .Properties("Name") = FormName
        .Properties("Height") = 30
        .Properties("Width") = BarLength
        .Properties("BackColor") = RGB(240, 240, 240)
        .Properties("SpecialEffect") = fmSpecialEffectFlat
        .Properties("BorderStyle") = fmBorderStyleNone
        'Ҫ�ڸô����ϴ����ؼ�������Ҫ����.Designer���������
    End With

    '���ز���ʾ�ô��塣ע����ƽʱ������ʾ����Ĳ�ͬ
    Set uForm = VBA.UserForms.Add(FormName)
    
    With uForm
        
        '������ʾ����
        Set lbl1 = .Controls.Add("Forms.Label.1", "Label1", True)
        With lbl1
            .Left = 0
            .Top = 0
            .Height = uForm.Width
            .Width = 0
            .Caption = ""
            .BackColor = RGB(128, 128, 255)
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleOpaque
            .BorderColor = .BackColor
            .ZOrder 1
        End With
        
        '������ʾ����
        Set lbl2 = .Controls.Add("Forms.Label.1", "Label1", True)
        With lbl2
            .Left = 0
            .Top = 9
            .Height = 12
            .Width = BarLength
            .Caption = ""
            .TextAlign = fmTextAlignLeft
            .Font.size = 9
            .Font.Bold = False
            .Font.Italic = False
            .Font.name = "����"""
            .ForeColor = RGB(0, 0, 0)
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            .ZOrder 0
        End With
        
        RemoveFormCaption uForm
        uForm.Show vbModeless
        
    End With
    
End Sub


Private Sub RemoveModual(n As String) '�Ƴ�����ָ�����Ƶ�ģ��
On Error Resume Next
    objApp.VBProject.VBComponents.Remove objApp.VBProject.VBComponents(n)
    objApp.Save
End Sub


Private Sub RemoveFormCaption(form As Object)

    If val(Application.Version) < 9 Then
        hwnd = FindWindow("ThunderXFrame", form.Caption)
    Else
        hwnd = FindWindow("ThunderDFrame", form.Caption)
    End If
    IStyle = GetWindowLong(hwnd, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    SetWindowLong hwnd, GWL_STYLE, IStyle
    DrawMenuBar hwnd
    
End Sub

