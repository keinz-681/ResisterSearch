VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResistorSearch 
   Caption         =   "��R���׋@ Ver.0.0.0"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "ResistorSearch.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ResistorSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��R���׋@bykeinz-681
'2022/06/15 ���J

Private Sub CommandButton1_Click()
    Value.Text = num1.Caption & num2.Caption & num3.Caption & String(mul.Caption, "0")
End Sub
Private Sub Error5_Change()
Select Case Error5.ListIndex '���X�g�̔ԍ����Q�Ƃ��A�Ή�����F��\������B
    Case 0 '��
    Error5.BackColor = &HB5
    Error5.ForeColor = &HFFFFFF
    Er.Caption = "+-1%"
    Case 1 '��
    Error5.BackColor = &HFF
    Error5.ForeColor = &HFFFFFF
    Er.Caption = "+-2%"
    Case 2 'F79B32�@��
    Error5.BackColor = &H329BF7
    Error5.ForeColor = &HFFFFFF
    Er.Caption = "+-0.05%"
    Case 3 '��
    Error5.BackColor = &HFF00&
    Error5.ForeColor = &H0
    Er.Caption = "+-0.5%"
    Case 4 '��
    Error5.BackColor = &HFF0000
    Error5.ForeColor = &HFFFFFF
    Er.Caption = "+-0.25%"
    Case 5 '6E2EA8�@��
    Error5.BackColor = &HA82E6E
    Error5.ForeColor = &HFFFFFF
    Er.Caption = "+-0.1%"
    Case 6 'F5D56A�@��
    Error5.BackColor = &H6AD5F5
    Error5.ForeColor = &H0
    Er.Caption = "+-5%"
    Case 7 '��
    Error5.BackColor = &HDEDEDE
    Error5.ForeColor = &H0
    Er.Caption = "+-10%"
    End Select
End Sub
Private Sub First_Change()
    Select Case First.ListIndex '���X�g�̔ԍ����Q�Ƃ��A�Ή�����F��\������B
    Case 0
    First.BackColor = &H0&
    First.ForeColor = &HFFFFFF
    num1.Caption = 0
    Case 1
    First.BackColor = &HB5
    First.ForeColor = &HFFFFFF
    num1.Caption = 1
    Case 2
    First.BackColor = &HFF
    First.ForeColor = &HFFFFFF
    num1.Caption = 2
    Case 3 'F79B32
    First.BackColor = &H329BF7
    First.ForeColor = &HFFFFFF
    num1.Caption = 3
    Case 4
    First.BackColor = &HFFFF&
    First.ForeColor = &H0
    num1.Caption = 4
    Case 5
    First.BackColor = &HFF00&
    First.ForeColor = &H0
    num1.Caption = 5
    Case 6
    First.BackColor = &HFF0000
    First.ForeColor = &HFFFFFF
    num1.Caption = 6
    Case 7 '6E2EA8
    First.BackColor = &HA82E6E
    First.ForeColor = &HFFFFFF
    num1.Caption = 7
    Case 8
    First.BackColor = &HD0D0D0
    First.ForeColor = &H0
    num1.Caption = 8
    Case 9
    First.BackColor = &HFFFFFF
    First.ForeColor = &H0
    num1.Caption = 9
    End Select
End Sub
Private Sub Multiplier_Change()
Select Case Multiplier.ListIndex '���X�g�̔ԍ����Q�Ƃ��A�Ή�����F��\������B
    Case 0
    Multiplier.BackColor = &H0&
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 0
    Case 1
    Multiplier.BackColor = &HB5
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 1
    Case 2
    Multiplier.BackColor = &HFF
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 2
    Case 3 'F79B32
    Multiplier.BackColor = &H329BF7
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 3
    Case 4
    Multiplier.BackColor = &HFFFF&
    Multiplier.ForeColor = &H0
    mul.Caption = 4
    Case 5
    Multiplier.BackColor = &HFF00&
    Multiplier.ForeColor = &H0
    mul.Caption = 5
    Case 6
    Multiplier.BackColor = &HFF0000
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 6
    Case 7 '6E2EA8
    Multiplier.BackColor = &HA82E6E
    Multiplier.ForeColor = &HFFFFFF
    mul.Caption = 7
    Case 8
    Multiplier.BackColor = &HFFFFFF
    Multiplier.ForeColor = &H0
    mul.Caption = -3
    Case 9 'F5D56A
    Multiplier.BackColor = &H6AD5F5
    Multiplier.ForeColor = &H0
    mul.Caption = -1
    Case 10
    Multiplier.BackColor = &HDEDEDE
    Multiplier.ForeColor = &H0
    mul.Caption = -2
    End Select
End Sub
Private Sub Second_Change()
    Select Case Second.ListIndex '���X�g�̔ԍ����Q�Ƃ��A�Ή�����F��\������B
    Case 0
    Second.BackColor = &H0&
    Second.ForeColor = &HFFFFFF
    num2.Caption = 0
    Case 1
    Second.BackColor = &HB5
    Second.ForeColor = &HFFFFFF
    num2.Caption = 1
    Case 2
    Second.BackColor = &HFF
    Second.ForeColor = &HFFFFFF
    num2.Caption = 2
    Case 3 'F79B32
    Second.BackColor = &H329BF7
    Second.ForeColor = &HFFFFFF
    num2.Caption = 3
    Case 4
    Second.BackColor = &HFFFF&
    Second.ForeColor = &H0
    num2.Caption = 4
    Case 5
    Second.BackColor = &HFF00&
    Second.ForeColor = &H0
    num2.Caption = 5
    Case 6
    Second.BackColor = &HFF0000
    Second.ForeColor = &HFFFFFF
    num2.Caption = 6
    Case 7 '6E2EA8
    Second.BackColor = &HA82E6E
    Second.ForeColor = &HFFFFFF
    num2.Caption = 7
    Case 8
    Second.BackColor = &HD0D0D0
    Second.ForeColor = &H0
    num2.Caption = 8
    Case 9
    Second.BackColor = &HFFFFFF
    Second.ForeColor = &H0
    num2.Caption = 9
    End Select
End Sub
Private Sub Third_Change()
    Select Case Third.ListIndex '���X�g�̔ԍ����Q�Ƃ��A�Ή�����F��\������B
    Case 0
    Third.BackColor = &H0&
    Third.ForeColor = &HFFFFFF
    num3.Caption = 0
    Case 1
    Third.BackColor = &HB5
    Third.ForeColor = &HFFFFFF
    num3.Caption = 1
    Case 2
    Third.BackColor = &HFF
    Third.ForeColor = &HFFFFFF
    num3.Caption = 2
    Case 3 'F79B32
    Third.BackColor = &H329BF7
    Third.ForeColor = &HFFFFFF
    num3.Caption = 3
    Case 4
    Third.BackColor = &HFFFF&
    Third.ForeColor = &H0
    num3.Caption = 4
    Case 5
    Third.BackColor = &HFF00&
    Third.ForeColor = &H0
    num3.Caption = 5
    Case 6
    Third.BackColor = &HFF0000
    Third.ForeColor = &HFFFFFF
    num3.Caption = 6
    Case 7 '6E2EA8
    Third.BackColor = &HA82E6E
    Third.ForeColor = &HFFFFFF
    num3.Caption = 7
    Case 8
    Third.BackColor = &HD0D0D0
    Third.ForeColor = &H0
    num3.Caption = 8
    Case 9
    Third.BackColor = &HFFFFFF
    Third.ForeColor = &H0
    num3.Caption = 9
    Case 10
    Third.BackColor = &HFFFFFF
    Third.ForeColor = &H0
    num3.Caption = ""
    End Select
End Sub
Private Sub UserForm_Deactivate()
    Unload Me
End Sub
Private Sub UserForm_Activate()
'�R���{�{�b�N�X�ւ̃A�C�e���ǉ�
    With First '��ꐔ��
    .AddItem "��", 0
    .AddItem "��", 1
    .AddItem "��", 2
    .AddItem "��", 3
    .AddItem "��", 4
    .AddItem "��", 5
    .AddItem "��", 6
    .AddItem "��", 7
    .AddItem "�D", 8
    .AddItem "��", 9
    End With
    With Second '��񐔎�
    .AddItem "��", 0
    .AddItem "��", 1
    .AddItem "��", 2
    .AddItem "��", 3
    .AddItem "��", 4
    .AddItem "��", 5
    .AddItem "��", 6
    .AddItem "��", 7
    .AddItem "�D", 8
    .AddItem "��", 9
    End With
    With Third '��O����
    .AddItem "��", 0
    .AddItem "��", 1
    .AddItem "��", 2
    .AddItem "��", 3
    .AddItem "��", 4
    .AddItem "��", 5
    .AddItem "��", 6
    .AddItem "��", 7
    .AddItem "�D", 8
    .AddItem "��", 9
    .AddItem "����4�{�ł��B", 10
    End With
    With Multiplier '�搔
    .AddItem "��", 0
    .AddItem "��", 1
    .AddItem "��", 2
    .AddItem "��", 3
    .AddItem "��", 4
    .AddItem "��", 5
    .AddItem "��", 6
    .AddItem "��", 7
    .AddItem "��", 8
    .AddItem "��", 9
    .AddItem "��", 10
    End With
    With Error5 '���e��R�l�덷
    .AddItem "��", 0
    .AddItem "��", 1
    .AddItem "��", 2
    .AddItem "��", 3
    .AddItem "��", 4
    .AddItem "��", 5
    .AddItem "��", 6
    .AddItem "��", 7
    End With
    mul.BackColor = ResistorSearch.BackColor
    
End Sub
