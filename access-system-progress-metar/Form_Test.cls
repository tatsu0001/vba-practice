VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' �i�����[�^�[�̊����A���݂̐i���A�i���X�e�b�v��
Private ProgressFinish As Integer
Private ProgressCurrent As Integer
Private ProgressStep As Integer

Private Sub Form_Timer()
    Me!���� = Now()
    
    ' �i���X�V
    ProgressCurrent = ProgressCurrent + ProgressStep
    
    ' ����
    If ProgressCurrent = ProgressFinish Then
        ' ���[�^�p��
        SysCmd acSysCmdRemoveMeter
        ' �^�C�}�[�C�x���g�I��
        Me.TimerInterval = 0
    Else
        ' ���[�^�X�V
        SysCmd acSysCmdUpdateMeter, ProgressCurrent
    End If
End Sub

Private Sub �R�}���h3_Click()
    ProgressFinish = 100
    ProgressCurrent = 0
    ProgressStep = 10

    ' �i�����[�^������
    ' �X�V�̃^�C�}�[�C�x���g�J�n(1�b���ƂɃ��[�^���X�V)
    SysCmd acSysCmdInitMeter, "������...", ProgressFinish
    Me.TimerInterval = 1000
    
    MsgBox "�������ł�" & vbCrLf & "�X�e�[�^�X�o�[�ɐi����\�����܂��B", vbInformation
End Sub