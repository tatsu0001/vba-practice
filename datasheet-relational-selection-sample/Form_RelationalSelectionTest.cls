VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RelationalSelectionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IRecordSelectEventListener

Private Sub IRecordSelectEventListener_OnRecordSelect(ByVal selectedRecord As Form)
    Dim targetField As Variant
    Dim targetForm As Form
    Set targetForm = Me!�v���_�N�g.Form
    
    ' �Ώۃt�H�[���̌��݂̃��R�[�h�̃R���g���[���̔w�i�F��ς���
    For Each targetField In targetForm
        ' ���x���ȊO(�����t�����̃v���p�e�B�����I�u�W�F�N�g�݂̂����e����)
        If targetField.ControlType <> acLabel Then
            ' �f�[�^�V�[�g�̏����t������ύX
            With targetField.FormatConditions
                .Delete
                With .Add(acExpression, acEqual, "[����]='" & selectedRecord("����").Value & "'")
                    .BackColor = RGB(0, 255, 0)
                End With
            End With
        End If
    Next

    targetForm.Repaint
End Sub


Private Sub Form_Load()
    Me!�v���O���~���O����.Form.RecordSelectEventListener = Me.Form
End Sub