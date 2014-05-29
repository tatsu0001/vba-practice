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
    Set targetForm = Me!プロダクト.Form
    
    ' 対象フォームの現在のレコードのコントロールの背景色を変える
    For Each targetField In targetForm
        ' ラベル以外(条件付書式のプロパティを持つオブジェクトのみを許容する)
        If targetField.ControlType <> acLabel Then
            ' データシートの条件付書式を変更
            With targetField.FormatConditions
                .Delete
                With .Add(acExpression, acEqual, "[言語]='" & selectedRecord("言語").Value & "'")
                    .BackColor = RGB(0, 255, 0)
                End With
            End With
        End If
    Next

    targetForm.Repaint
End Sub


Private Sub Form_Load()
    Me!プログラミング言語.Form.RecordSelectEventListener = Me.Form
End Sub
