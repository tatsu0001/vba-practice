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

' 進捗メーターの完了、現在の進捗、進捗ステップ数
Private ProgressFinish As Integer
Private ProgressCurrent As Integer
Private ProgressStep As Integer

Private Sub Form_Timer()
    Me!時刻 = Now()
    
    ' 進捗更新
    ProgressCurrent = ProgressCurrent + ProgressStep
    
    ' 完了
    If ProgressCurrent = ProgressFinish Then
        ' メータ廃棄
        SysCmd acSysCmdRemoveMeter
        ' タイマーイベント終了
        Me.TimerInterval = 0
    Else
        ' メータ更新
        SysCmd acSysCmdUpdateMeter, ProgressCurrent
    End If
End Sub

Private Sub コマンド3_Click()
    ProgressFinish = 100
    ProgressCurrent = 0
    ProgressStep = 10

    ' 進捗メータ初期化
    ' 更新のタイマーイベント開始(1秒ごとにメータを更新)
    SysCmd acSysCmdInitMeter, "処理中...", ProgressFinish
    Me.TimerInterval = 1000
    
    MsgBox "処理中です" & vbCrLf & "ステータスバーに進捗を表示します。", vbInformation
End Sub
