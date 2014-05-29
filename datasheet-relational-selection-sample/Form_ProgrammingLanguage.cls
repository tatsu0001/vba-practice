VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ProgrammingLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private recordSelectEventHandler As IRecordSelectEventListener

Public Sub Form_Load()
    Set recordSelectEventHandler = Nothing
End Sub

Public Property Let RecordSelectEventListener(ByVal listener As IRecordSelectEventListener)
    Set recordSelectEventHandler = listener
End Property

Private Sub Form_Current()
    If Not recordSelectEventHandler Is Nothing Then
        recordSelectEventHandler.OnRecordSelect Me
    End If
End Sub
