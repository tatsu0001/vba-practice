VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents Connection As ADODB.Connection
Attribute Connection.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Connection = Nothing
End Sub

Public Sub OpenConnection(ByVal dsn As String)
    Set Connection = New ADODB.Connection
    Connection.Open dsn
End Sub

Public Sub Disconnect()
    Connection.Close
    Set Connection = Nothing
End Sub

Public Function Execute(ByVal query As String) As ADODB.Recordset
    Set Execute = Connection.Execute(query)
End Function


Private Sub Connection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print pConnection.ConnectionString & " connect complete."
End Sub

Private Sub Connection_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print pConnection.ConnectionString & " connect complete."
End Sub

Private Sub Connection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print pCommand.CommandText & " execute complete."
End Sub
