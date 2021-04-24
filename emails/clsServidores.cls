VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsservidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Private DoingRequery As Boolean
Public Event MoveComplete()

Private Sub Class_Initialize()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & InAppPath2("List.mdb") & ";"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Email,Servidor,Nome from servidores Order by Servidor", db, adOpenStatic, adLockOptimistic

  DataMembers.Add "Primary"
  Set db = Nothing
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Select Case DataMember
  Case "Primary"
    Set Data = adoPrimaryRS
  End Select
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  RaiseEvent MoveComplete
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Public Property Get EditingRecord() As Boolean
  EditingRecord = (adoPrimaryRS.EditMode <> adEditNone)
End Property

Public Property Get AbsolutePosition() As Long
  AbsolutePosition = adoPrimaryRS.AbsolutePosition
End Property

Public Sub AddNew()
  adoPrimaryRS.AddNew
End Sub

Public Sub Delete()
  adoPrimaryRS.Delete
    MoveNext
End Sub

Public Sub Requery()
  adoPrimaryRS.Requery
  DataMemberChanged "Primary"
End Sub

Public Sub Update()
  With adoPrimaryRS
    .UpdateBatch adAffectAll
    If .EditMode = adEditAdd Then
      MoveLast
    End If
  End With
End Sub

Public Sub Cancel()
  With adoPrimaryRS
    .CancelUpdate
    If .EditMode = adEditAdd Then
      MoveFirst
    End If
  End With
End Sub

Public Sub MoveFirst()
  adoPrimaryRS.MoveFirst
End Sub

Public Sub MoveLast()
  adoPrimaryRS.MoveLast
End Sub

Public Sub MoveNext()
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
End Sub

Public Sub MovePrevious()
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
End Sub

Private Function InAppPath2(ByVal File As String, Optional Dir As String = "False") As String
Dim strDir As String
If Not Dir = "False" Then
    Dir = FixName(Dir)
    Dexist AppPath2 & Dir, True
    strDir = AppPath2 & Dir
Else
    strDir = AppPath2()
End If
InAppPath2 = strDir & File
End Function

Private Function AppPath2() As String
AppPath2 = FixName(App.Path)
End Function

Private Function FixName(ByVal Name As String) As String
If Not Right(Name, 1) = "\" Then
    FixName = Name & "\"
Else
    FixName = Name
End If
End Function

Private Function Dexist(strDirName As String, Optional Criar As Boolean = False) As Boolean
Dim Dexist1 As Boolean
    Dim strDummy As String

    ' Check for Existence of Directory
    On Error Resume Next
    strDirName = FixName(strDirName)
    strDummy = Dir$(strDirName & "*.*", vbDirectory)
    Dexist1 = Not (strDummy = "")
    If Criar Then
        If Not Dexist1 Then
            MkDir strDirName
            Dexist1 = True
        End If
    End If
    Dexist = Dexist1
    Dir$ AppPath, vbDirectory
    Err = 0
End Function

Private Sub Class_Terminate()
Set adoPrimaryRS = Nothing
End Sub
