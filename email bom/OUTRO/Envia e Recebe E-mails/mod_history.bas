Attribute VB_Name = "mod_history"
Option Explicit

'**********************************
'THIS IS THE MAGIC BEHIND DYNACOM
'YES, THE HISTORY SAVE,LOAD,ADD
'AND CROP ROUTINES, AWE AT THEIR
'SHEER PERFECTION!!!!
'**********************************

'HISTORY SAVE PROCEDURE
Public Sub save_history(ByVal save_combo As ComboBox, filename As String)
    Dim current_entry As Integer
    On Error GoTo FILE_ERROR
    Open filename For Output As #1
        For current_entry = 0 To save_combo.ListCount - 1
            Write #1, save_combo.List(current_entry)
        Next current_entry
    Close #1
    Exit Sub
FILE_ERROR:
    Exit Sub
End Sub

'HISTORY LOAD PROCEDURE
Public Sub load_history(ByVal load_combo As ComboBox, filename As String)
    Dim file_buffer As String
    On Error GoTo FILE_ERROR
    Open filename For Input As #1
        Do While Not EOF(1)
            Input #1, file_buffer
            load_combo.AddItem file_buffer
        Loop
    Close #1
    Exit Sub
FILE_ERROR:
    Exit Sub
End Sub

'HISTORY ADD ITEM PROCEDURE
Public Function add_history(ByVal add_combo As ComboBox, message As String, max_entries As Integer) As Boolean
    If max_entries = 0 Or message = "" Then Exit Function
    Dim entry_exists As Boolean
    Dim check_entry As Integer
    For check_entry = 0 To add_combo.ListCount
        If add_combo.List(check_entry) = message Then entry_exists = True
    Next check_entry
    If (entry_exists = False) Then
        If add_combo.ListCount < max_entries Then
            add_combo.AddItem message
        ElseIf add_combo.ListCount = max_entries Then
            For check_entry = 1 To max_entries - 1
                add_combo.List(check_entry - 1) = add_combo.List(check_entry)
            Next check_entry
            add_combo.List(max_entries - 1) = message
        End If
        add_history = True
    End If
End Function

'HISTORY CROP PROCEDURE
Public Function crop_history(ByVal crop_combo As ComboBox, max_items As Integer) As Integer
    Dim entries_cropped As Integer
    While crop_combo.ListCount > max_items
        crop_combo.RemoveItem (crop_combo.ListCount - 1)
        entries_cropped = entries_cropped + 1
    Wend
    crop_history = entries_cropped
End Function

'VERIFY FILENAME
Public Function verify_file(i_filename As String) As Boolean
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileexists(i_filename) Then
        verify_file = True
    Else
        verify_file = False
    End If
End Function
