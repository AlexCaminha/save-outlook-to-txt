Public Sub TestExport()
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Debug.Print "Namespace obtained"
    Debug.Print "Number of folders: " & ns.Folders.Count
    
    Dim i As Integer
    For i = 1 To ns.Folders.Count
        Debug.Print "Folder " & i & ": " & ns.Folders.Item(i).Name
    Next i
    
    MsgBox "Test completed - check Immediate Window (Ctrl+G)"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
End Sub