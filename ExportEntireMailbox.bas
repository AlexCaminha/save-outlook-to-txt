Option Explicit

' Outlook SaveAs format constants
Const olMSG As Long = 3
Const olTXT As Long = 0

Const BASE_PATH As String = "D:\tempdata\email"
Const MAX_FILENAME_LENGTH As Integer = 100

Private m_stopExport As Boolean

''' <summary>Exports all items from all mailboxes in Outlook to a specified directory.</summary>
''' <remarks>
''' This function serves as the entry point for exporting mailboxes. It initializes the export process,
''' creates necessary directories, and iterates through all mailboxes in the Outlook profile.
''' For each mailbox, it calls the ExportFolder function to handle the actual export of items.
''' It also manages error handling and updates the status bar to inform the user of progress.
''' Upon completion, it displays a message box indicating success or failure.
''' </remarks>
Public Sub ExportEntireMailbox()
    On Error GoTo ErrorHandler
    
    Dim stepInfo As String
    stepInfo = "Starting export"
    m_stopExport = False
    
    ' Create base directory if needed
    stepInfo = "Creating base directory: " & BASE_PATH
    If Dir(BASE_PATH, vbDirectory) = "" Then
        MkDir BASE_PATH
    End If
    
    stepInfo = "Getting namespace"
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    ' Loop through ALL accounts/mailboxes
    Dim olFolder As Outlook.Folder
    Dim i As Integer
    Dim totalAccounts As Integer
    
    stepInfo = "Counting folders"
    totalAccounts = ns.Folders.Count
    
    Debug.Print "Total accounts found: " & totalAccounts
    
    For i = 1 To totalAccounts
        stepInfo = "Getting folder " & i
        Set olFolder = ns.Folders.Item(i)
        
        stepInfo = "Processing folder: " & olFolder.Name
        Debug.Print "Account " & i & ": " & olFolder.Name
        
        ' Update status bar
        On Error Resume Next
        Application.StatusBar = "Exporting account " & i & " of " & totalAccounts & ": " & olFolder.Name
        On Error GoTo ErrorHandler
        
        stepInfo = "Calling ExportFolder for: " & olFolder.Name
        ExportFolder olFolder, BASE_PATH

        If m_stopExport Then Exit For
    Next i
    
    If Not m_stopExport Then
        ' Clear status bar and show completion
        On Error Resume Next
        Application.StatusBar = "Export completed successfully!"

        ' Reset status bar after 3 seconds
        Dim endTime As Double
        endTime = Timer + 3
        Do While Timer < endTime
            DoEvents
        Loop
        Application.StatusBar = ""
        On Error GoTo 0

        MsgBox "All mailboxes exported successfully!", vbInformation
    Else
        On Error Resume Next
        Application.StatusBar = ""
        On Error GoTo 0
    End If
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Application.StatusBar = ""
    On Error GoTo 0
    MsgBox "Error during export at step: " & stepInfo & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical
    Debug.Print "ERROR at: " & stepInfo
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub

''' <summary>Recursively exports all items in the specified Outlook folder to the designated parent path.</summary>
''' <param name="olFolder">The Outlook folder to be exported.</param>
''' <param name="parentPath">The path where the folder's contents will be exported.</param>
''' <remarks>
''' This function processes each item in the folder, saving mail items as TXT files,
''' and calls itself for any subfolders found within the specified folder.
''' It also handles folder creation and checks for path length constraints.
''' </remarks>
Private Sub ExportFolder(ByVal olFolder As Outlook.Folder, ByVal parentPath As String)
    If m_stopExport Then Exit Sub
    On Error Resume Next

    Dim folderPath As String
    folderPath = parentPath & "\" & CleanFolderName(olFolder.Name)
    
    ' Check path length (leave room for filename within Windows 260-char limit)
    If Len(folderPath) > 260 - MAX_FILENAME_LENGTH Then
        Debug.Print "Skipping folder (path too long): " & folderPath
        Exit Sub
    End If
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        If Err.Number <> 0 Then
            Debug.Print "Failed to create folder: " & folderPath & " - Error: " & Err.Description
            Err.Clear
            Exit Sub
        End If
    End If
    
    Dim item As Object
    Dim exportCount As Long
    Dim totalItems As Long
    exportCount = 0
    totalItems = olFolder.Items.Count
    
    ' Update status bar with folder name
    Application.StatusBar = "Processing folder: " & olFolder.Name & " (" & totalItems & " items)"
    
    Dim itemCounter As Long
    itemCounter = 0
    
    For Each item In olFolder.Items
        itemCounter = itemCounter + 1
        
        If TypeOf item Is Outlook.MailItem Then
            ' Update progress every 10 items to avoid slowdown
            If itemCounter Mod 10 = 0 Then
                Application.StatusBar = "Processing: " & olFolder.Name & " - " & itemCounter & " of " & totalItems
                DoEvents ' Allow UI to update
            End If
            
            If SaveMailAsTXT(item, folderPath) Then
                exportCount = exportCount + 1
            End If

            If g_lastSaveError <> 0 Then
                Debug.Print "Error " & g_lastSaveError & " in folder: " & olFolder.Name & ", stopping export."
                MsgBox "Export stopped due to error " & g_lastSaveError & " in folder """ & olFolder.Name & """" & vbCrLf & _
                       "Item: " & itemCounter & " of " & totalItems, vbCritical
                m_stopExport = True
                Exit Sub
            End If
        End If
    Next item
    
    Debug.Print "Exported " & exportCount & " emails from: " & olFolder.Name
    
    Dim subFolder As Outlook.Folder
    For Each subFolder In olFolder.Folders
        If m_stopExport Then Exit For
        ExportFolder subFolder, folderPath
    Next subFolder
End Sub

''' <summary>Cleans and sanitizes a folder name by removing or replacing invalid characters.</summary>
''' <param name="folderName">The original folder name to be cleaned.</param>
''' <returns>A cleaned folder name with invalid characters removed or replaced, suitable for file system operations.</returns>
''' <remarks>
''' This function is typically used to ensure folder names are compatible with file system constraints
''' and can be safely used when exporting or organizing mailbox data.
''' </remarks>
Public Function CleanFolderName(ByVal folderName As String) As String
    If Len(folderName) = 0 Then
        CleanFolderName = "NoName"
        Exit Function
    End If
    
    ' Remove email address symbols that cause issues in paths
    folderName = Replace(folderName, "@", "_")
    folderName = Replace(folderName, ":", "_")
    folderName = Replace(folderName, "\", "_")
    folderName = Replace(folderName, "/", "_")
    folderName = Replace(folderName, "*", "_")
    folderName = Replace(folderName, "?", "_")
    folderName = Replace(folderName, """", "_")
    folderName = Replace(folderName, "<", "_")
    folderName = Replace(folderName, ">", "_")
    folderName = Replace(folderName, "|", "_")
    
    ' Remove other problematic characters
    folderName = Replace(folderName, vbCr, " ")
    folderName = Replace(folderName, vbLf, " ")
    folderName = Replace(folderName, vbTab, " ")
    
    ' Remove multiple spaces
    Do While InStr(folderName, "  ") > 0
        folderName = Replace(folderName, "  ", " ")
    Loop
    
    ' Trim spaces
    folderName = Trim(folderName)
    
    ' Remove trailing dots and spaces
    Do While Len(folderName) > 0 And (Right(folderName, 1) = "." Or Right(folderName, 1) = " ")
        folderName = Left(folderName, Len(folderName) - 1)
    Loop
    
    ' Truncate to MAX_FILENAME_LENGTH to keep paths short
    If Len(folderName) > MAX_FILENAME_LENGTH Then
        folderName = Left(folderName, MAX_FILENAME_LENGTH)
    End If

    ' If empty after cleaning, use default
    If Len(folderName) = 0 Then
        folderName = "NoName"
    End If

    CleanFolderName = folderName
End Function
