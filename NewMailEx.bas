Option Explicit

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Const BASE_PATH As String = "D:\tempdata\email"

    On Error GoTo ErrorHandler

    Debug.Print "NewMailEx fired: " & EntryIDCollection

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim mail As Outlook.MailItem
    Set mail = ns.GetItemFromID(EntryIDCollection)

    If Not mail Is Nothing Then
        ' Get the account name from the parent folder's store
        Dim accountName As String
        accountName = mail.Parent.Store.DisplayName

        ' Build path: BASE_PATH\AccountName\FolderName\
        Dim accountPath As String
        accountPath = BASE_PATH & "\" & CleanFolderName(accountName)

        Dim folderPath As String
        folderPath = accountPath & "\" & CleanFolderName(mail.Parent.Name)

        ' Create base directory if needed
        If Dir(BASE_PATH, vbDirectory) = "" Then
            MkDir BASE_PATH
        End If

        ' Create folder structure
        If Dir(accountPath, vbDirectory) = "" Then
            MkDir accountPath
        End If
        If Dir(folderPath, vbDirectory) = "" Then
            MkDir folderPath
        End If

        SaveMailAsTXT mail, folderPath

        Debug.Print "Auto-exported: " & mail.Subject
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "NewMailEx error: " & Err.Number & ": " & Err.Description
End Sub
