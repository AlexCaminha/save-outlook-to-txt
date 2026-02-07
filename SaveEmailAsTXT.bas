Option Explicit

Public g_lastSaveError As Long

' SaveEmailAsTXT Function
'
' This function is responsible for saving an email in the .txt format, which is a standard format for email messages.
' It returns a Boolean value indicating the success or failure of the operation.
'
' The function utilizes the CleanFileName function to ensure that the filename generated for the saved email
' adheres to valid naming conventions, avoiding any illegal characters that may cause issues when saving the file.
'
' Parameters:
'   mail - The Outlook.MailItem object to be saved as .txt format.
'   folderPath - The directory path where the .txt file will be saved.
'
' Returns:
'   Boolean - True if the email was successfully saved as .txt, False otherwise.
Public Function SaveMailAsTXT(ByVal mail As Outlook.MailItem, ByVal folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    g_lastSaveError = 0

    Dim fileName As String
    Dim timeStamp As String
    Dim subject As String

    ' Validate received time
    On Error Resume Next
    timeStamp = Format(mail.ReceivedTime, "yyyy-MM-dd HHmmss")
    If Err.Number <> 0 Then
        timeStamp = Format(mail.CreationTime, "yyyy-MM-dd HHmmss")
        If Err.Number <> 0 Then
            timeStamp = Format(Now, "yyyy-MM-dd HHmmss")
        End If
    End If
    On Error GoTo ErrorHandler

    subject = CleanFileName(mail.subject)

    ' Truncate subject if too long
    If Len(subject) > 100 Then
        subject = Left(subject, 100)
    End If

    ' Handle empty subjects
    If Len(subject) = 0 Then
        subject = "NoSubject"
    End If

    fileName = timeStamp & " - " & subject & ".txt"

    ' Check total path length
    Dim fullPath As String
    fullPath = folderPath & "\" & fileName

    ' Windows path limit check
    If Len(fullPath) > 240 Then
        subject = Left(subject, 30)
        fileName = timeStamp & " - " & subject & ".txt"
        fullPath = folderPath & "\" & fileName
    End If

    ' Skip if already saved (e.g. from a previous run)
    If Dir(fullPath) <> "" Then
        Debug.Print "  Already exists, skipping: " & fullPath
        SaveMailAsTXT = True
    Else
        ' Debug output
        Debug.Print "Saving to: " & fullPath

        ' Try to save
        mail.SaveAs fullPath, olTXT

        ' Verify file was created
        If Dir(fullPath) <> "" Then
            SaveMailAsTXT = True
            Debug.Print "  Success!"
        Else
            Debug.Print "  File not found after save"
            SaveMailAsTXT = False
        End If
    End If

    ' Save attachments
    If mail.Attachments.Count > 0 Then
        On Error Resume Next
        Dim att As Outlook.Attachment
        Dim attCleanName As String
        Dim attExt As String
        Dim attBase As String
        Dim attFileName As String
        Dim attFullPath As String
        Dim extPos As Long
        Dim maxNameLen As Long

        ' Max filename length that keeps full path within 240 chars
        maxNameLen = 240 - Len(folderPath) - 1

        Dim a As Long
        For a = 1 To mail.Attachments.Count
            Set att = mail.Attachments.Item(a)

            ' Clean and split attachment filename into base + extension
            attCleanName = CleanFileName(att.fileName)
            If Len(attCleanName) = 0 Or attCleanName = "NoName" Then
                attCleanName = "attachment_" & a
            End If

            extPos = InStrRev(attCleanName, ".")
            If extPos > 0 Then
                attExt = Mid(attCleanName, extPos)
                attBase = Left(attCleanName, extPos - 1)
            Else
                attExt = ""
                attBase = attCleanName
            End If

            ' Build filename: timestamp - subject - attachment.ext
            attFileName = timeStamp & " - " & subject & " - " & attBase & attExt

            ' Truncate progressively to fit within path limit
            If Len(attFileName) > maxNameLen Then
                ' Try shortening subject (keep 6 chars for two " - " separators)
                Dim availSubj As Long
                availSubj = maxNameLen - Len(timeStamp) - 6 - Len(attBase) - Len(attExt)
                If availSubj > 0 Then
                    attFileName = timeStamp & " - " & Left(subject, availSubj) & " - " & attBase & attExt
                Else
                    ' Drop subject: timestamp - attachment.ext
                    attFileName = timeStamp & " - " & attBase & attExt
                    If Len(attFileName) > maxNameLen Then
                        ' Truncate attachment name, preserve extension
                        Dim availAtt As Long
                        availAtt = maxNameLen - Len(timeStamp) - 3 - Len(attExt)
                        If availAtt > 0 Then
                            attFileName = timeStamp & " - " & Left(attBase, availAtt) & attExt
                        Else
                            attFileName = timeStamp & attExt
                        End If
                    End If
                End If
            End If

            attFullPath = folderPath & "\" & attFileName

            ' Skip if attachment already exists
            If Dir(attFullPath) <> "" Then
                Debug.Print "  Attachment already exists, skipping: " & attFullPath
            Else
                att.SaveAsFile attFullPath

                If Err.Number <> 0 Then
                    Debug.Print "  Failed to save attachment: " & att.fileName & " - Error " & Err.Number & ": " & Err.Description
                    Err.Clear
                Else
                    Debug.Print "  Saved attachment: " & attFullPath
                End If
            End If
        Next a
        On Error GoTo 0
    End If

    Exit Function

ErrorHandler:
    g_lastSaveError = Err.Number
    Debug.Print "Failed to save: " & mail.subject
    Debug.Print "  Path: " & fullPath
    Debug.Print "  Error " & Err.Number & ": " & Err.Description
    Debug.Print "  Path length: " & Len(fullPath)

    If Err.Number = -2147024809 Then
        Debug.Print "  === E_INVALIDARG details ==="
        Debug.Print "  folderPath arg: " & folderPath
        Debug.Print "  fullPath built: " & fullPath
        Debug.Print "  fileName built: " & fileName
        Debug.Print "  timeStamp: " & timeStamp
        Debug.Print "  subject (cleaned): " & subject
        On Error Resume Next
        Debug.Print "  mail.Subject (raw): " & mail.subject
        Debug.Print "  mail.ReceivedTime: " & mail.ReceivedTime
        Debug.Print "  mail.SenderName: " & mail.SenderName
        Debug.Print "  mail.EntryID: " & mail.EntryID
        Debug.Print "  SaveAs type: " & olTXT
        On Error GoTo 0
        Debug.Print "  =============================="
    End If

    ' Last resort: try with minimal filename
    On Error Resume Next
    fileName = timeStamp & ".txt"
    fullPath = folderPath & "\" & fileName
    If Dir(fullPath) <> "" Then
        Debug.Print "  Minimal filename already exists, skipping: " & fullPath
        SaveMailAsTXT = True
    Else
        Debug.Print "  Trying minimal filename: " & fullPath
        mail.SaveAs fullPath, olTXT
        If Dir(fullPath) <> "" Then
            SaveMailAsTXT = True
            Debug.Print "  Success with minimal filename!"
        Else
            SaveMailAsTXT = False
        End If
    End If
End Function

''' <summary>Cleans and sanitizes a file name by removing or replacing invalid characters.</summary>
''' <param name="fileName">The original file name to be cleaned.</param>
''' <returns>A cleaned file name with invalid characters removed or replaced, suitable for file system operations.</returns>
''' <remarks>
''' This function is typically used to ensure file names are compatible with file system constraints
''' and can be safely used when exporting or organizing mailbox data.
''' </remarks>
Public Function CleanFileName(ByVal fileName As String) As String
    If Len(fileName) = 0 Then
        CleanFileName = "NoName"
        Exit Function
    End If
    
    ' Remove email address symbols that cause issues in file names
    fileName = Replace(fileName, "@", "_at_")
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    Dim c As Variant
    For Each c In invalidChars
        fileName = Replace(fileName, CStr(c), "_")
    Next c
    
    ' Replace other problematic characters
    fileName = Replace(fileName, vbCr, " ")
    fileName = Replace(fileName, vbLf, " ")
    fileName = Replace(fileName, vbTab, " ")
    
    ' Remove multiple spaces
    Do While InStr(fileName, "  ") > 0
        fileName = Replace(fileName, "  ", " ")
    Loop
    
    ' Trim spaces
    fileName = Trim(fileName)
    
    ' Remove trailing dots and spaces
    Do While Len(fileName) > 0 And (Right(fileName, 1) = "." Or Right(fileName, 1) = " ")
        fileName = Left(fileName, Len(fileName) - 1)
    Loop
    
    ' If empty after cleaning, use default
    If Len(fileName) = 0 Then
        fileName = "NoName"
    End If
    
    CleanFileName = fileName
End Function