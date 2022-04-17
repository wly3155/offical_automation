Attribute VB_Name = "SaveAttachment"
Private Sub SaveAttachment(ByVal Item As Object, path$, Optional condition$ = "*")
    Dim olAtt As Attachment
    Dim i As Integer
    Dim m As Long
    Dim s As String

    If Item.Attachments.Count > 0 Then
        For i = 1 To Item.Attachments.Count
            Set olAtt = Item.Attachments(i)
            If olAtt.FileName Like condition Then
                olAtt.SaveAsFile path & "\" & olAtt.FileName
            End If
        Next
    End If
    Set olAtt = Nothing
End Sub
Public Sub SaveAttach(Item As Outlook.mailitem)
    Dim saveDir As String
    Dim dateFormat

    saveDir = "d:\tmp"
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        On Error Resume Next
    End If

    dateFormat = Format(Now, "yyyy-mm-dd")
    saveDir = saveDir & "\" & dateFormat
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        On Error Resume Next
    End If

    saveDir = saveDir & "\" & Item.Subject
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        On Error Resume Next
    End If

    SaveAttachment Item, saveDir
End Sub
