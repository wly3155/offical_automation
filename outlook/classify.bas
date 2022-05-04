Attribute VB_Name = "classify"
Option Explicit
Option Compare Text

'Private WithEvents inboxItems As Outlook.Items
Private Const SAVED_ROOT_DIR As String = "C:\TMP"
Private Const EXCEL_EXIST_SAVE_DIR As String = "excel"
Private Const UN_CLASSIFY_DIR As String = "Others"
Private Const WORKSHEET_NAME As String = "xyz" 'must using low case string
Private Const KEYWORD As String = "123"
Private Const TIME_09AM As String = "_09_00_00"
Private Const TIME_12AM As String = "_12_00_00"
Private globalSavedDir As String

Private Function checkAndCreateDir(filePath As String)
    If Len(Dir(filePath, vbDirectory)) = 0 Then
        'MsgBox filePath & " not exist, create now"
        mkDirRecursion (filePath)
    End If
End Function

Private Function checkReceivedData(ByVal Item As Object) As String
    Dim recv_date As String
    Dim last_day_012 As String
    Dim now_day_09 As String
    Dim now_day_012 As String
    Dim next_day_09 As String

    last_day_012 = Format(now - 1, "yyyy_mm_dd") & TIME_12AM
    now_day_09 = Format(now, "yyyy_mm_dd") & TIME_09AM
    now_day_012 = Format(now, "yyyy_mm_dd") & TIME_12AM
    next_day_09 = Format(now + 1, "yyyy_mm_dd") & TIME_09AM
    recv_date = Format(Item.ReceivedTime, "yyyy_mm_dd_hh_mm_ss")
    If recv_date < last_day_012 Then
        End 'skip this scenario
    ElseIf recv_date <= now_day_09 Then
        checkReceivedData = last_day_012 & "-" & now_day_09
    ElseIf recv_date > now_day_09 And recv_date < now_day_012 Then
        checkReceivedData = now_day_09 & "-" & now_day_012
    Else
        checkReceivedData = now_day_012 & "-" & next_day_09
    End If

    globalSavedDir = globalSavedDir & "\" & checkReceivedData
    Debug.Print "globalSavedDir after checkReceivedData:" & globalSavedDir
    checkAndCreateDir globalSavedDir
End Function

Private Function readFormDataInExcelFile(filePath As String) As String
    Dim excelApp As Excel.Application
    Dim xWb As Excel.workBook
    Dim xWs As Excel.Worksheet
    Dim cell As Range
    Dim value As String

    Debug.Print filePath
    On Error Resume Next

    Set excelApp = CreateObject("Excel.Application")
    Set xWb = excelApp.Workbooks.Open(filePath, True, True)
    Set xWs = xWb.Worksheets(WORKSHEET_NAME)
    If Err.Number <> 0 Then
        Debug.Print filePath & ":spa not found"
        readFormDataInExcelFile = UN_CLASSIFY_DIR
        GoTo exit_read
    End If

    With xWs.UsedRange
        Set cell = .Find(KEYWORD, LookIn:=xlValues)
        If Not cell Is Nothing Then
            Debug.Print cell.Address
            MsgBox xWs.Cells(cell.Row, cell.Column + 1).value
            value = xWs.Cells(cell.Row, cell.Column + 1).value
            Select Case value
            Case "abc"
                readFormDataInExcelFile = "a"
            Case "def"
                readFormDataInExcelFile = "b"
            Case "gpl"
                readFormDataInExcelFile = "c"
            Case Else
                readFormDataInExcelFile = UN_CLASSIFY_DIR
            End Select
        End If
    End With

exit_read:
    xWb.Close False
    Set xWs = Nothing
    Set xWb = Nothing
End Function

Private Function saveAttachTempory(ByVal attach As Object) As String
    Dim now_time As String
    
    now_time = Format(now, "yyyy_mm_dd_hh_mm_ss")
    attach.SaveAsFile SAVED_ROOT_DIR & "\" & now_time & "_" & attach.filename
    saveAttachTempory = SAVED_ROOT_DIR & "\" & now_time & "_" & attach.filename
End Function

Private Function DestoryTempory(filePath As String)
    Kill filePath
End Function

Private Function doSaveOneAttachment(ByVal olAtt As Object, path$, Optional condition$ = "*")
    If olAtt.filename Like condition Then
        checkAndCreateDir path
        olAtt.SaveAsFile path & "\" & olAtt.filename
    End If
End Function

Private Function doSaveAllAttachment(ByVal Item As Object, path$)
    Dim olAtt As Attachment
    Dim i As Integer
    Dim m As Long
    Dim s As String

    If Item.Attachments.Count > 0 Then
        checkAndCreateDir path
        For i = 1 To Item.Attachments.Count
            Set olAtt = Item.Attachments(i)
            olAtt.SaveAsFile path & "\" & olAtt.filename
        Next
    End If
    Set olAtt = Nothing
End Function

Private Function doSaveMailBody(ByVal Item As Object, path$, Optional condition$ = "*")
    checkAndCreateDir path
    Item.SaveAs path & "\" & Item.Subject & ".msg", OlSaveAsType.olMSG
End Function

Private Function classByAttach(ByVal Item As Object) As String
    Dim olAtt As Attachment
    Dim i As Integer
    Dim tmpXlsPath As String
    Dim tmpSavedDir As String
    Dim excelExist As Boolean
    Dim now_time As String

    Debug.Print "classByAttach start"
    excelExist = False
    If Item.Attachments.Count > 0 Then
        For i = 1 To Item.Attachments.Count
            tmpSavedDir = ""
            Set olAtt = Item.Attachments(i)
            If olAtt.filename Like "*.xls*" Or olAtt.filename Like "*.csv" Then
                Debug.Print "邮件主题：" & Item.Subject & " excel found"
                excelExist = True
                tmpXlsPath = saveAttachTempory(olAtt)
                Debug.Print "start to run readFormDataInExcelFile" & tmpXlsPath
                tmpSavedDir = readFormDataInExcelFile(tmpXlsPath)
                Debug.Print "tmpSavedDir: " & tmpSavedDir
                tmpSavedDir = EXCEL_EXIST_SAVE_DIR & "\" & tmpSavedDir & "\" & Split(olAtt.filename, ".")(0)
                DestoryTempory (tmpXlsPath)
                tmpSavedDir = globalSavedDir & "\" & tmpSavedDir
                If Dir(tmpSavedDir, vbDirectory) <> Empty Then
                    now_time = Format(now, "hh_mm_ss")
                    tmpSavedDir = tmpSavedDir & "_" & now_time
                End If
                Debug.Print "邮件主题：" & Item.Subject & "保存位置：" & tmpSavedDir
                checkAndCreateDir tmpSavedDir
                doSaveOneAttachment olAtt, tmpSavedDir
                doSaveMailBody Item, tmpSavedDir
            End If
        Next
    End If

    If excelExist = False Then
        Debug.Print "邮件主题：" & Item.Subject & " excel not found"
        classByAttach = UN_CLASSIFY_DIR
        tmpSavedDir = globalSavedDir & "\" & classByAttach & "\" & Item.Subject
        Debug.Print tmpSavedDir
        doSaveMailBody Item, tmpSavedDir
        If Item.Attachments.Count > 0 Then
            doSaveAllAttachment Item, tmpSavedDir
        End If
    End If

End Function

Private Function SavedDirInit()
    globalSavedDir = SAVED_ROOT_DIR
    checkAndCreateDir SAVED_ROOT_DIR
End Function

Public Sub main(Item As Outlook.mailitem)
    SavedDirInit
    checkReceivedData Item
    classByAttach Item
End Sub
