Attribute VB_Name = "classify"
Option Explicit

'Private WithEvents inboxItems As Outlook.Items
Private Const SAVED_ROOT_DIR As String = "C:\TMP\"
Private Const UN_CLASSIFY_DIR As String = "Others\"
Private Const WORKSHEET_NAME As String = "spa"
Private Const KEYWORD As String = "vendor"

Private Function classByReceivedData(ByVal Item As Object) As String
    Dim recv_date As String
    Dim last_day_012 As String
    Dim now_day_09 As String
    Dim now_day_012 As String
    Dim next_day_09 As String

    last_day_012 = Format(Now - 1, "yyyy_mm_dd") & "_12_00_00"
    now_day_09 = Format(Now, "yyyy_mm_dd") & "_09_00_00"
    now_day_012 = Format(Now, "yyyy_mm_dd") & "_12_00_00"
    next_day_09 = Format(Now + 1, "yyyy_mm_dd") & "_09_00_00"
    recv_date = Format(Item.ReceivedTime, "yyyy_mm_dd_hh_mm_ss")
    If recv_date < last_day_012 Then
        End 'skip this scenario
    ElseIf recv_date <= now_day_09 Then
        classByReceivedData = last_day_012 & "-" & now_day_09
    ElseIf recv_date > now_day_09 And recv_date < now_day_012 Then
        classByReceivedData = now_day_09 & "-" & now_day_012
    Else
        classByReceivedData = now_day_012 & "-" & next_day_09
    End If

    MsgBox classByReceivedData
End Function

Private Function readFormDataInExcelFile(filepath As String) As String
    Dim excelApp As Excel.Application
    Dim xWb As Excel.workBook
    Dim xWs As Excel.Worksheet
    Dim cell As Range
    
    readFormDataInExcelFile = UN_CLASSIFY_DIR
    Set excelApp = CreateObject("Excel.Application")
    Set xWb = excelApp.Workbooks.Open(filepath, True, True)
    Set xWs = xWb.Worksheets(WORKSHEET_NAME)
    If xWs Is Nothing Then
        Set xWs = xWb.Worksheets(LCase(WORKSHEET_NAME))
        If xWs Is Nothing Then
            End
        End If
    End If
    
    With xWs.UsedRange
        Set cell = .Find(KEYWORD, LookIn:=xlValues)
        If Not cell Is Nothing Then
            Debug.Print cell.Address
            Debug.Print "find row: " & cell.Row & "find colum: " & cell.Column
            MsgBox "find row: " & cell.Row & "find colum: " & cell.Column
            MsgBox xWs.Cells(cell.Row, cell.Column + 1).Value
            readFormDataInExcelFile = xWs.Cells(cell.Row, cell.Column + 1).Value
        End If
    End With

    MsgBox readFormDataInExcelFile
    xWb.Close False
    Set xWs = Nothing
    Set xWb = Nothing

End Function

Private Function saveAttachTempory(ByVal attach As Object) As String
    attach.SaveAsFile SAVED_ROOT_DIR & attach.filename
    saveAttachTempory = SAVED_ROOT_DIR & attach.filename
End Function

Private Function DestoryTempory(filepath As String)
    Kill filepath
End Function

Private Function classByAttach(ByVal Item As Object) As String
    Dim olAtt As Attachment
    Dim i As Integer
    Dim tmpXlsPath As String

    classByAttach = UN_CLASSIFY_DIR
    If Item.Attachments.Count > 0 Then
        For i = 1 To Item.Attachments.Count
            Set olAtt = Item.Attachments(i)
            If olAtt.filename Like "*.xls*" Or olAtt.filename Like "*.csv" Then
                tmpXlsPath = saveAttachTempory(olAtt)
                classByAttach = readFormDataInExcelFile(tmpXlsPath)
                DestoryTempory (tmpXlsPath)
            End If
        Next
    End If
End Function

Private Function choseAndCreateSaveDir(Item As Outlook.mailitem) As String
    Dim saveDir As String
    Dim recvDateClassRes As String
    Dim attchClassRes As String
    
    saveDir = SAVED_ROOT_DIR
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        MsgBox saveDir
        On Error Resume Next
    End If

    recvDateClassRes = classByReceivedData(Item)
    saveDir = saveDir & "\" & recvDateClassRes
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        MsgBox saveDir
        On Error Resume Next
    End If

    attchClassRes = classByAttach(Item)
    saveDir = saveDir & "\" & attchClassRes
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir (saveDir)
        MsgBox saveDir
        On Error Resume Next
    End If

    choseAndCreateSaveDir = saveDir
    Debug.Print choseAndCreateSaveDir
End Function

Private Function doSaveAttachment(ByVal Item As Object, path$, Optional condition$ = "*")
    Dim olAtt As Attachment
    Dim i As Integer
    Dim m As Long
    Dim s As String

    If Item.Attachments.Count > 0 Then
        For i = 1 To Item.Attachments.Count
            Set olAtt = Item.Attachments(i)
            If olAtt.filename Like condition Then
                olAtt.SaveAsFile path & "\" & olAtt.filename
            End If
        Next
    End If
    Set olAtt = Nothing
End Function

Private Function doSaveMailBody(ByVal Item As Object, path$, Optional condition$ = "*")
    Item.SaveAs path & "\" & Item.Subject & ".msg", OlSaveAsType.olMSG
End Function

Public Sub main(Item As Outlook.mailitem)
    Dim saveDir As String

    saveDir = choseAndCreateSaveDir(Item)
    doSaveAttachment Item, saveDir
    doSaveMailBody Item, saveDir
End Sub
