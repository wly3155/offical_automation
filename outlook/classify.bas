Attribute VB_Name = "classify"
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
        End 'skip class this scenario
    ElseIf recv_date <= now_day_09 Then
        classByReceivedData = last_day_012 & "-" & now_day_09
    ElseIf recv_date > now_day_09 And recv_date < now_day_012 Then
        classByReceivedData = now_day_09 & "-" & now_day_012
    Else
        classByReceivedData = now_day_012 & "-" & next_day_09
    End If
End Function
Private Function findKeyWord(ByVal sheet As Object) As String
'    r = sheet.Range("a1:z500").Find(Range("G3"), Lookat:=xlWhole)
'    MsgBox sheet.Range("a1:d5")
'    MsgBox r
End Function
Private Function classByAttach(ByVal Item As Object) As String
    Dim olAtt As Attachment
    Dim i As Integer
    Dim j As Integer
    Dim tmpXlsPath As String
    Dim exlApp As Excel.Application
    Dim exlWorkBook As Excel.workBook
    Dim exlSheet As Excel.Worksheet
    Dim exlRange As Excel.Range
    Dim myrange1
    Dim hangs1
    Dim lies1
    Dim keyWord As String

    keyWord = "fdasf"
    classByAttach = "others"
    Debug.Print keyWord
    
    Set exlApp = New Excel.Application
    With exlApp
        .Visible = True
        .EnableEvents = True
        .UserControl = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .ScreenUpdating = True
    End With

    Debug.Print Item.Attachments.Count
    If Item.Attachments.Count > 0 Then
        For i = 1 To Item.Attachments.Count
            Set olAtt = Item.Attachments(i)
            Debug.Print olAtt.FileName
            If olAtt.FileName Like "*.xls*" Or olAtt.FileName Like "*.csv" Then
                Debug.Print olAtt.FileName
                tmpXlsPath = "c:\temp\" & olAtt.FileName
                Debug.Print tmpXlsPath
                olAtt.SaveAsFile tmpXlsPath
                Set exlWorkBook = exlApp.Workbooks.Open("c:\temp\test.xlsx")
                Set exlSheet = exlWorkBook("sheet2")
                exlSheet.Activate
                exlSheet.Columns(1).Font.Bold = True
'                areaCount = Selection.Areas.Count

'                Set exlSheet = exlWorkBook.Sheets("sheet2")
'                Set exlSheet = exlWorkBook.ActiveSheet
'                Debug.Print exlSheet.Name
'                Debug.Print exlWorkBook.Name
'                exlApp.ScreenUpdating = True
'                Set exlRange = exlSheet.Columns("A:A").Find("*", oSheet.[a1], xlValues, , xlByRows, xlPrevious)
'                exlRange.Cells.ClearContents


'                tmpXlsPath = "c:\temp\test.xlsx"
'                'Import process (Error starts next line)
'                tmpXlsPath = Application.GetOpenFilename("Text Files (*.PRN),*.PRN", , "Please select text file...")
'                With wb.QueryTables.Add(Connection:="TEXT;" & tmpXlsPath, Destination:=ws.Range("A1"))
'                    .TextFileParseType = xlDelimited
'                    .TextFileCommaDelimiter = True
'                    .Refresh
'                End With
'
'                'open File
'                Set exlWorkBook = exlApp.Workbooks.Open(tmpXlsPath, , , , , , , , , True)
'                Set exlSheet = exlWorkBook.Worksheets("Sheet2")
'                exlWorkBook.Activate

 'Save and Close
     'Workbooks("BigPic 2019.xlsx").Close SaveChanges:=True
'                MsgBox exlWorkBook.Sheets.Count
'                Debug.Print exlWorkBook.Sheets.Count
'                For j = 1 To exlWorkBook.Sheets.Count
'                    Set exlSheet = exlWorkBook.Sheets(j).Name
'                    MsgBox exlSheet
'                    If exlSheet.Name = "asdfspasdfa" Or exlSheet.Name = "asfx" Then
'                        MsgBox "´æÔÚ"
'                        MsgBox sheet.Range("a1:z500").Find("fdasf", Lookat:=xlWhole)
'                        findKeyWord (sheet)
'                        With exlSheet.Range("a1:z500")
'                            Set c = .Find(keyWord, LookIn:=xlValues)
'                            If Not c Is Nothing Then
'                                MsgBox c.Address
'                                Exit For
'                            End If
'                        End With
'                    Else
'                        MsgBox "²»´æÔÚ"
'                    End If
'                Next
'
'                Kill "c:\temp\" & olAtt.FileName
                exlWorkBook.Close False
                exlApp.Quit
            End If
        Next
    End If
End Function
Private Function choseAndCreateSaveDir(Item As Outlook.mailitem) As String
    Dim saveDir As String
    Dim recvDateClassRes As String
    Dim attchClassRes As String
    
    saveDir = "c:\tmp"
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

    Debug.Print "classbyattach"
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
            If olAtt.FileName Like condition Then
                olAtt.SaveAsFile path & "\" & olAtt.FileName
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

