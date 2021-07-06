
Public Function MergeArrays(Arr1() As Variant, Arr2() As Variant) As Variant
    Dim arr1UB As Integer
    Dim new_arr() As Variant
    arr1UB = 0
    
    If Not Not Arr1 Then 'Testing to see if the first array has values
        For i = LBound(Arr1) To UBound(Arr1) 'Storing the contents of array 1 parameter into new array
            ReDim Preserve new_arr(i)
            new_arr(i) = Arr1(i)
            arr1UB = arr1UB + i
        Next
    Else 'Setting the upper bound to 0 since the first array is empty
        arr1UB = 0
    End If
    
    
    
    For i = LBound(Arr2) To UBound(Arr2) 'Storing the contents of array 2 parameter into new array
        ReDim Preserve new_arr(i + arr1UB)
        new_arr(i + arr1UB) = Arr2(i)
    Next i

    MergeArrays2 = new_arr
    
End Function



Public Function convToTable(wsName As String, startCell As String) As ListObject
    Dim StartPoint As Range
    Dim LastColumn As Long
    Dim LastRow As Long
    Dim ws As Worksheet

    Set ws = ActiveWorkbook.Sheets(wsName)
    Set StartPoint = ws.Range(startCell)
    LastRow = ws.Cells(ws.rows.count, StartPoint.Column).End(xlUp).row
    LastColumn = ws.Cells(StartPoint.row, ws.Columns.count).End(xlToLeft).Column
    
    ws.Activate
    ws.Range(StartPoint, ws.Cells(LastRow, LastColumn)).Select
    
    Set convToTable = ws.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    
End Function

Public Function Remove_excess(catalog As String, choice As Integer) As String
    Dim i As Long
    Select Case choice
        Case 1
            Remove_excess = ""
            
            For i = 1 To Len(catalog)
                If Not UCase(Mid(catalog, i, 1)) Like "[/(), ]" Then
                    Remove_excess = Remove_excess & Mid(catalog, i, 1)
                End If
            Next i
        Case 2
            Remove_excess = ""
            
            For i = 1 To Len(catalog)
                If Not UCase(Mid(catalog, i, 1)) Like "()" Then
                    Remove_excess = Remove_excess & Mid(catalog, i, 1)
                End If
            Next i
    End Select
    
    
End Function

Public Function nextEmptyCell(sheetName As String, address As String) As Range
    Dim startCell
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim col As String
    
    Set ws = ActiveWorkbook.Sheets(sheetName)
    Set startCell = ws.Range(address)
    
    LastRow = ws.Cells(ws.rows.count, startCell.Column).End(xlUp).row
    
    col = ""
    For i = 1 To Len(address)
        If UCase(Mid(address, i, 1)) Like "[A-Z]" Then
            col = col & Mid(address, i, 1)
        End If
    Next
    
    Set nextEmptyCell = ws.Cells(LastRow + 1, col)
    
End Function
Public Function getRowNum(address As String) As Integer
    'Getting the row number for starting row in number of rows to insert
    'In main code, the number of records in the recordset will determine the ending row number--therefore total rows needed to insert
    row = ""
    For i = 1 To Len(address)
        If UCase(Mid(address, i, 1)) Like "[0-9]" Then
            row = row & Mid(address, i, 1)
        End If
    Next
    
    getRowNum = CInt(row)

End Function
Public Function getRange(sheetName As String, address As String) As Range
    Dim startCell As Range
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    
    Set ws = ActiveWorkbook.Sheets(sheetName)
    Set startCell = ws.Range(address)
    
    LastRow = ws.Cells(ws.rows.count, startCell.Column).End(xlUp).row
    LastColumn = ws.Cells(ws.Columns.count).End(xlToLeft).Column
    
    Set getRange = ws.Range(startCell, ws.Cells(LastRow, LastColumn))

End Function

Public Function getParts(tblName As ListObject, colNum As Integer, dtype As String) As String
    Dim strCrit As String
    Select Case dtype
        Case "String"
            If tblName.ListRows.count > 1 Then
                For i = 1 To tblName.ListRows.count
                    If i = 1 Then
                        strCrit = "('" & tblName.ListColumns(colNum).DataBodyRange(i) & "',"
                    ElseIf i = tblName.ListRows.count Then
                        strCrit = strCrit & "'" & tblName.ListColumns(colNum).DataBodyRange(i) & "')"
                    Else
                        strCrit = strCrit & "'" & tblName.ListColumns(colNum).DataBodyRange(i) & "',"
                    End If
                Next
            Else
                strCrit = "('" & tblName.ListColumns(colNum).DataBodyRange(i) & "')"
            End If
            getParts = strCrit
        Case "Integer"
            If tblName.ListRows.count > 1 Then
                For i = 1 To tblName.ListRows.count
                    If i = 1 Then
                        strCrit = "(" & tblName.ListColumns(colNum).DataBodyRange(i) & ","
                    ElseIf i = tblName.ListRows.count Then
                        strCrit = strCrit & tblName.ListColumns(colNum).DataBodyRange(i) & ")"
                    Else
                        strCrit = strCrit & tblName.ListColumns(colNum).DataBodyRange(i) & ","
                    End If
                Next
            End If
            getParts = strCrit
    End Select
End Function

Public Sub clearResults(sheetName As String, address As String, tblName As String)
    Dim ws As Worksheet
    Dim startCell As Range
    Dim clearRange As Range
    Dim lo As ListObject
    
    Application.ScreenUpdating = False
    Set ws = ActiveWorkbook.Sheets(sheetName)
    Set startCell = ws.Range(address)
    
    If startCell.value <> "" Then
        For Each lo In ws.ListObjects
            If lo.Name = tblName Then
                lo.Unlist
            End If
        Next lo
        LastRow = ws.Cells(ws.rows.count, startCell.Column).End(xlUp).row
        LastColumn = ws.Cells(startCell.row, ws.Columns.count).End(xlToLeft).Column
        Set clearRange = ws.Range(startCell, ws.Cells(LastRow, LastColumn))
        clearRange.ClearContents
        clearRange.Interior.color = xlNone
        clearRange.Borders.color = xlNone
    End If
'    startCell.value = ""
'    clearResults = startCell.value
    Application.ScreenUpdating = True
End Sub

Public Function tableExists(tblName As String, sheetName As String) As Boolean
    Dim ListObject As ListObject
    For Each ListObject In ActiveWorkbook.Sheets(sheetName).ListObjects
        Select Case ListObject.Name
            Case tblName
                tableExists = True
        End Select
    Next ListObject
End Function


'Public Sub send_email(recipient As String, content As String, wb As Workbook)
'    'GoTo tools-> references -> enable Microsoft Outlook 15.0 Object Library
'    'Used in bom tree application for sending emails or mismatching records
'
'    Dim MailApp As Outlook.Application
'    Dim MailItem As Outlook.MailItem
'
'    Dim source As String 'For attachment file
'    Dim body As String 'For email content
'
'    Set MailApp = New Outlook.Application
'    Set MailItem = New Outlook.MailItem
'
'    MailItem.To = recipient
'
'    MailItem.HTMLBody = "" & _
'    "<b> Mismatching records found, part number : " & content & " --see attached </b>"
'    source = wb.FullName
'    MailItem.Attachments.Add source
'    MailItem.Send
'
'End Sub
