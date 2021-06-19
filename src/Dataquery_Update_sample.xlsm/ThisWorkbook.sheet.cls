Option Explicit

Public Sub updateStatus()

Dim I As Byte
Dim rngRules As Range, rngRow As Range
Dim objConnection As Object
Dim sSQL As String, sWhere As String, sStatus As String
Dim sRuleColumn As String, sRuleValue As String
Dim bScreenUpdating As Boolean

    ' Don't refresh the screen every time
    bScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Use late binding to reduce configuration problems on some computers
    Set objConnection = CreateObject("ADODB.Connection")
    If (objConnection Is Nothing) Then
        MsgBox "Sorry, it seems you don't have ADODB.Connection on your computer." & vbCrLf & vbCrLf & _
            "The macro present in this workbook requires that DLL to be able to do his job", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    shData.Activate
    
    ' Prepare the connection to the spreadsheet
    Set objConnection = CreateObject("ADODB.Connection")
    
    ' We'll work on this workbook, it's a .xlsm file and we've the first row with headers
    objConnection.Open _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source='" & ThisWorkbook.FullName & "';" & _
        "Extended Properties=""Excel 12.0 Macro;HDR=Yes;"";"
    
    ' Reset the status cell
    objConnection.Execute "UPDATE [" & ActiveSheet.Name & "$] SET [Status]='' "
        
    ' Process our rules
    Set rngRules = shRules.Range("rngRules")
    
    sWhere = ""
    
    For Each rngRow In rngRules.Rows
        ' Don't process the first row of the rules; it's the header row
        If rngRow.Row > 3 Then
        
            ' Get the status to apply
            sStatus = rngRow.Cells(1, 5).Value
            
            ' Nothing to do if the status cell is empty
            If Not (Trim(sStatus) = "") Then
                ' There are four columns in order to determine the status
                ' i.e.  client / gender / ZipCode / Comment
                For I = 1 To 4
                    If (rngRow.Columns(I).Value <> "") Then
                    
                        sRuleColumn = rngRules.Cells(1, I).Value
                        sRuleValue = rngRow.Cells(1, I).Value
                        
                        ' We must make sure to not use single quote if the cell contains a number
                        If (sRuleColumn <> "ZipCode") Then
                            sRuleValue = "'" & sRuleValue & "'"
                        End If
                        
                        sWhere = sWhere & "([" & sRuleColumn & "] = " & sRuleValue & ") AND "
                    End If
                Next
            
                ' Remove the final " AND "
                If Not (Trim(sWhere) = "") Then
                    sWhere = Left(sWhere, Len(sWhere) - 5)
                End If
        
                sSQL = "UPDATE [" & ActiveSheet.Name & "$] SET [Status]='" & sStatus & "' WHERE " & sWhere
                
                'Debug.Print "Row " & rngRow.Row & " - *" & sWhere & "*"
                'Debug.Print sSQL
                
                ' Run the update i.e. set the status in the Data sheet
                objConnection.Execute sSQL
                
                ' Reset the variable
                sWhere = ""
            End If
        End If
    Next
    
    ' Close and release
    objConnection.Close
    Set objConnection = Nothing
    
    ' Restore
    Application.ScreenUpdating = bScreenUpdating
    
End Sub