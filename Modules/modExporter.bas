Attribute VB_Name = "modExporter"
Public sQueryStatement As String
Public iExportCounter As Long
Public iFieldCount As Integer
Public sType As String
Dim iRecordCount As Long
Dim iCounter As Long
Dim recCounter As Long
Dim sInsertStatement As String
Dim dName As String
Dim sName As String
Dim sfd As String       '   formatted date
Dim sft As String       '   formatted date

Dim err_pass As Long    '   error counter

Public dest_conn As New ADODB.Connection
Public bDestOpened As Boolean

Public Sub construct_statement()

    iFieldCount = view_rs.Fields.Count
    
    sQueryStatement = "create table " & sTableName & " ("
    
    iExportCounter = 0
    
    '   get field's name, type, and defined size
    While iExportCounter < iFieldCount
    
        Select Case view_rs.Fields(iExportCounter).Type
            Case adArray
                    sType = "Array"
                        MsgBox "array"
            Case adBigInt
                    sType = "Big Integer"
                        MsgBox "big int"
            Case adBinary
                    sType = "Binary"
                        MsgBox "binary"
            Case adBoolean                                     '   boolean
                    sType = "logical"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adbyref       '   ???
                    sType = "ByRef"
            Case adBSTR
                    sType = "BSTR"
            Case adChar
                    sType = "Character"
                        MsgBox "character"
            Case adCurrency
                    sType = "Currency"
            Case adDate
                    sType = "Date"
            Case adDBDate                                      '   date
                    sType = "date"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adDBTime
                    'sType = "Database Time"
                    sType = "time"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adDBTimeStamp
                    sType = "Database Time Stamp"
            Case adDecimal
                    sType = "Decimal"
                    MsgBox "decimal"
            Case adDouble                                      '   double
                    sType = "double"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adEmpty
                    sType = "Empty"
            Case adError
                    sType = "Error"
            Case adGUID
                    sType = "GUID"
            Case adIDispatch
                    sType = "ID Dispatch"
            Case adInteger                                     '   integer
                    sType = "int"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adIUnknown
                    sType = "Unknown Type"
                        MsgBox "unknown type"
            Case adLongVarBinary
                    '   i haven't found yet the conversion of blob type to access
                    '   so i use long
                    sType = "long"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adLongVarChar
                    sType = "memo"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
            Case adLongVarWChar
                    sType = "Long Variant Word Character"
                        MsgBox "long variant word character"
            Case adNumeric
                    sType = "Numeric"
                        MsgBox "numeric"
            Case adSingle
                    sType = "Single"
            Case adSmallInt
                    sType = "Small Integer"
                        MsgBox "small integer"
            Case adTinyInt
                    sType = "tinyint"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType
                    
            Case adUnsignedInt
                    sType = "Unsigned Integer"
                        MsgBox "unsigned integer"
            Case adUnsignedSmallInt
                    sType = "Unsigned Small Integer"
            Case adUnsignedTinyInt
                    sType = "Unsigned Tiny Integer"
            Case adUserDefined
                    sType = "Defined"
            Case adVarBinary
                    sType = "Variant Binary"
            Case adVarChar                                    '   character
                    sType = "char"
                    sQueryStatement = sQueryStatement & view_rs(iExportCounter).Name & "_" & " " & sType & "(" & view_rs(iExportCounter).DefinedSize & ")"
            Case adVariant
                    sType = "Variant"
            Case advector
                    sType = "Vector"
        End Select
    
        iExportCounter = iExportCounter + 1
        
        If iExportCounter = iFieldCount Then
            'do nothing
        Else
            sQueryStatement = sQueryStatement & ","
        End If
        
    Wend
    
    sQueryStatement = sQueryStatement & ");"
    
End Sub

Public Sub create_table()
    dest_conn.Execute (sQueryStatement)
End Sub

Public Sub transfer()
On Error GoTo err_encountered
    
    err_counter = 0
    
    frmTransfer.Show
    frmTransfer.Refresh
    
    iRecordCount = view_rs.RecordCount
    
    '   set progress bar max value
    frmTransfer.ProgressBar1.Max = iRecordCount
    recCounter = 0
    
    view_rs.MoveFirst

    While Not view_rs.EOF
    
    '   reset counter
    iCounter = 0
    
    sInsertStatement = "insert into " & sTableName & " ("
        
        While iCounter < iFieldCount
            sInsertStatement = sInsertStatement & view_rs(iCounter).Name & "_"
            iCounter = iCounter + 1
            If iCounter = iFieldCount Then
                '   do nothing
            Else
                sInsertStatement = sInsertStatement & ","
            End If
        Wend
        
        sInsertStatement = sInsertStatement & ") values ("
        
    '   reset counter
    iCounter = 0
        
        '   this lines of code filters records
        While iCounter < iFieldCount
            
            
            If IsNull(view_rs(iCounter)) Then
                        
                sInsertStatement = sInsertStatement & "NULL"
                        
            Else
            
                If view_rs(iCounter).Type = adVarChar Then
                    dName = Trim(view_rs(iCounter))
                    sName = Replace(dName, "'", "_")
                    sInsertStatement = sInsertStatement & "'" & sName & "'"
                ElseIf view_rs(iCounter).Type = adDBDate Then
                    sfd = FormatDateTime(Trim(view_rs(iCounter)), vbShortDate)
                    sInsertStatement = sInsertStatement & "'" & sfd & "'"
                ElseIf view_rs(iCounter).Type = adLongVarBinary Then
                    '   i haven't found yet the convertion type for blob to access
                    '   so no value will be copied
                    sInsertStatement = sInsertStatement & "NULL"
                ElseIf view_rs(iCounter).Type = adLongVarChar Then
                    sInsertStatement = sInsertStatement & "'" & view_rs(iCounter) & "'"
                ElseIf view_rs(iCounter).Type = adDBTime Then
                    sft = FormatDateTime(Trim(view_rs(iCounter)), vbLongTime)
                    sInsertStatement = sInsertStatement & "'" & sft & "'"
                Else
                    sInsertStatement = sInsertStatement & view_rs(iCounter)
                End If
                
            End If
            
            iCounter = iCounter + 1
            
            If iCounter = iFieldCount Then
                    '   do nothing
            Else
                sInsertStatement = sInsertStatement & ","
            End If
        
        Wend
        
            sInsertStatement = sInsertStatement & ");"
        
        '   insert record
        dest_conn.Execute (sInsertStatement)
        
        frmTransfer.Label2.Caption = recCounter & " of " & iRecordCount
        frmTransfer.Refresh
        
        recCounter = recCounter + 1
        
        frmTransfer.ProgressBar1.Value = recCounter
        
        view_rs.MoveNext
       
    Wend
    
        '   close destination
        dest_conn.Close
        Set dest_conn = Nothing
        
        Unload frmTransfer
        
        MsgBox "Done", vbOKOnly + vbInformation
        Exit Sub
        
err_encountered:

    err_counter = err_counter + 1
    frmTransfer.Label3.Caption = "Pass Through Error : " & err_counter
    Resume Next

End Sub
