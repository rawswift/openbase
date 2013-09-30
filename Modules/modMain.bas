Attribute VB_Name = "modMain"
'===============================================
'   nBit's OpenBase Professional version.1.0
'===============================================
'   Core Developer
'-----------------------------------------------
'   Ryan Yonzon
'   e-mail: madpacker360@yahoo.com
'===============================================

'   declare global variables
Public main_conn As New ADODB.Connection
Public test_conn As New ADODB.Connection
Public schema_rs As New ADODB.Recordset
Public view_rs As New ADODB.Recordset
Public query_rs As New ADODB.Recordset

Public sDBFileName As String
Public bTableOpened As Boolean
Public bQueryOpened As Boolean
Public sTableName As String
Public provChoice As Byte

Public sOpenStatement As String
Public bTest As Boolean
Public bConnEstablished As Boolean

'   main application startup
Public Sub main()

    '   show splash screen
    frmSplash.Show
    frmSplash.Refresh
    
    '   set main connection
    Set test_conn = New ADODB.Connection
    
    '   show main form
    frmMain.Show

End Sub

'   end this application
Public Sub kill_app()
    End
End Sub

Public Sub feature()
    MsgBox "Sorry but this feature is not yet available", vbOKOnly + vbInformation
End Sub

Public Sub close_connection()
'   close opened recordset(s) and connection(s)
    If bTableOpened Then
        view_rs.Close
        Set view_rs = Nothing
    End If
    
    If bQueryOpened Then
        query_rs.Close
        Set query_rs = Nothing
        bQueryOpened = False
    End If
    
    schema_rs.Close
    Set schema_rs = Nothing
    main_conn.Close
    Set main_conn = Nothing
    
    '   clear status bar
    frmMain.StatusBar1.Panels(1).Text = ""
    frmMain.StatusBar1.Panels(2).Text = ""
    
End Sub

