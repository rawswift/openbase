VERSION 5.00
Object = "{2B12169D-6738-11D2-BF5B-00A024982E5B}#29.2#0"; "axbutton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Connectivity"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin axButtonControl.axButton axButton1 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   238
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   -2147483633
      Style           =   3
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Provider"
      TabPicture(0)   =   "frmOpen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Connection"
      TabPicture(1)   =   "frmOpen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).Control(2)=   "Command2(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Advanced"
      TabPicture(2)   =   "frmOpen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton Command2 
         Caption         =   "&Test Connection"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   -71760
         TabIndex        =   9
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Source File Name"
         Height          =   1455
         Index           =   1
         Left            =   -74880
         TabIndex        =   6
         Top             =   2280
         Width           =   5055
         Begin VB.CommandButton Command2 
            Caption         =   "&Browse"
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   480
            TabIndex        =   11
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Select Database"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Connection Source"
         Height          =   1455
         Index           =   0
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   5055
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   "Data Source Name"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Advanced options are not currently available"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -74160
         TabIndex        =   13
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Select the provider to connect to:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   declare private variables
Dim sProvider As String
Dim sFilter As String

'   click event
Private Sub Command1_Click(Index As Integer)
On Error GoTo err_encountered
    Select Case Index
        Case 0
            '   check what provider then test connection
            Select Case List1.ListIndex
                Case 0  '   Advantage
                    bTest = False
                    main_conn.Open ("PROVIDER=Advantage.OLEDB.1;DATA SOURCE=" & Text4.Text & ";SERVERTYPE=ADS_LOCAL_SERVER;TABLETYPE=ADS_ADT;")
                    provChoice = 1
                    bConnEstablished = True
                Case 1  '   ODBC
                    
                    sOpenStatement = ""
                    sOpenStatement = "PROVIDER=MSDASQL;DSN=" & Trim(Text1.Text) & ";"
                
                    bTest = False
                    frmConnPass.Show vbModal
                    
                    'main_conn.Open ("PROVIDER=MSDASQL;DSN=" & Text1.Text)
                    
                    provChoice = 2
                    bConnEstablished = True
                Case 2  '   Jet/Access
                    bTest = False
                    main_conn.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Text4.Text)
                    provChoice = 3
                    bConnEstablished = True
                Case 3
                    bTest = False
                    main_conn.Open ("DSN=" & Text1.Text)
                    provChoice = 4
                    bConnEstablished = True
            End Select
            
            '   enable drop button
            frmMain.Toolbar1.Buttons(3).Enabled = True
            
            If bConnEstablished Then
                '   show database view, information, and tables
                Load frmView
                frmView.Show
                '   unload frmOpen
                Unload Me
            End If
            
        Case 1
            '   enable toolbar button
            frmMain.Toolbar1.Buttons(1).Enabled = True
            '   enable menu --> Open Connection
            frmMain.file_sub_open.Enabled = True
            '   unload this form
            Unload Me
    End Select
        Exit Sub
err_encountered:
    MsgBox "Connection failed", vbOKOnly + vbCritical, "OpenBase Connection"
    bConnEstablished = False
    Exit Sub
End Sub

'   click event
Private Sub Command2_Click(Index As Integer)
On Error GoTo err_conn

    Select Case Index
        Case 0  '   test connection
            '   check what provider then test connection
            Select Case List1.ListIndex
                Case 0  '   Advantage
                    test_conn.Open ("PROVIDER=Advantage.OLEDB.1;DATA SOURCE=" & Trim(Text4.Text) & ";SERVERTYPE=ADS_LOCAL_SERVER;TABLETYPE=ADS_ADT;")
                    MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                    test_conn.Close
                    Set test_conn = Nothing
                Case 1  '   ODBC
                
                    sOpenStatement = ""
                    sOpenStatement = "PROVIDER=MSDASQL;DSN=" & Trim(Text1.Text) & ";"
                
                    bTest = True
                    frmConnPass.Show vbModal
                    
                Case 2  '   Jet/Access
                    
                    test_conn.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Trim(Text4.Text) & ";")
                    MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                    test_conn.Close
                    Set test_conn = Nothing
                    
                Case 3
                    test_conn.Open ("DSN=" & Trim(Text1.Text) & ";")
                    MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                    test_conn.Close
                    Set test_conn = Nothing
            End Select
        
        Case 1  '   browse
            '   filter filename
            CommonDialog1.Filter = sFilter
            '   show dialog open
            CommonDialog1.ShowOpen
            '   get selected filename
            sDBFileName = CommonDialog1.FileName
            '   show selected filename
            Text4.Text = sDBFileName
    End Select
    Exit Sub
err_conn:
    MsgBox "Unexpected Error Occured", vbOKOnly + vbCritical
    Exit Sub
End Sub

'   events when this form is loaded
Private Sub Form_Load()
    
    '   disable this toolbar button
    frmMain.Toolbar1.Buttons(1).Enabled = False
    '   disable menu --> Open Connection
    frmMain.file_sub_open.Enabled = False
    
    '   populate list1
    With List1
        .AddItem "Advantage OLE DB Provider"
        .AddItem "Microsoft OLE DB Provider for ODBC Drivers"
        .AddItem "Microsoft Jet 4.0 OLE DB Provider"
        .AddItem "MySQL ODBC Driver"
    End With
    
        '   set default provider
        List1.ListIndex = 1
    
End Sub
    
'   event list (provider) click
Private Sub List1_Click()
    sProvider = Trim(List1.Text)
End Sub

'   event sstab click
Private Sub SSTab1_Click(PreviousTab As Integer)
    '   check previoustab
    If PreviousTab = 0 Then
    
        '   check what provider
        Select Case List1.ListIndex
            Case 0
                enable_source_file_name
                disable_connection_source
            Case 1
                enable_connection_source
                disable_source_file_name
            Case 2
                enable_source_file_name
                disable_connection_source
            Case 3
                enable_connection_source
                disable_source_file_name
        End Select
    
        '   check what provider
        Select Case List1.ListIndex
            Case 0
                sFilter = "Advantage Data Directory (*.add)|*.add"
            Case 2
                sFilter = "Microsoft Access (*.mdb)|*.mdb"
        End Select
    End If
End Sub

'   enable/disable procedures
Private Sub enable_connection_source()
    Frame1(0).Enabled = True
    Text1.Enabled = True
    Text1.BackColor = &H8000000E
    Label2(0).Enabled = True
End Sub

Private Sub disable_connection_source()
    Frame1(0).Enabled = False
    Text1.Enabled = False
    Text1.BackColor = &H8000000B
    Label2(0).Enabled = False
End Sub

Private Sub disable_source_file_name()
    Frame1(1).Enabled = False
    Label2(3).Enabled = False
    Text4.Enabled = False
    Text4.BackColor = &H8000000B
    Command2(1).Enabled = False
End Sub

Private Sub enable_source_file_name()
    Frame1(1).Enabled = True
    Label2(3).Enabled = True
    Text4.Enabled = True
    Text4.BackColor = &H8000000E
    Command2(1).Enabled = True
End Sub

Private Sub Text1_Change()
    '   enable button Open
    Command2(0).Enabled = True
    Command1(0).Enabled = True
End Sub

Private Sub Text4_Change()
    '   enable button Open
    Command2(0).Enabled = True
    Command1(0).Enabled = True
End Sub
