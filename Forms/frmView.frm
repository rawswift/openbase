VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmView 
   Caption         =   "Connection Viewer"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9315
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   2760
      ScaleHeight     =   3975
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Table View"
      TabPicture(0)   =   "frmView.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SQL &Query"
      TabPicture(1)   =   "frmView.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   8421504
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   1
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   8421504
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   1
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7320
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":1180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":171A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":1874
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":1CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":2118
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":256A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Information"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6165
      _Version        =   393217
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   declare variables
Dim iIndex As Integer
Dim sTable As String
Dim mNode As Node
Dim iTreeTop As Integer

'   form load event
Private Sub Form_Load()
    
    '   disable Open Connection button from frmMain
        frmMain.Toolbar1.Buttons(1).Enabled = False
        frmMain.file_sub_open.Enabled = False
    
    '   default form arrangement
        frmView.Left = 0
        frmView.Top = 0
    
    '   arrange default location
    ListView1.Move ScaleLeft, ScaleTop, 3000, 3000
    TreeView1.Move ScaleLeft, ListView1.Height + 100, ListView1.Width, ScaleHeight - ListView1.Height - 100
    SSTab1.Move 100 + TreeView1.Width, ScaleTop, frmView.Width - (TreeView1.Width + 225), ScaleHeight
    '   tab 1
    DataGrid1.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 400
    '   tab 2
    DataGrid2.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 2000
    Text1.Move 50, DataGrid2.Height + 100, SSTab1.Width - 100, SSTab1.Height - DataGrid2.Height - 500
    
    '   get connection info
    get_info
    '   get database schema
    Set schema_rs = main_conn.OpenSchema(adSchemaTables)
                
    '   get table names
    schema_rs.Filter = "table_type = 'table'"
    
    '   set root
        With TreeView1
            .Sorted = True
            Set mNode = .Nodes.Add()
            .LabelEdit = False
            .LineStyle = tvwRootLines
        End With
        
        If Trim(main_conn.Provider) = "Microsoft.Jet.OLEDB.4.0" Then
            mNode.Text = sDBFileName
        Else
            mNode.Text = main_conn.DefaultDatabase
        End If
            mNode.Image = 1
        
        '   create parent tree
        Set mNode = TreeView1.Nodes.Add(1, tvwChild, , "Tables", 3)
            mNode.Tag = "noview"    '   set tag for this node
        Set mNode = TreeView1.Nodes.Add(1, tvwChild, , "Views", 4)
            mNode.Tag = "noview"    '   set tag for this node
        
        '   get number of tables in the database
        iIndex = 1
        '   populate tree with table names
        While Not schema_rs.EOF
            Set mNode = TreeView1.Nodes.Add(2, tvwChild, , schema_rs("TABLE_NAME"), 2)
            iIndex = iIndex + 1
            schema_rs.MoveNext
        Wend
        
        schema_rs.MoveFirst
        '   get view names
        
        schema_rs.Filter = "table_type = 'view'"
        
        '   populate tree with view names
        While Not schema_rs.EOF
            Set mNode = TreeView1.Nodes.Add(3, tvwChild, , schema_rs("TABLE_NAME"), 5)
            iIndex = iIndex + 1
            schema_rs.MoveNext
        Wend
        
        '   set status bar
        frmMain.StatusBar1.Panels(1).Width = TreeView1.Width
        frmMain.StatusBar1.Panels(1).Text = iIndex - 1 & " table(s) found"
        
        '   set to false cause no table opened yet!
        bTableOpened = False
        
End Sub

'   query unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '   drop the current connection
    close_connection
    
    '   enable Open Connection button
    frmMain.Toolbar1.Buttons(1).Enabled = True
    frmMain.file_sub_open.Enabled = True
    frmMain.tables_sub_filter.Enabled = False
    frmMain.tables_sub_export.Enabled = False
    
    '   disable Field Description and Drop button
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    
    '   disable filter and export buttons
    frmMain.Toolbar2.Buttons(1).Enabled = False
    frmMain.Toolbar2.Buttons(2).Enabled = False
    

    '   unload this form
    Unload Me
    
End Sub

'   the following codes are used for resizing controls and
'   emulate splitter (simple yet effective)
'   the idea of splitter is from TimeBillingUI by COM Express ;)
'   the following are coded by ryan ...from scratch =)
Private Sub Form_Resize()
    ListView1.Move ScaleLeft, ScaleTop, 4000, 3500
    TreeView1.Move ScaleLeft, ListView1.Height + 100, ListView1.Width, ScaleHeight - ListView1.Height - 100
    SSTab1.Move 100 + TreeView1.Width, ScaleTop, frmView.Width - (TreeView1.Width + 225), ScaleHeight
    '   tab 1
    DataGrid1.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 400
    '   tab 2
    DataGrid2.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 2000
    Text1.Move 50, DataGrid2.Height + 100, SSTab1.Width - 100, SSTab1.Height - DataGrid2.Height - 500

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > ListView1.Width Then
        frmView.MousePointer = 9
    Else
        frmView.MousePointer = 0
    End If
    
    If Button = vbLeftButton Then
        If X > 500 And X < frmView.Width - 1000 Then
            Picture1.Move X, ScaleTop, 100, ScaleHeight
            ListView1.Move ScaleLeft, ScaleTop, Picture1.Left, 3500
            TreeView1.Move ListView1.Left, ListView1.Height + 100, ListView1.Width, ScaleHeight - ListView1.Height - 100
            SSTab1.Move 100 + TreeView1.Width, ScaleTop, frmView.Width - (TreeView1.Width + 225), ScaleHeight
            
            '   tab 1
            DataGrid1.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 400
            '   tab 2
            DataGrid2.Move 50, 50, SSTab1.Width - 100, SSTab1.Height - 2000
            Text1.Move 50, DataGrid2.Height + 100, SSTab1.Width - 100, SSTab1.Height - DataGrid2.Height - 500
            
            frmMain.StatusBar1.Panels(1).Width = TreeView1.Width
            frmView.Refresh
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmView.MousePointer = 0
    Picture1.Visible = False
End Sub

Private Sub get_info()

                ListView1.ListItems.Clear
                
                ListView1.Icons = ImageList2
                ListView1.View = lvwReport
                
                '   default database name
                ListView1.ListItems.Add 1, , "Database Name", , 1
                ListView1.ListItems(1).ListSubItems.Add , , main_conn.DefaultDatabase
                '   provider
                ListView1.ListItems.Add 2, , "Provider", , 2
                ListView1.ListItems(2).ListSubItems.Add , , main_conn.Provider
                '   isolation level
                ListView1.ListItems.Add 3, , "Isolation Level", , 4
                
                    Select Case main_conn.IsolationLevel
                        Case adXactUnspecified:
                            sVariant = "Unspecified"
                        Case adXactChaos:
                            sVariant = "Chaos"
                        Case adXactBrowse:
                            sVariant = "Browse"
                        Case adXactReadUncommitted:
                            sVariant = "Read Uncommitted"
                        Case adXactCursorStability:
                            sVariant = "Cursor Stability"
                        Case adXactReadCommitted:
                            sVariant = "Read Committed"
                        Case adXactRepeatableRead:
                            sVariant = "Repeatable Read"
                        Case adXactIsolated:
                            sVariant = "Isolated"
                        Case adXactSerializable:
                            sVariant = "Serializable"
                    End Select
    
                ListView1.ListItems(3).ListSubItems.Add , , sVariant
                
                '   mode
                ListView1.ListItems.Add 4, , "Mode", , 4
                
                    Select Case main_conn.Mode
                        Case adModeUnknown:
                            sVariant = "Unknown"
                        Case adModeRead:
                            sVariant = "Read"
                        Case adModeWrite:
                            sVariant = "Write"
                        Case adModeReadWrite:
                            sVariant = "Read and Write"
                        Case adModeShareDenyRead:
                            sVariant = "Share, Deny, and Read"
                        Case adModeWrite:
                            sVariant = "Write"
                        Case adModeShareDenyWrite:
                            sVariant = "Share, Deny, and Write"
                        Case adModeShareExclusive:
                            sVariant = "Share Exclusive"
                        Case adModeShareDenyNone:
                            sVariant = "Share, Deny, & None"
                    End Select

                ListView1.ListItems(4).ListSubItems.Add , , sVariant
                
                '   ADO version
                ListView1.ListItems.Add 5, , "ADO Version", , 3
                ListView1.ListItems(5).ListSubItems.Add , , main_conn.Version
                
                   
                   '    other informations
                   ListView1.ListItems.Add 6, , "DBMS Name", , 3
                   ListView1.ListItems(6).ListSubItems.Add , , main_conn.Properties("DBMS Name")
                   
                   ListView1.ListItems.Add 7, , "DBMS Version", , 3
                   ListView1.ListItems(7).ListSubItems.Add , , main_conn.Properties("DBMS Version")
                   
                   ListView1.ListItems.Add 8, , "OLE DB Version", , 3
                   ListView1.ListItems(8).ListSubItems.Add , , main_conn.Properties("OLE DB Version")
                   
                   ListView1.ListItems.Add 9, , "Provider Name", , 3
                   ListView1.ListItems(9).ListSubItems.Add , , main_conn.Properties("Provider Name")
                   
                   ListView1.ListItems.Add 10, , "Provider Version", , 3
                   ListView1.ListItems(10).ListSubItems.Add , , main_conn.Properties("Provider Version")
                
                '   driver name and driver version
                '   error occur when this lines is included on a advantage connection
                'If conn_db_type <> "Advantage" Then
                '    ListView1.ListItems.Add 11, , "Driver Name", , 2
                '    ListView1.ListItems(11).ListSubItems.Add , , main_conn.Properties("Driver Name")
                    '   driver version
                '    ListView1.ListItems.Add 12, , "Driver Version", , 2
                '    ListView1.ListItems(12).ListSubItems.Add , , main_conn.Properties("Driver Version")
                'End If
                
                
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmView.MousePointer = 0
End Sub

Private Sub datagrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmView.MousePointer = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        DataGrid1.Visible = False
        DataGrid2.Visible = True
        Text1.Visible = True
        Text1.SetFocus
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        DataGrid1.Visible = True
        DataGrid2.Visible = False
        Text1.Visible = False
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmView.MousePointer = 0
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        '   do nothing
    End If
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmView.MousePointer = 0
End Sub

'   event treeview single click
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '   change mouse pointer to hourglass
    TreeView1.MousePointer = ccHourglass
    
    '   get table
    sTable = Node.Text
    '   get current node index
    iTreeTop = Node.Index
    
    '   if top level of the tree then assume that is the root
    If iTreeTop = 1 Or Node.Tag = "noview" Then
        '   do nothing
    Else
        '   show form Wait
        frmWait.Label1.Caption = "Please wait while loading table : " & sTable
        frmWait.Show
        frmWait.Refresh
            view_table (sTable)
        Unload frmWait
        
    End If
    
    '   restore default mouse pointer
    TreeView1.MousePointer = 0
    
    '   set sstab to view first tab
    SSTab1.Tab = 0
    
End Sub

'   view table
Private Sub view_table(dTable As String)
On Error GoTo err_open_table
    
    If bTableOpened Then
        view_rs.Close
        Set view_rs = Nothing
    End If
    
    '   check what provider then set cursor location and open recordset
            Select Case provChoice
                Case 2, 3, 4  '   msdasql or jet connection
                    view_rs.CursorLocation = adUseClient
            End Select
            
    view_rs.Open ("select * from " & dTable), main_conn, adOpenStatic, adLockOptimistic
    
    Set DataGrid1.DataSource = view_rs
    
    '   enable field description
    frmMain.Toolbar1.Buttons(2).Enabled = True
    '   enable filter and export
    frmMain.tables_sub_filter.Enabled = True
    frmMain.tables_sub_export.Enabled = True
    
    '   enable filter and export button in toolbar2
    frmMain.Toolbar2.Buttons(1).Enabled = True
    frmMain.Toolbar2.Buttons(2).Enabled = True
    
        Select Case provChoice
            Case 1
                frmMain.export_access.Enabled = True
                frmMain.export_mysql = True
                frmMain.export_advantage.Enabled = False
                    
                '   enable/disable needed button for a specific provider
                frmMain.Toolbar2.Buttons(2).ButtonMenus(1).Enabled = True
                frmMain.Toolbar2.Buttons(2).ButtonMenus(2).Enabled = False
                frmMain.Toolbar2.Buttons(2).ButtonMenus(3).Enabled = True
            
            Case 2, 3
                frmMain.export_advantage.Enabled = True
                frmMain.export_mysql = True
                frmMain.export_access.Enabled = False
                
                '   enable/disable needed button for a specific provider
                frmMain.Toolbar2.Buttons(2).ButtonMenus(1).Enabled = False
                frmMain.Toolbar2.Buttons(2).ButtonMenus(2).Enabled = True
                frmMain.Toolbar2.Buttons(2).ButtonMenus(3).Enabled = True

            Case 4
                frmMain.export_advantage.Enabled = True
                frmMain.export_access.Enabled = True
                frmMain.export_mysql = False
                
                '   enable/disable needed button for a specific provider
                frmMain.Toolbar2.Buttons(2).ButtonMenus(1).Enabled = True
                frmMain.Toolbar2.Buttons(2).ButtonMenus(2).Enabled = True
                frmMain.Toolbar2.Buttons(2).ButtonMenus(3).Enabled = False

        End Select
    
    sTableName = dTable
    bTableOpened = True
    
    frmMain.StatusBar1.Panels(2).Width = DataGrid1.Width
    frmMain.StatusBar1.Panels(2).Text = view_rs.RecordCount & " record(s) found in table : " & sTable
    
    Exit Sub
err_open_table:
    MsgBox "Error opening table : " & dTable, vbOKOnly + vbCritical
    Exit Sub
End Sub
