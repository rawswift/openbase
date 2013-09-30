VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   Caption         =   "Export"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1_Access 
      Caption         =   "Destination"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test Connection"
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test Connection"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Specific Database"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Data Source Name"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDBPath As String

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0  '   test connection 1
        
            If Len(Text1.Text) = 0 Then
                MsgBox "Please enter Data Source Name", vbOKOnly + vbInformation
            Else
                '   check what provider then test connection
                Select Case sExportTo
                    Case "access"  '   Access
                        test_conn.Open ("PROVIDER=MSDASQL;DSN=" & Text1.Text)
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                    Case "mysql"
                        test_conn.Open ("DSN=" & Text1.Text)
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                End Select
            End If
    
        Case 1  '   test connection 2
        
            If Len(Text2.Text) = 0 Then
                MsgBox "Please select specific database name", vbOKOnly + vbInformation
            Else
                '   check what provider then test connection
                Select Case sExportTo
                    Case "advantage"  '   Advantage
                        test_conn.Open ("PROVIDER=Advantage.OLEDB.1;DATA SOURCE=" & Text2.Text & ";SERVERTYPE=ADS_LOCAL_SERVER;TABLETYPE=ADS_ADT;")
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                    Case "access"  '   Jet/Access
                        test_conn.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Text2.Text)
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                End Select
            End If
        
        Case 2  '   browse
            
            CommonDialog1.Filter = "Microsoft Access (*.mdb)|*.mdb"
            '   show dialog open
            CommonDialog1.ShowOpen
            '   get selected filename
            sDBPath = CommonDialog1.FileName
            '   show selected filename
            Text2.Text = sDBPath

    End Select
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            
            Select Case sExportTo
                    Case "access"  '   Access
                        If Option1(0) Then
                            dest_conn.Open ("PROVIDER=MSDASQL;DSN=" & Text1.Text)
                            MsgBox "Connection Opened : MSDSQL"
                        ElseIf Option1(1) Then
                            dest_conn.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Text2.Text)
                            MsgBox "Connection Opened : JET"
                        End If
                            'construct_statement
                            
                    Case "mysql"
                        
                        test_conn.Open ("DSN=" & Text1.Text)
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                End Select

        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Option1(0).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            '   diable specific database
            Text2.Enabled = False
            Text2.BackColor = &H8000000F
            Command1(1).Enabled = False
            Command1(2).Enabled = False
            '   enable data source name
            Text1.Enabled = True
            Text1.BackColor = &H80000005
            Command1(0).Enabled = True
        Case 1
            '   disable data source name
            Text1.Enabled = False
            Text1.BackColor = &H8000000F
            Command1(0).Enabled = False
            '   enable specific database
            Text2.Enabled = True
            Text2.BackColor = &H80000005
            Text2.SetFocus
            Command1(1).Enabled = True
            Command1(2).Enabled = True
    End Select
End Sub
