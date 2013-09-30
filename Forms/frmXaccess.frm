VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmXaccess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export to Access"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Start  >>"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destination"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Specific Database"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ODBC Data Source Name"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmXaccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDBPath As String

Private Sub Command1_Click()
            CommonDialog1.Filter = "Microsoft Access (*.mdb)|*.mdb"
            '   show dialog open
            CommonDialog1.ShowOpen
            '   get selected filename
            sDBPath = CommonDialog1.FileName
            '   show selected filename
            Text2.Text = sDBPath
            Command2(0).Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo err_x_conn
    Select Case Index
        Case 0  '   start >>
            If Option1(0) Then
                dest_conn.Open ("PROVIDER=MSDASQL;DSN=" & Trim(Text1.Text) & ";")
                bDestOpened = True
            ElseIf Option1(1) Then
                dest_conn.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Text2.Text)
                bDestOpened = True
            End If
            
            '   construct SQL statement to create table
            construct_statement
            '   create the table
            create_table
            
            '   hide forms
            frmXaccess.Visible = False
            frmMain.Visible = False
            
            '   transfer records
            transfer
            
            '   restore form(s)
            frmMain.Visible = True
            
        Case 1  '   cancel
            If bDestOpened Then
                dest_conn.Close
                Set dest_conn = Nothing
            End If
            Unload Me
    End Select
    Exit Sub
err_x_conn:
    MsgBox "Unexpected error occured", vbOKOnly + vbCritical
    bDestOpened = False
    Exit Sub
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            '   diable specific database
            Text2.Enabled = False
            Text2.BackColor = &H8000000F
            Command1.Enabled = False
            '   enable data source name
            Text1.Enabled = True
            Text1.BackColor = &H80000005
            Text1.SetFocus
        Case 1
            '   disable data source name
            Text1.Enabled = False
            Text1.BackColor = &H8000000F
            Command1.Enabled = True
            '   enable specific database
            Text2.Enabled = True
            Text2.BackColor = &H80000005
            Text2.SetFocus
    End Select
End Sub

Private Sub Text1_Change()
    Command2(0).Enabled = True
End Sub
