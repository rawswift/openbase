VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "nBit's OpenBase Professional Version 1.0"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10170
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1323
      BandCount       =   2
      _CBWidth        =   10170
      _CBHeight       =   750
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   3135
      NewRow1         =   0   'False
      Child2          =   "Toolbar2"
      MinHeight2      =   330
      Width2          =   255
      NewRow2         =   -1  'True
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   582
         ButtonWidth     =   1561
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Filter"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Export"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Access"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Advantage"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "MySQL"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Text File"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Excel File"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   582
         ButtonWidth     =   2910
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open Connection"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Field Description"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Drop Connection"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Execute Query"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6540
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu file_sub_open 
         Caption         =   "&Open Connection"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu file_sub_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_tables 
      Caption         =   "T&ables"
      Begin VB.Menu tables_sub_filter 
         Caption         =   "Filter"
         Enabled         =   0   'False
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu tables_sub_export 
         Caption         =   "Export to"
         Enabled         =   0   'False
         Begin VB.Menu export_access 
            Caption         =   "Access"
         End
         Begin VB.Menu export_advantage 
            Caption         =   "Advantage"
         End
         Begin VB.Menu export_mysql 
            Caption         =   "MySQL"
         End
         Begin VB.Menu export_text 
            Caption         =   "Text File"
         End
         Begin VB.Menu export_excel 
            Caption         =   "Excel File"
         End
      End
   End
   Begin VB.Menu menu_window 
      Caption         =   "&Window"
      Begin VB.Menu sub_window_horizontal 
         Caption         =   "T&ile Horizontally"
      End
      Begin VB.Menu sub_window_vertical 
         Caption         =   "&Tile Vertically"
      End
      Begin VB.Menu sub_window_cascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      Begin VB.Menu help_sub_openbase 
         Caption         =   "nBit's OpenBase &Help"
      End
      Begin VB.Menu help_sub_about 
         Caption         =   "&About OpenBase"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub export_access_Click()
    frmXaccess.Show vbModal
End Sub

Private Sub export_advantage_Click()
    MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
End Sub

Private Sub export_excel_Click()
    MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
End Sub

Private Sub export_mysql_Click()
    MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
End Sub

Private Sub export_text_Click()
    MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
End Sub

'   menu File -> Exit
Private Sub file_sub_exit_Click()
    '   unload this form
    Unload Me
End Sub

Private Sub file_sub_open_Click()
    Load frmOpen
    frmOpen.Show
End Sub

Private Sub tables_sub_filter_Click()
    frmFilter.Show vbModal
End Sub

'   menu Help -> About
Private Sub help_sub_about_Click()
    '   show form about
    frmAbout.Show vbModal
End Sub

Private Sub MDIForm_Load()
    '   disable view_field and drop button
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    
    '   unload splash screen
    Unload frmSplash
    
End Sub

'   prompt before exiting
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   prompt this message then get user's choice
    myResponse = MsgBox("Do you want to close this application?", vbYesNo + vbInformation, "Exit OpenBase")
        '   user's response?
        If myResponse = vbNo Then
            Cancel = True
        Else
            Unload Me
        End If
End Sub

'   close this app
Private Sub MDIForm_Unload(Cancel As Integer)
    '   call sub kill_app to close this application
    kill_app
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err_query

    Select Case Button.Index
        Case 1  '   open connection button
            '   show form Open
            Load frmOpen
            frmOpen.Show
        Case 2
            Load frmField
            frmField.Show vbModal
        Case 3  '   drop connection button
            idiscon = MsgBox("Do you want to drop the current connection?", vbYesNo + vbExclamation, "Disconnect Current Connection")
    
                If idiscon = vbYes Then
                    Unload frmView
                End If
            provChoice = 0
        Case 4  '   execute query
        
            If bQueryOpened Then
                query_rs.Close
                Set query_rs = Nothing
                bQueryOpened = False
            End If
            
                If Len(Trim(frmView.Text1.Text)) = 0 Then
                    MsgBox "No SQL statement found", vbOKOnly + vbExclamation
                Else
                                        
                '   check what provider then set cursor location and open recordset
                    Select Case provChoice
                        Case 2, 3, 4  '   msdasql or jet connection
                            query_rs.CursorLocation = adUseClient
                    End Select

                    query_rs.Open (Trim(frmView.Text1.Text)), main_conn, adOpenStatic, adLockOptimistic
                    Set frmView.DataGrid2.DataSource = query_rs
                    frmView.DataGrid2.Refresh
                    bQueryOpened = True
                End If
            
    End Select
    
    Exit Sub

err_query:

    MsgBox "ERROR : Unable to process query" & vbCrLf & "Statement : " & frmView.Text1.Text & vbCrLf _
        & "Please check syntax", vbOKOnly + vbCritical

    Exit Sub
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            frmFilter.Show vbModal
        Case 2
            If provChoice = 2 Or provChoice = 3 Then  '   2 = access
                MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
            Else
                'frmXaccess.Show vbModal
                frmXaccess.Show
            End If
    End Select
End Sub

Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1
            frmXaccess.Show vbModal
        Case 2
            MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
        Case 3
            MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
        Case 4
            MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
        Case 5
            MsgBox "Sorry but this feature is not currently available", vbOKOnly + vbInformation
    End Select
End Sub
