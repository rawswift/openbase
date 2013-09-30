VERSION 5.00
Object = "{2B12169D-6738-11D2-BF5B-00A024982E5B}#29.2#0"; "axbutton.ocx"
Begin VB.Form frmConnPass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Login"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin axButtonControl.axButton axButton1 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "Admin"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmConnPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
'   if error goto err_conn
On Error GoTo err_connpass

    Select Case Index
        Case 0
            If Len(Text1.Text) <> 0 Then
                sOpenStatement = sOpenStatement & "UID=" & Trim(Text1.Text) & ";"
            End If
    
            If Len(Text2.Text) <> 0 Then
                sOpenStatement = sOpenStatement & "PWD=" & Trim(Text2.Text) & ";"
            End If
            
                    If bTest Then
                        test_conn.Open (sOpenStatement)
                        MsgBox "Test Connection Succeeded", vbOKOnly + vbInformation, "OpenBase Connection"
                        test_conn.Close
                        Set test_conn = Nothing
                    Else
                        main_conn.Open (sOpenStatement)
                    End If
                    
            Unload Me
        Case 1
            Unload Me
            bConnEstablished = False
    End Select
    
    Exit Sub
    
'   error here
err_connpass:

    MsgBox "Error : Connection failed", vbOKOnly + vbCritical, "OpenBase Connection Test"
    bConnEstablished = False
    Exit Sub
    
End Sub

