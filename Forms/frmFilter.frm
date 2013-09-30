VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo err_filter
    Select Case Index
        Case 0
            If Len(Text1.Text) = 0 Then
                MsgBox "No filter found in the textbox", vbOKOnly + vbInformation
            Else
                Command1(1).Caption = "&Close"
                view_rs.Filter = Trim(Text1.Text)
            End If
        Case 1
            Unload Me
    End Select
    Exit Sub
err_filter:
    MsgBox "Error : Please check your syntax", vbOKOnly + vbCritical
    Command1(1).Caption = "&Cancel"
    Exit Sub
End Sub
