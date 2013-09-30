VERSION 5.00
Object = "{2B12169D-6738-11D2-BF5B-00A024982E5B}#29.2#0"; "axbutton.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About OpenBase"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   1665
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin axButtonControl.axButton axButton1 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
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
   Begin VB.Label Label1 
      Caption         =   "This product is unregistered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0442
      Height          =   855
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Version 0.1.0                 Copyright(C) 2004 nBit Enterprise"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "nBit Enterprise                       OpenBase Professional"
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
