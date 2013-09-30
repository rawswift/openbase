VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmField 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "List of field names and their types/description"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iFieldCount As Integer
Dim iType As Integer
Dim myCounter As Integer
Dim myIndex As Integer
Dim sType As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    '   set form caption which is the currently opened table
    frmField.Caption = sTableName
    
    '   configure listview properties
    ListView1.View = lvwReport
    
    ListView1.ListItems.Clear
    
    '   get number of field in the table
    iFieldCount = view_rs.Fields.Count
    
    myCounter = 0
    myIndex = 1
    
    While myCounter < iFieldCount
        ListView1.ListItems.Add myIndex, , view_rs(myCounter).Name
        ListView1.ListItems(myIndex).Bold = True
            get_type (myIndex)
        myCounter = myCounter + 1
        myIndex = myIndex + 1
    Wend
    
    Label1(1).Caption = myCounter & " field(s) found in table " & sTableName
    
End Sub

Private Sub get_type(dIndex As Integer)
    
    iType = view_rs(myCounter).Type
    
    Select Case iType
        Case adArray
                sType = "Array"
        Case adBigInt
                sType = "Big Integer"
        Case adBinary
                sType = "Binary"
        Case adBoolean
                sType = "Boolean"
        Case adByRef    '   ???
                sType = "Pointer"
        Case adBSTR
                sType = "Null-terminated Character String (Unicode)"
        Case adChar
                sType = "Character"
        Case adCurrency
                sType = "Currency"
        Case adDate
                sType = "Date"
        Case adDBDate
                sType = "Date (yyyymmdd)"
        Case adDBTime
                sType = "Time (hhmmss)"
        Case adDBTimeStamp
                sType = "Date-Time Stamp (yyyymmddhhmmss)"
        Case adDecimal
                sType = "Decimal (Exact Numeric)"
        Case adDouble
                sType = "Double-Precision Floating Point"
        Case adEmpty
                sType = "No Value"
        Case adError
                sType = "32-bit Error Code"
        Case adGUID
                sType = "Globally Unique Identifier"
        Case adIDispatch
                sType = "Pointer to an IDispatch Interface"
        Case adInteger
                sType = "Integer (4-byte Signed Integer)"
        Case adIUnknown
                sType = "Pointer to an IUnnknown Interface"
        Case adLongVarBinary
                sType = "Long Binary"
        Case adLongVarChar
                sType = "Long String"
        Case adLongVarWChar
                sType = "Long Null-terminated String"
        Case adNumeric
                sType = "Numeric (Exact Numeric)"
        Case adSingle:
                sType = "Single-Precision Floating-Point"
        Case adSmallInt
                sType = "Small Integer (2-byte Signed Integer)"
        Case adTinyInt
                sType = "Tiny Integer (1-byte Signed Integer)"
        Case adUnsignedBigInt
                sType = "Big Integer (8-byte Unsigned Integer)"
        Case adUnsignedInt
                sType = "Unsigned Integer (4-byte Unsigned Integer)"
        Case adUnsignedSmallInt
                sType = "Unsigned Small Integer (2-byte Unsigned Integer)"
        Case adUnsignedTinyInt
                sType = "Unsigned Tiny Integer (1-byte Unsigned Integer)"
        Case adUserDefined
                sType = "User-Defined"
        Case adVarBinary
                sType = "Binary"
        Case adVarChar
                sType = "String"
        Case adVariant
                sType = "Automation Variant"
        Case advector   '   ???
                sType = "Data is a DBVECTOR Structure"
        Case adVarWChar
                sType = "Null-terminated Unicode Character String"
        Case adWChar
                sType = "Null-terminated Unicode Character String"
    End Select
        
        '   show field type
        ListView1.ListItems(dIndex).ListSubItems.Add , , sType
        
End Sub
