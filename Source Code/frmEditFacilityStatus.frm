VERSION 5.00
Begin VB.Form frmFacilityStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Facility Status"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4290
   Icon            =   "frmEditFacilityStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4290
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   390
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1020
   End
   Begin VB.TextBox txtACNO 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtFlg 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblHeading 
      Caption         =   "       Edit Facility Status....!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblACNO 
      Caption         =   "Enter A/C No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label lblFlg 
      Caption         =   "Flag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1080
   End
End
Attribute VB_Name = "frmFacilityStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare variables
Dim conn As ADODB.Connection
Dim rec As ADODB.Recordset
Dim connstr As String
Dim counter As Integer




Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If txtFlg.Text = "" Then
txtFlg.SetFocus
MsgBox "Please Enter the Flag"
Exit Sub
End If

rec!intFlag = Trim(txtFlg.Text)
rec.Update
MsgBox " Changed Successfully...!"
cmdSave.Enabled = False
txtFlg.Text = ""
txtACNO.Enabled = True
txtFlg.Enabled = False
txtACNO.Text = ""
txtACNO.SetFocus

End Sub




Private Sub cmdSearch_Click()
 On Error GoTo errormsg
 
 If txtACNO.Text = "" Then
 txtACNO.SetFocus
 MsgBox "Please Enter the Account Number...!"
 Exit Sub
 End If
         
        Set conn = New ADODB.Connection
        Set rec = New ADODB.Recordset
        connstr = "dsn=CreditInfoDSN"
        conn.Open connstr
        rec.Open "select * from facilities", conn, adOpenStatic, adLockOptimistic
 
 
 
 rec.MoveFirst
 For counter = 1 To rec.RecordCount
 If Trim(txtACNO.Text) = rec!vcharAccNo Then
 txtFlg.Enabled = True
 txtFlg.Text = rec!intFlag
 cmdSave.Enabled = True
 txtACNO.Enabled = False
 Exit Sub
 End If
 rec.MoveNext
 Next counter
 MsgBox "Account Not Found....!"
 
    
           

Exit Sub


errormsg:
    If Err.Number = -2147467259 Then
       MsgBox " Database Server not responding, Please Check the connections / DSN..! "
       Unload Me
        ElseIf Err.Number = -2147217865 Then
       MsgBox " Table not found in database, Please Check..!"
       Unload Me
        Else
        MsgBox Err.Description
        Unload Me
    End If





End Sub
  

Private Sub txtACNO_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub


Private Sub txtFlg_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 51 Then
KeyAscii = 0
End If
End Sub
