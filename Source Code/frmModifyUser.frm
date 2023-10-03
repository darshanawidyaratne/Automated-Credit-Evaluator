VERSION 5.00
Begin VB.Form frmModifyUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Users"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4290
   Icon            =   "frmModifyUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4290
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFlag 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtUserName 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label lblFlag 
      BackColor       =   &H8000000A&
      Caption         =   "Flag"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblHeading 
      Caption         =   " Modify Users...!!!"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblUserName 
      Caption         =   "User Name"
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
End
Attribute VB_Name = "frmModifyUser"
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

Private Sub cmdModify_Click()
    
    If txtUserName.Text = "" Then
        MsgBox "Please Enter the User Name to Delete..!"
        txtUserName.SetFocus
        
       Else
        
       On Error GoTo errormsg
    'open datadase and create the recordset
    Set conn = New ADODB.Connection
    Set rec = New ADODB.Recordset
    connstr = "dsn=CreditInfoDSN"
    conn.Open connstr
    rec.Open "select * from UserInfo", conn, adOpenStatic, adLockOptimistic
    
   If rec.RecordCount < 1 Then
        MsgBox " Database does not contain Users Names, Please Check...!"
        conn.Close
        Unload Me
        
        Else
        
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
            If Trim(txtUserName.Text) = Trim(rec!UserName) Then
            cmdSave.Enabled = True
            cmdModify.Enabled = False
            txtUserName.Enabled = False
            txtFlag.Enabled = True
            txtFlag.Text = rec!Flag
            Exit Sub
            
               
    Else
        rec.MoveNext
        End If
        Next counter
        MsgBox " User" + " " + txtUserName.Text + " " + "Not Found...!!"
        txtUserName.SetFocus
        
        
       End If
    End If
    

Exit Sub


errormsg:
    If Err.Number = -2147467259 Then
       MsgBox " Database Server not responding, Please Check the connections / DSN..! "
       Unload Me
        ElseIf Err.Number = -2147217865 Then
       MsgBox " Table 'Supervisors' not found in database, Please Check..!"
       Unload Me
        Else
        MsgBox Err.Description
        Unload Me
    End If


End Sub
  





Private Sub cmdSave_Click()
If txtFlag.Text = "" Then
txtFlag.SetFocus
MsgBox "Please enter the Flag!"
Exit Sub
ElseIf txtFlag.Text = "0" Then
rec!Flag = Trim(txtFlag.Text)
rec.Update
MsgBox "Modified Successfully!"
Exit Sub
ElseIf txtFlag.Text = "1" Then
rec!Flag = Trim(txtFlag.Text)
rec.Update
MsgBox "Modified Successfully!"
Exit Sub
Else
txtFlag.SetFocus
MsgBox " Flag must be either 0 or 1 "
Exit Sub

End If


End Sub
