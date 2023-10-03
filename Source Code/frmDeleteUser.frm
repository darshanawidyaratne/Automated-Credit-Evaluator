VERSION 5.00
Begin VB.Form frmDeleteUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Users"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4290
   Icon            =   "frmDeleteUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4290
   Begin VB.TextBox txtUserName 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
   Begin VB.Label lblHeading 
      Caption         =   " Delete Users...!!!"
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
Attribute VB_Name = "frmDeleteUser"
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

Private Sub cmdDelete_Click()
    
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
            rec.Delete
            rec.Update
            MsgBox " User" + " " + txtUserName.Text + " " + "Deleted Successfully....!"
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
  





