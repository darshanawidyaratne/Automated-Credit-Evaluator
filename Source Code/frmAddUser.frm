VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Users"
   ClientHeight    =   3885
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4860
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4860
   Begin VB.TextBox txtFlag 
      Height          =   285
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtReEnter 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   390
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label lblUser 
      Caption         =   "Normal User = 0"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblAmin 
      Caption         =   "Administrator  = 1"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblFlag 
      Caption         =   "Flag ( 0 OR 1 )"
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
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblReEnter 
      Caption         =   "Re Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label lblHeading 
      Caption         =   " Add Users...!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   1815
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
Attribute VB_Name = "frmAddUser"
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

Private Sub cmdAdd_Click()
    
     If txtUserName.Text = "" Then
        txtUserName.SetFocus
        MsgBox "Please Enter the User Name to Add..!"
        Exit Sub
     End If
       
     If txtPassword.Text = "" Then
        MsgBox "Please Enter the Password..!"
        txtPassword.SetFocus
        Exit Sub
     End If
       
        If txtReEnter.Text = "" Then
        txtReEnter.SetFocus
        MsgBox "Please Re-enter the Password..!"
        Exit Sub
     End If
       
     If txtFlag.Text = "" Then
        txtFlag.SetFocus
        MsgBox "Please Enter the Flag...!"
        Exit Sub
      End If
     
     If Trim(txtPassword.Text) <> Trim(txtReEnter.Text) Then
        txtPassword.Text = ""
        txtReEnter.Text = ""
        txtPassword.SetFocus
        MsgBox " Password not matched....Please enter again...!"
        Exit Sub
     Else
       'open datadase and create the recordset
        On Error GoTo errormsg
        Set conn = New ADODB.Connection
        Set rec = New ADODB.Recordset
        connstr = "dsn=CreditInfoDSN"
        conn.Open connstr
        rec.Open "select * from UserInfo", conn, adOpenStatic, adLockOptimistic
        
        If rec.RecordCount < 1 Then
        rec.AddNew
        rec!UserName = Trim(txtUserName.Text)
        rec!Password = Encrypt(Trim(txtPassword))
        rec!flag = Trim(txtFlag.Text)
        rec.Update
        MsgBox " User Successfully Added....!"
        Exit Sub
        Else
        
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
        If Trim(rec!UserName) = Trim(txtUserName.Text) Then
        MsgBox "User Name already exist...!"
        Exit Sub
        End If
        rec.MoveNext
        Next counter
                
      rec.MoveLast
        rec.AddNew
        rec!UserName = Trim(txtUserName.Text)
        rec!Password = Encrypt(Trim(txtPassword))
        rec!flag = Trim(txtFlag.Text)
        rec.Update
        MsgBox " User Successfully Added....!"
        Exit Sub
    End If
      
errormsg:
    If Err.Number = -2147467259 Then
       MsgBox " Database Server not responding, Please Check the connections / DSN..! "
       Unload Me
        ElseIf Err.Number = -2147217865 Then
       MsgBox " Table 'Users' not found in database, Please Check..!"
       Unload Me
        Else
        MsgBox Err.Description
        Unload Me
    End If

End If

End Sub
  
Public Function Encrypt(ByVal icText As String) As String
 Dim icLen As Integer
 Dim icNewText As String
 Dim icChar As String
 Dim i As Integer
  
 icChar = ""
    icLen = Len(icText)
    For i = 1 To icLen
        icChar = Mid(icText, i, 1)
        Select Case Asc(icChar)
            Case 65 To 90   ' A - Z
                icChar = Chr(Asc(icChar) + 127)
            Case 97 To 122  ' a - z
                icChar = Chr(Asc(icChar) + 121)
            Case 48 To 57   ' 0 - 9
                icChar = Chr(Asc(icChar) + 196)
            Case 32
                icChar = Chr(32)    ' Space
        End Select
        icNewText = icNewText + icChar
    Next
    Encrypt = icNewText
End Function


Private Sub xtxFlag_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 49 Then
KeyAscii = 0
End If
End Sub

