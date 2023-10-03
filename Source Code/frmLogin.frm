VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3060
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5820
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":164A
   Moveable        =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   360
      Picture         =   "frmLogin.frx":2C94
      ScaleHeight     =   1995
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtUserName 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4440
      TabIndex        =   5
      Top             =   2280
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblLabels 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
   Begin VB.Label lblLabels 
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rec As ADODB.Recordset

Private Sub cmdCancel_Click()
conn.Close
Unload Me
End Sub

Private Sub cmdOK_Click()
       'checking the input
       If txtUserName.Text = "" Then
        MsgBox "Please Enter the User Name....!"
        txtUserName.SetFocus
        ElseIf txtPassword = "" Then
        MsgBox "Please Enter the Password....!"
        txtPassword.SetFocus
       
       Else
          If rec.RecordCount < 1 Then
        MsgBox " Database does not contain User Names, Please Check...!"
        conn.Close
        Unload Me
        
        Else
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
        If Trim(txtUserName.Text) = Trim(rec!UserName) Then
            If Trim(txtPassword.Text) = Decrypt(Trim(rec!Password)) Then
                Unload Me
                frmMDIForm.Show
            If rec!Flag = "0" Then
            frmMDIForm.mnuMaintenance.Visible = False
            frmMDIForm.mnuEditFacilityStatus.Visible = False
            frmMDIForm.mnuEditBlackListFlag.Visible = False
            End If
            conn.Close
            Exit Sub
            
            Else
                MsgBox "Invalid Password, try again..!"
                txtPassword.Text = ""
                txtPassword.SetFocus
            Exit Sub
       End If
       
       Else
        rec.MoveNext
        End If
        Next counter
       MsgBox "User not Fount..!"
        txtPassword.Text = ""
        txtUserName.Text = ""
        txtUserName.SetFocus
    End If
   End If
Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo errormsg
    'open database and create the record set
    Call basGlobal.opendb
    Set rec = New ADODB.Recordset
    rec.Open "select * from UserInfo", conn, adOpenStatic, adLockOptimistic
    rec.MoveFirst
Exit Sub

'error handler
errormsg:

If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Unload Me
    Exit Sub
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Unload Me
    Exit Sub
    Else
    MsgBox Err.Description
    Unload Me
    Exit Sub
    End If
End Sub

'decrypt the password
Public Function Decrypt(ByVal pwd As String) As String
  Dim Length As Integer
 Dim NewText As String
 Dim Charactor As String
 Dim i As Integer
 
 Charactor = ""
    Length = Len(pwd)
    For i = 1 To Length
        Charactor = Mid(pwd, i, 1)
        Select Case Asc(Charactor)
            Case 192 To 217     ' À - Ù ( A - Z )
                Charactor = Chr(Asc(Charactor) - 127)
            Case 218 To 243     ' Ú - ó ( a - z )
                Charactor = Chr(Asc(Charactor) - 121)
            Case 244 To 253     ' ô - ý ( 0 - 9 )
                Charactor = Chr(Asc(Charactor) - 196)
            Case 32             ' Space
                Charactor = Chr(32)
        End Select
        NewText = NewText + Charactor
    Next
    Decrypt = NewText
End Function

