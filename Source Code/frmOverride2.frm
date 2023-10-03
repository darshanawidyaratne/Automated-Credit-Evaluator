VERSION 5.00
Begin VB.Form frmOverride2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supervisor Override"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4290
   Icon            =   "frmOverride2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4290
   Begin VB.TextBox txtSupervisorID 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblHeading 
      Caption         =   "Supervisor Override Required......!!!"
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
      TabIndex        =   6
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblSupervisorID 
      Caption         =   "Supervisor ID"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1080
   End
End
Attribute VB_Name = "frmOverride2"
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
Dim overrideStatus As Boolean



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If txtSupervisorID.Text = "" Then
        MsgBox "Please Enter the Supervisor ID....!"
        txtSupervisorID.SetFocus
        ElseIf txtPassword = "" Then
        MsgBox "Please Enter the Password....!"
        txtPassword.SetFocus
       Else
        
       On Error GoTo errormsg
    'open datadase and create the recordset
    Set conn = New ADODB.Connection
    Set rec = New ADODB.Recordset
    connstr = "dsn=CreditInfoDSN"
    conn.Open connstr
    rec.Open "select * from supervisors", conn, adOpenStatic, adLockOptimistic
    
   If rec.RecordCount < 1 Then
        MsgBox " Database does not contain Supervisor ID's, Please Check...!"
        conn.Close
        Unload Me
        
        Else
        
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
            If Trim(txtSupervisorID.Text) = Trim(rec!SupervisorID) Then
                           
                If Trim(txtPassword.Text) = Decrypt(Trim(rec!Password)) Then
                frmAddUser.Show
                Unload Me
                
                
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
        
        MsgBox "Supervisor not Fount..!"
        txtPassword.Text = ""
        txtSupervisorID.Text = ""
        txtSupervisorID.SetFocus
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

End If



End Sub
  

Private Sub txtSupervisorID_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub


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

