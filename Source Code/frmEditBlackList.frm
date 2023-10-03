VERSION 5.00
Begin VB.Form frmEditBlackList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Black List Flag"
   ClientHeight    =   2040
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3765
   Icon            =   "frmEditBlackList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3765
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1020
   End
   Begin VB.TextBox txtCIF 
      Height          =   348
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblHeading 
      Caption         =   " Edit Black List Flag....!!! Status.....!!"
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
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblCIF 
      Caption         =   "CIF Number."
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
End
Attribute VB_Name = "frmEditBlackList"
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





Private Sub cmdChange_Click()
On Error GoTo errormsg

If txtCIF.Text = "" Then
txtCIF.SetFocus
MsgBox "Please Enter the CIF Number"
Exit Sub
End If

        Set conn = New ADODB.Connection
        Set rec = New ADODB.Recordset
        connstr = "dsn=CreditInfoDSN"
        conn.Open connstr
        rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
        If rec.RecordCount < 1 Then
        MsgBox "Customer Not Found....!"
        Exit Sub
        End If
        
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
        If Trim(txtCIF.Text) = Trim(rec!vcharCIF) Then
                    
                    If rec!Flag = "1" Then
                    rec!Flag = "0"
                    rec.Update
                    MsgBox "Changed Successfully..!!"
                    Exit Sub
                    End If
                    
                    If rec!Flag = "0" Then
                    rec!Flag = "1"
                    rec.Update
                    MsgBox "Changed Successfully..!!"
                    Exit Sub
                    End If
                    
                    
         End If
         rec.MoveNext
         Next counter
         txtCIF.SetFocus
         MsgBox "Customer Not Found...!!"
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


Private Sub txtCIF_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub


