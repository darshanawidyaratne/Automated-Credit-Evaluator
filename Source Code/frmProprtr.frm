VERSION 5.00
Begin VB.Form frmProprtr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proprietorship"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmProprtr.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5940
   Begin VB.Frame FrameLarge 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtTelephone 
         Height          =   285
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtPrvAddress 
         Height          =   285
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblTelephone 
         Caption         =   "Telephone No"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblPrvAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblFullName 
         Caption         =   " Name"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Caption         =   "   Proprietorship Evaluation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmProprtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNext_Click()
On Error GoTo errormsg
'check for the input
If txtFullName.Text = "" Or txtPrvAddress.Text = "" Then
txtFullName.SetFocus
MsgBox "Please Enter Details...!"
Exit Sub
End If


'open database & create a reccordset
Call basGlobal.opendb
Set rec = New ADODB.Recordset
rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
rec.MoveFirst
basGlobal.CusKey = ""
For counter = 1 To rec.RecordCount
'check whether existing
If rec!vcharName = Trim(txtFullName.Text) Then
basGlobal.CusKey = rec!vcharCIF
basGlobal.pd = Trim(rec!Flag)
frmPropEv.Show
Exit Sub
Else
rec.MoveNext
End If
Next counter

'not existing
'checking for highest CIF
Dim temp As Integer
rec.MoveFirst
For temp = 1 To rec.RecordCount
If Val(rec!vcharCIF) > Val(basGlobal.CusKey) Then
basGlobal.CusKey = rec!vcharCIF

End If
rec.MoveNext
Next temp
'Assigning a CIF value
basGlobal.CusKey = Str(Val(basGlobal.CusKey) + 10)
basGlobal.pd = "1"
'writing to database
rec.MoveLast
rec.AddNew
rec!vcharName = Trim(txtFullName.Text)
rec!vcharAddress = Trim(txtPrvAddress.Text)
rec!vcharCIF = Trim(basGlobal.CusKey)
rec!vcharTelephone = Trim(txtTelephone.Text)
rec!Flag = 1
rec.Update

frmPropEv.Show
Exit Sub

'error handler
errormsg:

If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Exit Sub
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Exit Sub
    Else
    MsgBox Err.Description
    Exit Sub
    End If



End Sub


Private Sub txtFullName_KeyPress(KeyAscii As Integer)

If KeyAscii < 65 Or KeyAscii > 122 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    ElseIf KeyAscii = 46 Then
    KeyAscii = 46
    Exit Sub
    ElseIf KeyAscii = 32 Then
    KeyAscii = 32
    Exit Sub
    End If
    If KeyAscii = 91 Then
    KeyAscii = 0
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub



Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    Exit Sub
    End If
End Sub

