VERSION 5.00
Begin VB.Form frmSearch2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Borrower"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   FillColor       =   &H00FFFFFF&
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5940
   Begin VB.Frame FrameLarge 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.C.No"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblFullName 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Caption         =   "Personal Borrower Evaluation"
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
Attribute VB_Name = "frmSearch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbAccommodation_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbCivilStatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbDependants_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbNationality_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdNext_Click()
On Error GoTo errormsg

If txtNIC.Text = "" Then
txtNIC.SetFocus
MsgBox ("Please Enter the NIC Number!!")
Exit Sub
End If

 
   
    Call basGlobal.opendb
    Set rec = New ADODB.Recordset
    rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
    rec.MoveFirst
For counter = 1 To rec.RecordCount - 1
If rec!vcharNIC = txtNIC.Text Then
Form2.Show
Exit Sub
Else
rec.MoveNext
End If
Next counter
rec.MoveLast
rec.AddNew
rec!vcharName = txtFullName.Text
rec!vcharAddress = txtPrvAddress.Text
rec!vcharNIC = txtNIC.Text
rec!vcharOfficeAdress = txtOffAddress.Text
rec.Update
form1.Show
Exit Sub

errormsg:

If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    'Unload Me
    Exit Sub
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    'Unload Me
    Exit Sub
    Else
    MsgBox Err.Description
    'Unload Me
    Exit Sub
    End If



End Sub








Private Sub txtDOB_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    ElseIf KeyAscii = 47 Then
    KeyAscii = 47
    Exit Sub
    End If
    KeyAscii = 0
    Exit Sub
    End If
End Sub

Private Sub txtEmployer_KeyPress(KeyAscii As Integer)
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



Private Sub txtNIC_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    ElseIf KeyAscii = 86 Then
    KeyAscii = 86
    Exit Sub
    ElseIf KeyAscii = 88 Then
    KeyAscii = 88
    Exit Sub
    End If
    KeyAscii = 0
    Exit Sub
    End If
End Sub







Private Sub txtSpouse_KeyPress(KeyAscii As Integer)
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
