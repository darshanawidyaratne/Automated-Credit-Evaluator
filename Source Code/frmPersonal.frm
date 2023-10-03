VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Borrower"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmPersonal.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5940
   Begin VB.Frame FrameLarge 
      Height          =   6375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txtEmployer 
         Height          =   285
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtOffAddress 
         Height          =   285
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Frame FrameSmall 
         Height          =   2895
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   3975
         Begin MSComCtl2.DTPicker DTPicker 
            Height          =   255
            Left            =   2280
            TabIndex        =   5
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   38285
         End
         Begin VB.ComboBox cmbNationality 
            Height          =   315
            ItemData        =   "frmPersonal.frx":164A
            Left            =   2280
            List            =   "frmPersonal.frx":1654
            TabIndex        =   6
            Text            =   "Please Select"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTelephone 
            Height          =   285
            Left            =   2280
            MaxLength       =   14
            TabIndex        =   11
            Top             =   2400
            Width           =   1455
         End
         Begin VB.ComboBox cmbAccommodation 
            Height          =   315
            ItemData        =   "frmPersonal.frx":1675
            Left            =   2280
            List            =   "frmPersonal.frx":1682
            TabIndex        =   10
            Text            =   "Please Select"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox cmbDependants 
            Height          =   315
            ItemData        =   "frmPersonal.frx":169F
            Left            =   2280
            List            =   "frmPersonal.frx":16AF
            TabIndex        =   9
            Text            =   "Please Select"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cmbCivilStatus 
            Height          =   315
            ItemData        =   "frmPersonal.frx":16C3
            Left            =   2280
            List            =   "frmPersonal.frx":16D0
            TabIndex        =   8
            Text            =   "Please Select"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtNIC 
            Height          =   285
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTelephone 
            Caption         =   "Telephone No"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label lblAccommodation 
            Caption         =   "Accommodation"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblDependants 
            Caption         =   "No. of Dependants"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblCililStatus 
            Caption         =   "Civil Status"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblNIC 
            Caption         =   "N.I.C.No"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblNationality 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblDOB 
            Caption         =   "Date of Birth (dd/mm/yyyy)"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1935
         End
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
      Begin VB.Label lblEmployer 
         Caption         =   "Present Employer"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblOffAddress 
         Caption         =   "Office Address"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPrvAddress 
         Caption         =   "Private Address"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblFullName 
         Caption         =   "Full Name"
         Height          =   255
         Left            =   360
         TabIndex        =   14
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
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Diable keypress
Private Sub cmbAccommodation_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
'check for the input
Private Sub cmbAccommodation_LostFocus()
If cmbAccommodation.Text = "" Then
cmbAccommodation.Text = "Please Select"
End If

End Sub
'Diable keypress
Private Sub cmbCivilStatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
'check for the input
Private Sub cmbCivilStatus_LostFocus()
If cmbCivilStatus.Text = "" Then
cmbCivilStatus.Text = "Please Select"
End If
End Sub
'Diable keypress
Private Sub cmbDependants_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
'check for the input
Private Sub cmbDependants_LostFocus()
If cmbDependants.Text = "" Then
cmbDependants.Text = "Please Select"
End If

End Sub
'Diable keypress
Private Sub cmbNationality_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

'check for the input
Private Sub cmbNationality_LostFocus()
If cmbNationality.Text = "" Then
cmbNationality.Text = "Please Select"
End If

End Sub

Private Sub cmdNext_Click()
On Error GoTo errormsg
'check for the input
If txtFullName.Text = "" Or txtPrvAddress.Text = "" Or txtNIC.Text = "" Then
txtFullName.SetFocus
MsgBox "Please Enter Details...!"
Exit Sub
End If

'check for the input
If cmbAccommodation.Text = "" Or cmbCivilStatus.Text = "" Or cmbNationality.Text = "" Or cmbDependants.Text = "" Then
cmbNationality.SetFocus
MsgBox "Please Enter Details...!"
End If


'open database & create a reccordset
Call basGlobal.opendb
Set rec = New ADODB.Recordset
rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
rec.MoveFirst

For counter = 1 To rec.RecordCount
'check whether existing
If rec!vcharNIC = Trim(txtNIC.Text) Then
basGlobal.CusKey = rec!vcharCIF
basGlobal.pd = Trim(rec!Flag)
frmPersonalEv.Show
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
rec!vcharNIC = Trim(txtNIC.Text)
rec!vcharOfficeAdress = Trim(txtOffAddress.Text)
rec!vcharCIF = Trim(basGlobal.CusKey)
rec!vcharEmployer = Trim(txtEmployer.Text)
rec!dtDOB = DTPicker.Value
rec!Nationality = Trim(cmbNationality.Text)
rec!CivilStatus = Trim(cmbCivilStatus.Text)
rec!vcharTelephone = Trim(txtTelephone.Text)
rec!vcharDepen = Trim(cmbDependants.Text)
rec!vcharAccomodation = Trim(cmbAccommodation.Text)
rec!Flag = 1
rec.Update

frmPersonalEv.Show
Exit Sub

'error handler
errormsg:

If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Exit Sub
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Exit Sub
    ElseIf Err.Number = -2147352571 Then
    MsgBox "Date of Birth Not Valid....!"
    Exit Sub
    Else
    MsgBox Err.Description
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
