VERSION 5.00
Begin VB.Form frmPersonall 
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
   Begin VB.TextBox txtSpouse 
      Height          =   285
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Frame FrameLarge 
      Height          =   6375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   4560
         TabIndex        =   29
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
         Height          =   3255
         Left            =   360
         TabIndex        =   17
         Top             =   2880
         Width           =   3975
         Begin VB.ComboBox cmbNationality 
            Height          =   315
            Left            =   2280
            TabIndex        =   7
            Text            =   "Please Select"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTelephone 
            Height          =   285
            Left            =   2280
            MaxLength       =   14
            TabIndex        =   13
            Top             =   2760
            Width           =   1455
         End
         Begin VB.ComboBox cmbAccommodation 
            Height          =   315
            Left            =   2280
            TabIndex        =   12
            Text            =   "Please Select"
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtAge 
            Height          =   285
            Left            =   2280
            TabIndex        =   11
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox cmbDependants 
            Height          =   315
            Left            =   2280
            TabIndex        =   10
            Text            =   "Please Select"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cmbCivilStatus 
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Text            =   "Please Select"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtNIC 
            Height          =   285
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtDOB 
            Height          =   285
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblTelephone 
            Caption         =   "Telephone No"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label lblAccommodation 
            Caption         =   "Accommodation"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblAge 
            Caption         =   "Their Ages"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblDependants 
            Caption         =   "No. of Dependants"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblCililStatus 
            Caption         =   "Civil Status"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblNIC 
            Caption         =   "N.I.C.No"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblNationality 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblDOB 
            Caption         =   "Date of Birth (dd/mm/yyyy)"
            Height          =   255
            Left            =   120
            TabIndex        =   18
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
      Begin VB.Label lblSpouseName 
         Caption         =   "Full Name of Spouse"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblEmployer 
         Caption         =   "Present Employer"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblOffAddress 
         Caption         =   "Office Address"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPrvAddress 
         Caption         =   "Private Address"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblFullName 
         Caption         =   "Full Name"
         Height          =   255
         Left            =   360
         TabIndex        =   15
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
Attribute VB_Name = "frmPersonall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNext_Click()
On Error GoTo errormsg
    Call basGlobal.opendb
    Set rec = New ADODB.Recordset
    rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
    rec.MoveFirst
For Counter = 1 To rec.RecordCount - 1
If rec.vcharNIC = txtNIC.Text Then
Form2.Show
Else
rec.MoveNext
End If
Next Counter

rec.MoveLast
rec.AddNew
vcharName = txtFullName.Text
vcharAddress = txtPrvAddress.Text
vcharNIC = txtNIC.Text





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
