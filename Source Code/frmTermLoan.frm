VERSION 5.00
Begin VB.Form frmTermLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Term  Loan Installement Computation "
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmTermLoan.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6255
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   492
      Left            =   4560
      TabIndex        =   10
      Top             =   2880
      Width           =   1332
   End
   Begin VB.TextBox txtInstalment 
      Enabled         =   0   'False
      Height          =   372
      Left            =   2640
      TabIndex        =   8
      Top             =   2880
      Width           =   1692
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   492
      Left            =   4560
      TabIndex        =   7
      Top             =   2160
      Width           =   1332
   End
   Begin VB.TextBox txtInterest 
      Height          =   372
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtPeriod 
      Height          =   372
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtLoanAmount 
      Height          =   372
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label lblHeading 
      Caption         =   "    Term Loan Installement Computation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblInstalment 
      Caption         =   "Installment"
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
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblInterest 
      Caption         =   "Interest Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblPeriod 
      Caption         =   "Period (in months)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblLoanAmont 
      Caption         =   "Loan Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmTermLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     
'declaring variables
    Dim LoanAmount As Double
    Dim Period As Integer
    Dim rate As Double
    Dim Instalment As Double
        
     

Private Sub cmdCalculate_Click()
        
Call calculation
     
End Sub
  
Private Sub calculation()
    
       
    On Error GoTo errmsg
    'calculate the instalment
    LoanAmount = Val(txtLoanAmount.Text)
    Period = Val(txtPeriod.Text)
    rate = ((Val(txtInterest.Text)) / 12) / 100
    Instalment = ((LoanAmount / Period) + (LoanAmount * rate))
    txtInstalment.Text = Round(Val(Instalment))

Exit Sub
'error handler
errmsg:
MsgBox Err.Description

End Sub

Private Sub cmdClear_Click()
    'clear text boxes
    txtInstalment.Text = ""
    txtLoanAmount.Text = ""
    txtPeriod.Text = ""
    txtInterest.Text = ""
    txtLoanAmount.SetFocus

End Sub

Private Sub txtLoanAmount_KeyPress(KeyAscii As Integer)

    If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    If KeyAscii = 46 Then
    KeyAscii = 46
    Exit Sub
    End If
    KeyAscii = 0
    End If

End Sub

Private Sub txtLoanAmount_LostFocus()
    'validating the input
    If Val(txtLoanAmount.Text) = 0 Then
    txtLoanAmount.SetFocus
    MsgBox "Please Enter the Loan Amount..!!"
    End If

End Sub



Private Sub txtPeriod_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If

End Sub

Private Sub txtPeriod_LostFocus()
    'validating the input
    If txtPeriod.Text = "" Or Val(txtPeriod.Text) = 0 Then
    txtPeriod.SetFocus
    MsgBox "Invalid Period..!!"
    End If
    
End Sub

Private Sub txtInterest_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    If KeyAscii = 46 Then
    KeyAscii = 46
    Exit Sub
    End If
    KeyAscii = 0
    End If

End Sub

Private Sub txtInterest_LostFocus()
    'validating the input
    If txtInterest.Text = "" Then
    txtInterest.SetFocus
    MsgBox "Interest rate can not be blank..!!"
    Exit Sub
    End If
        
    If Val(txtInterest.Text) = 0 Then
    txtInterest.SetFocus
    MsgBox "Interest rate can not be zero..!!"
    End If

End Sub




