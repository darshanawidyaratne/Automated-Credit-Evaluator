VERSION 5.00
Begin VB.Form frmPersonalEv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Borrower"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmPersonalEv.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   5970
   Begin VB.Frame Frame2 
      Caption         =   "Evaluation Results"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   240
      TabIndex        =   27
      Top             =   5160
      Width           =   5415
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   3000
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   4200
         TabIndex        =   32
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblReason2 
         Height          =   615
         Left            =   1320
         TabIndex        =   31
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label lblResult2 
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Comment :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblResult 
         Caption         =   "Result :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbPurpose 
      Height          =   315
      ItemData        =   "frmPersonalEv.frx":164A
      Left            =   2040
      List            =   "frmPersonalEv.frx":1654
      TabIndex        =   2
      Text            =   "Please Select"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame FrameLarge 
      Height          =   4935
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtInterest 
         Height          =   285
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox cmbInfluence 
         Height          =   315
         ItemData        =   "frmPersonalEv.frx":1674
         Left            =   1920
         List            =   "frmPersonalEv.frx":1681
         TabIndex        =   12
         Text            =   "Please Select"
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox txtAge 
         Height          =   255
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbSecurity 
         Height          =   315
         ItemData        =   "frmPersonalEv.frx":16B8
         Left            =   1920
         List            =   "frmPersonalEv.frx":16D4
         TabIndex        =   11
         Text            =   "Please Select"
         Top             =   3960
         Width           =   3015
      End
      Begin VB.ComboBox cmbCharacter 
         Height          =   315
         ItemData        =   "frmPersonalEv.frx":17A2
         Left            =   1920
         List            =   "frmPersonalEv.frx":17AF
         TabIndex        =   10
         Text            =   "Please Select"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtIncome 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   9
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox cmbRepayment 
         Height          =   315
         ItemData        =   "frmPersonalEv.frx":17F0
         Left            =   1920
         List            =   "frmPersonalEv.frx":1800
         TabIndex        =   8
         Text            =   "Please Select"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.ComboBox cmbExchange 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPersonalEv.frx":184D
         Left            =   1920
         List            =   "frmPersonalEv.frx":1857
         TabIndex        =   7
         Text            =   "Please Select"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Text            =   "Please Select"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton cmdEvaluate 
         Caption         =   "Evaluate"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtPeriod 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest Rate"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Influence 
         Caption         =   "Influence Range"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label lblAge 
         Caption         =   "Age (in years)"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblSecurity 
         Caption         =   "Security"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblCharacter 
         Caption         =   "Character"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblIncome 
         Caption         =   "Monthly Income"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblRepayment 
         Caption         =   "Repayment Source"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblExch 
         Caption         =   "Exchange Regulns"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lbPeriod 
         Caption         =   "Period (in months)"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPurpose 
         Caption         =   "Purpose"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblType 
         Caption         =   "Facility Type"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
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
Attribute VB_Name = "frmPersonalEv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'validating inputs
Private Sub cmbType_Click()
If Trim(cmbType.Text) = "Letter of Credit" Then
cmbExchange.Enabled = True
Else
cmbExchange.Enabled = False
End If
End Sub
'validating inputs
Private Sub cmbType_GotFocus()
txtInterest.Enabled = True
txtPeriod.Enabled = True
cmbSecurity.Enabled = True
End Sub

Private Sub cmdPrint_Click()
frmPersonal.PrintForm
Me.PrintForm
End Sub

Private Sub cmdSave_Click()
If lblResult2.Caption = "" Or lblReason2.Caption = "" Then
MsgBox "Please Evaluate First...!"
Exit Sub
End If

Set evrec2 = New ADODB.Recordset
evrec2.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
evrec2.MoveFirst
Dim counterrr As Integer
For counterrr = 1 To evrec2.RecordCount
If Trim(evrec2!vcharProductName) = Trim(cmbType.Text) Then
basGlobal.proid = Trim(evrec2!vcharProductID)
End If
evrec2.MoveNext
Next counterrr

Set recEvl = New ADODB.Recordset
recEvl.Open "select * from CreditInfo", conn, adOpenStatic, adLockOptimistic
recEvl.MoveLast
recEvl.AddNew
recEvl!vcharCIF = Trim(basGlobal.CusKey)
recEvl!dtDateVisited = Date
recEvl!vcharFacilityCode = Trim(basGlobal.proid)
recEvl!vcharAmount = Trim(txtAmount.Text)
recEvl!vcharResult = Trim(lblResult2.Caption)
recEvl!vcharComments = Trim(lblReason2.Caption)
recEvl!vcharAttempts = Str(Val(recEvl!vcharAttempts) + 1)
recEvl.Update
MsgBox " Saved Successfully...!"
End Sub

Private Sub Form_Load()
On Error GoTo errormsg
    'open database and create the record set
    Call basGlobal.opendb
    Set rec = New ADODB.Recordset
    rec.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
    rec.MoveFirst
If rec.RecordCount < 1 Then
MsgBox "Facility Types not available in the database...!!"
Exit Sub
End If

'adding facilities to combo
For counter = 1 To rec.RecordCount
If Trim(rec!personal) = 1 Then
cmbType.AddItem (rec!vcharProductName)
End If
rec.MoveNext
Next counter

Set rec2 = New ADODB.Recordset
rec2.Open "select * from SBL", conn, adOpenStatic, adLockOptimistic
'Initializing Single Borrower Limit
basGlobal.SBL = Val(rec2!SBL)
'close the database
 conn.Close

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

'start evaluation
Private Sub cmdEvaluate_Click()
 'validating inputs
 If cmbType.Text = "Please Select" Or cmbPurpose.Text = "Please Select" Or cmbSecurity.Text = "Please Select" Or cmbRepayment.Text = "Please Select" Or cmbCharacter.Text = "Please Select" Or cmbInfluence.Text = "Please Select" Or txtAge.Text = "" Or txtAmount.Text = "" Or txtIncome.Text = "" Or txtInterest.Text = "" Or txtPeriod.Text = "" Then
 MsgBox "Please Fill the Details..!"
 Exit Sub
 End If
 'validate the interest rate
 If Val(txtInterest.Text) > 50 Or Val(txtInterest.Text) = 0 Then
 MsgBox "Interest rate is not valid...!"
 Exit Sub
 End If
 
 Call basGlobal.opendb
 Set evrec = New ADODB.Recordset
 evrec.Open "select * from Facilities", conn, adOpenStatic, adLockOptimistic
 evrec.MoveFirst
 Dim counterr As Integer
  For counterr = 1 To evrec.RecordCount
  If Trim(evrec!vcharCIF) = Trim(basGlobal.CusKey) Then
  basGlobal.arr = evrec!intFlag
  End If
  evrec.MoveNext
  Next counterr
  
  If basGlobal.pd = "0" Then
  basGlobal.Result = "Rejected"
  basGlobal.Comment = " Black Listed Customer"
  lblResult2.Caption = basGlobal.Result
  lblReason2.Caption = basGlobal.Comment
    
  Exit Sub
  ElseIf basGlobal.arr = "1" Then
  basGlobal.Result = "Rejected"
  basGlobal.Comment = "Customer is having arrears facilities"
  lblResult2.Caption = basGlobal.Result
  lblReason2.Caption = basGlobal.Comment
  Exit Sub
  ElseIf basGlobal.arr = "2" Then
  basGlobal.Result = "Rejected"
  basGlobal.Comment = "Customer is in irregular"
  lblResult2.Caption = basGlobal.Result
  lblReason2.Caption = basGlobal.Comment
  Exit Sub
  End If
  
 'selecting the input
  Select Case (cmbType.Text)
 
 'regular overdraft
 Case ("ROD")
 Call ROD
 'Guarantees
 Case ("Guarantees")
 Call Guarantees
 'credit cards
 Case ("Credit Cards")
 Call CreditCards
 'LBP
 Case ("Local Bill Purchase")
 Call LocalBillPurchase
 'Temporary OD
 Case ("TOD")
 Call TOD
 'term loan
 Case ("Term Loan")
 Call TermLoan
 'personal loan
 Case ("Pahana")
 Call Pahana
 'housing loan
 Case ("Nivahana")
 Call Nivahana
 'lease
 Case ("Lease")
 Call Lease
 'import loan
 Case ("IDL")
 Call IDL
 'preshipments
 Case ("Preshipment Loan")
 Call PreshipmentLoan
 'LC
 Case ("Letter of Credit")
 Call LetterofCredit
 End Select
End Sub
'Evaluating ROD
Private Sub ROD()
Dim msg As VbMsgBoxResult
msg = MsgBox("Are the account turnovers satisfactory?", vbYesNo, "Credit Evaluator")
If msg = 7 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No sufficient turnovers"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment
Exit Sub
End If

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Monthly Salary" And Val(txtIncome.Text) < 20000 Then
basGlobal.Result = "Not suggested, but may consider depending on the need"
basGlobal.Comment = " Not for a working capital requirement.!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Bussiness" And Val(txtIncome.Text) < 30000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Bussiness not sound, Income not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Bussiness" And Val(txtIncome.Text) < 40000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Bussiness not sound, Income not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Backgound not clear, identify the activities of the borrower "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is not sound "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Overdraft can not be considered for a Personal Guarantee"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Better go for a Lease..!! Security is not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 250000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub
'evaluating g'tees
Private Sub Guarantees()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(Val(txtIncome.Text)) < 30000 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(Val(txtIncome.Text)) < 20000 And Trim(cmbInfluence.Text) = "Highly Influential" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" And Trim(txtAmount.Text) > 20000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A realistic security must be obtained"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Such a security is not applicable for a guarantee"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If


End Sub
'evaluating credit cards
Private Sub CreditCards()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > 1000000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Maximum limit allowed is Rs.1,000,000/-"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting "
basGlobal.Comment = " Age is above 75 years, risky"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A card can not be issued to a pensioner "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtIncome.Text) < 10000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Capacity of the borrower is not sound"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (Val(txtIncome.Text) * 150 / 100) < Val(txtAmount.Text) And Trim(cmbInfluence.Text) = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is not sound "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating LBP
Private Sub LocalBillPurchase()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(Val(txtIncome.Text)) < 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is not sound "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, A guarantee must be obtained "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Such security is not valid for LBP"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, But obtain a personal guarantee"
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips, But obtain a personal guarantee"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub
'evaluating TOD
Sub TOD()

Dim msg As VbMsgBoxResult
msg = MsgBox("Are the account turnovers satisfactory?", vbYesNo, "Credit Evaluator")
If msg = 7 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No sufficient turnovers"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment
Exit Sub
End If

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 6 Or Val(txtPeriod.Text) < 1 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 79 And Trim(txtAmount.Text) > 500000 Then
basGlobal.Result = "Rejected,"
basGlobal.Comment = " Age is above 79 years, it is fairly risky since the amount is also high"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbRepayment.Text) = "Monthly Salary" And Val(txtIncome.Text) < 20000 Then
basGlobal.Result = "Not suggested, but may consider depending on the need"
basGlobal.Comment = " Not for a working capital requirement.!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Bussiness" And Val(txtIncome.Text) < 30000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Bussiness not sound, Income not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Bussiness" And Val(txtIncome.Text) < 40000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Bussiness not sound, Income not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" And Val(txtAmount.Text) > 300000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Amount is high when compared with the security"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Such a security is not realistic for a temporary facility. Proposed to submit a personal guarantee"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Property Mortgage (Acceptable Property)" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Such a security is not realistic for a temporary facility. Proposed to submit a personal guarantee or grant if leeway available"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Machinery Mortgage)" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Such a security is not realistic for a temporary facility. Proposed to submit a personal guarantee or grant if leeway available"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating term loan
Private Sub TermLoan()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 48 Or Val(txtPeriod.Text) < 12 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" And Val(txtAmount.Text) > 300000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Amount is high when compared with the security"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Better go for a Lease..!! Security is not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub
'evaluating personal loan
Private Sub Pahana()
If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > 750000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the maximum amount under the particuar product"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 48 Or Val(txtPeriod.Text) < 12 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) / Val(txtPeriod.Text)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Trim(cmbCharacter.Text) = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, obtain two guarantors atleast! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Can consider with two guarantors "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Can consider with two guarantors "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Better go for a Lease..!! Security is not sufficient"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Property Mortgage (Acceptable Property)" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " but it is proposed to go for a term loan or obtain a guarantor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Machinery Mortgage" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " but it is proposed to go for a term loan or obtain a guarantor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Stock Mortgage" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " but it is proposed to go for a term loan or obtain a guarantor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" And Trim(cmbCharacter.Text) = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "It is too risky to lend for an unknown customer only with a guarantee"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating housing loan
Private Sub Nivahana()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 300 Or Val(txtPeriod.Text) < 60 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 50 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "It is mandatory that the borrower should be below 50 years "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (Val(txtAge.Text) + (Val(txtPeriod.Text)) / 12) > 55 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Combination of age and the period is not acceptable "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)) / (((1 + ((Val(txtInterest.Text) / 100) / 12)) ^ (Val(txtPeriod.Text)) - 1)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And Trim(cmbInfluence.Text) = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - ((Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)) / (((1 + ((Val(txtInterest.Text) / 100) / 12)) ^ (Val(txtPeriod.Text)) - 1)) + (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12)))) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 1000000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor.Though influential,amount is high..! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Nivahana can not be considered without a security..! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Machinery Mortgage" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A property mortgage is essential..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Stock Mortgage" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "A property mortgage is essential..! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "A property mortgage is essential..! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A property mortgage is essential..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "A property mortgage is essential..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "A property mortgage is essential..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "A property mortgage is essential..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating lease
Sub LeaseRent()
If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 60 Or Val(txtPeriod.Text) < 12 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influential or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - basGlobal.rentalAmt) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - basGlobal.rentalAmt) / Val(txtIncome.Text)) * 100) < 40 And Val(txtAmount.Text) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " It is essential to get the absolute ownership over the asset "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Allowed"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


Else

basGlobal.Result = " Allowed subject to CRIB clearence and obtaining absolute ownership over the asset "
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub
'computing lease rental
Private Sub Lease()
frmLease2.Show
Me.Hide
End Sub

'evaluating IDL
Private Sub IDL()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 6 Or Val(txtPeriod.Text) < 1 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting if the customer is not involved in business"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Monthly Salary" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Borrower should be in bussiness "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment


ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A realistic security must be obtained"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Security is in the natue of a lease facility. Can't accept"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating preshipment
Private Sub PreshipmentLoan()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAmount.Text) > basGlobal.SBL Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Request exceeds the Single Borrower Limit"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtPeriod.Text) > 6 Or Val(txtPeriod.Text) < 1 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Period out of the range"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting if the customer is not involved in business"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) < 40 And cmbInfluence.Text = "Average Customer" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf (((Val(txtIncome.Text) - (Val(txtAmount.Text) * ((Val(txtInterest.Text) / 100) / 12))) / Val(txtIncome.Text)) * 100) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Repayment Capacity is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Monthly Salary" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Borrower should be in bussiness "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And cmbCharacter.Text = "Can't Judge" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential, risky to lend if can't judge "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And cmbInfluence.Text = "Average Customer" And Val(txtAmount.Text) > 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Not influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Highly Influential" And Val(Trim(txtAmount.Text)) > 200000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" And Trim(cmbInfluence.Text) = "Influential" And Val(txtAmount.Text) > 100000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " No Security, Amount is high though influential "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A realistic security must be obtained"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Security is in the natue of a lease facility. Can't accept"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Life Insurance Policy " And Trim(Val(txtAmount.Text)) > 300000 Then
basGlobal.Result = "Not highly recommended, "
basGlobal.Comment = " It is risky if the surrender value is not realistic..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" And Trim(Val(txtAmount.Text)) > 500000 Then
basGlobal.Result = "Not highly recommended, though blue chips"
basGlobal.Comment = " If the share comes down due to any reason, loss is high!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'evaluating LC
Private Sub LetterofCredit()

If cmbPurpose.Text = "Not Acceptable" Then

basGlobal.Result = "Rejected"
basGlobal.Comment = " Purpose Not Acceptable"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(Val(cmbExchange.Text)) = "Violating Regulations" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Violating Regulations"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) < 18 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Age is below 18 years,Still a minor"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Val(txtAge.Text) > 75 Then
basGlobal.Result = "Avoid granting whether influencial or not"
basGlobal.Comment = " Age is above 75 years, A high risk involved"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Purely on Pension" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Pension can not be accepted as a repayment source "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbCharacter.Text) = "Defaulter in an Other Bank" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " A defaulter in an other bank. Risky..!! "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(Val(txtIncome.Text)) < 50000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbRepayment.Text) = "Monthly Salary" And Trim(Val(txtIncome.Text)) < 60000 Then
basGlobal.Result = "Rejected"
basGlobal.Comment = " Capacity of the borrower is poor "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "No Security" Then
basGlobal.Result = "Rejected, but can consider with the title over goods imported"
basGlobal.Comment = "No Security, Better collect the margin in order to consider  "
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Personal Guarantee" Or Trim(cmbSecurity.Text) = "Property Mortgage (Acceptable Property)" Or Trim(cmbSecurity.Text) = "Machinery Mortgage" Or Trim(cmbSecurity.Text) = "Stock Mortgage" Or Trim(cmbSecurity.Text) = "Shares (Blue Chip Company)" Or Trim(cmbSecurity.Text) = "Life Insurance Policy" Then
basGlobal.Result = "Allowed subject to CRIB Clearence"
basGlobal.Comment = " Title over goods imported should be obtained"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

ElseIf Trim(cmbSecurity.Text) = "Absolute Ownership over the assests" Then
basGlobal.Result = "Rejected"
basGlobal.Comment = "Such a security is not applicable for a guarantee"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

Else

basGlobal.Result = " Allowed subject to CRIB clearence"
basGlobal.Comment = "Evaluation Satisfactory..!"
lblResult2.Caption = basGlobal.Result
lblReason2.Caption = basGlobal.Comment

End If

End Sub

'disabling unwanted selections
Private Sub cmbType_LostFocus()

If Trim(cmbType.Text) = "" Then
cmbType.Text = "Please Select"
End If

If Trim(cmbType.Text) = "Local Bill Purchase" Then
txtInterest.Text = "11"
txtPeriod.Text = "11"
txtInterest.Enabled = False
txtPeriod.Enabled = False
End If


If Trim(cmbType.Text) = "Letter of Credit" Then
cmbExchange.Enabled = True
txtInterest.Text = "11"
txtPeriod.Text = "11"
txtInterest.Enabled = False
txtPeriod.Enabled = False
End If

If Trim(cmbType.Text) = "Credit Cards" Then

txtInterest.Text = "11"
txtPeriod.Text = "11"
cmbSecurity.Text = "111111111"

txtInterest.Enabled = False
txtPeriod.Enabled = False
cmbSecurity.Enabled = False

End If

If Trim(cmbType.Text) = "ROD" Then
txtPeriod.Text = "11"
txtPeriod.Enabled = False
End If

If Trim(cmbType.Text) = "Guarantees" Then
txtInterest.Text = "11"
txtPeriod.Text = "11"
txtInterest.Enabled = False
txtPeriod.Enabled = False
End If


End Sub

Private Sub cmbPurpose_LostFocus()
If cmbPurpose.Text = "" Then
cmbPurpose.Text = "Please Select"
End If
End Sub

Private Sub cmbExchange_LostFocus()
If cmbExchange.Text = "" Then
cmbExchange.Text = "Please Select"
End If
End Sub

Private Sub cmbRepayment_LostFocus()
If cmbRepayment.Text = "" Then
cmbRepayment.Text = "Please Select"
End If
End Sub

Private Sub cmbCharacter_LostFocus()
If cmbCharacter.Text = "" Then
cmbCharacter.Text = "Please Select"
End If
End Sub

Private Sub cmbSecurity_LostFocus()
If cmbSecurity.Text = "" Then
cmbSecurity.Text = "Please Select"
End If
End Sub

Private Sub cmbInfluence_LostFocus()
If cmbInfluence.Text = "" Then
cmbInfluence.Text = "Please Select"
End If
End Sub

'disabling keypress
Private Sub cmbType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbPurpose_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbExchange_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbRepayment_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbCharacter_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbSecurity_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbInfluence_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
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

Private Sub txtIncome_KeyPress(KeyAscii As Integer)
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

Private Sub txtPeriod_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If
End Sub

