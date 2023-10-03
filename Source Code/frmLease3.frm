VERSION 5.00
Begin VB.Form frmLease3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5430
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmLease3.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   492
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   1332
   End
   Begin VB.TextBox txtInitialRentals 
      Height          =   372
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3000
      Width           =   492
   End
   Begin VB.TextBox txtInterest 
      Height          =   372
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2400
      Width           =   492
   End
   Begin VB.TextBox txtPeriod 
      Height          =   372
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1800
      Width           =   612
   End
   Begin VB.TextBox txtLeaseAmount 
      Height          =   372
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label lblHeading 
      Caption         =   "Lease Rental Computation"
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
      Left            =   600
      TabIndex        =   9
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblInitialRentals 
      Caption         =   "No. of Initial rentals"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
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
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblPeriod 
      Caption         =   "Period"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblLeaseAmont 
      Caption         =   "Lease Amount"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmLease3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declaring variables
    Dim rec As ADODB.Recordset
    Dim LeaseAmount As Double
    Dim Period As Integer
    Dim rate As Double
    Dim Initial As Integer
    Dim Rental As Double
    Dim VAT As Double
    Dim grossRental As Double
    Dim VATRate As Double
    Dim counter As Integer
    
    
    

Private Sub cmdCalculate_Click()
     'create the recordset and get the tax percentage
    If rec.RecordCount < 1 Then
        MsgBox " Database does not contain Tax Percentage Information..!"
        conn.Close
        Unload Me
    Else
        rec.MoveFirst
        For counter = 1 To rec.RecordCount
        If Trim(rec!TaxName) = "VAT" Then
           VATRate = CDbl(rec!percentage)
           Call calculation
           
           Exit Sub
           
        Else
        rec.MoveNext
        End If
        Next counter
                       
        MsgBox " VAT Percentage not available in Database...!"
        conn.Close
        Unload Me
        
     End If
   
   Exit Sub
   

          
        
  End Sub
  
Private Sub calculation()
    
       
    On Error GoTo errmsg
    'calculate the rental
    LeaseAmount = Val(txtLeaseAmount.Text)
    Period = Val(txtPeriod.Text)
    rate = ((Val(txtInterest.Text)) / 12) / 100
    Initial = Val(txtInitialRentals.Text)
    Rental = Round((LeaseAmount * rate) / (((Initial * rate) + 1) - ((1 + rate) ^ (Initial - Period))))
    VAT = Round(Rental * VATRate / 100)
    grossRental = Rental + VAT
    basGlobal.rentalAmt = grossRental
    frmPropEv.LeaseRent
    frmPropEv.Show
    Me.Hide
Exit Sub
'error handler
errmsg:
MsgBox Err.Description



End Sub


Private Sub Form_Load()
On Error GoTo errormsg
'open the database and reate a record set
Call basGlobal.opendb
Set rec = New ADODB.Recordset
rec.Open "select * from TaxInfo", conn, adOpenStatic, adLockOptimistic

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

Private Sub txtLeaseAmount_KeyPress(KeyAscii As Integer)

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

Private Sub txtLeaseAmount_LostFocus()
    'validating the input
    If Val(txtLeaseAmount.Text) = 0 Then
    txtLeaseAmount.SetFocus
    MsgBox "Please Enter the Lease Amount..!!"
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
    If Val(txtPeriod.Text) = 0 Then
    txtPeriod.SetFocus
    MsgBox "Invalid Period..!!"
    End If
    If Val(txtPeriod.Text) = 1 Then
    txtPeriod.SetFocus
    MsgBox "Sorry...! Facility can not be considered for one month!"
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
    End If
    If txtInterest.Text = "0" Or txtInterest.Text = "00" Then
    txtInterest.SetFocus
    MsgBox "Interest rate can not be zero..!!"
    End If

End Sub

Private Sub txtInitialRentals_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    KeyAscii = 0
    End If

End Sub

Private Sub txtInitialRentals_LostFocus()
'validating the input
    If Val(txtInitialRentals.Text) > Val(txtPeriod.Text) Then
    txtInitialRentals.SetFocus
    MsgBox ("Sorry....!!   Initial rentals can not be greater than the period..!!")
    End If
    If Val(txtInitialRentals.Text) = Val(txtPeriod.Text) Then
    txtInitialRentals.SetFocus
    MsgBox ("Sorry....!!   Initial rentals can not be equal to the period..!!")
    End If
    If Val(txtInitialRentals.Text) = 0 Then
    txtInitialRentals.SetFocus
    MsgBox "Sorry...! Atleast one rental has to be paid!"
    End If

End Sub


