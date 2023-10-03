VERSION 5.00
Begin VB.Form frmInputNewAdvances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input New Advances"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmInputNewAdvances.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5940
   Begin VB.TextBox txtNIC 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame FrameLarge 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtACNO 
         Height          =   285
         Left            =   3960
         MaxLength       =   13
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox cmbFacility 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "Please Select"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtCIF 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   7
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1320
         MaxLength       =   70
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblACNO 
         Caption         =   "A/C No"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblNIC 
         Caption         =   "NIC"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblCIF 
         Caption         =   "CIF No"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblFacility 
         Caption         =   "Facility "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblHeading 
         Caption         =   "       Input New Advances"
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
Attribute VB_Name = "frmInputNewAdvances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFacility_KeyPress(KeyAscii As Integer)

KeyAscii = 0
End Sub

Private Sub cmbFacility_LostFocus()
If cmbFacility.Text = "" Then
cmbFacility.Text = "Please Select"
End If

End Sub

Private Sub cmdClear_Click()
   'Clear the Form
    
    txtCIF.Text = ""
    txtAddress.Text = ""
    txtName.Text = ""
    txtNIC.Text = ""
    txtAmount.Text = ""
    cmbFacility.Text = "Please Select"
    txtACNO.Text = ""
    
End Sub

Private Sub cmdSearch_Click()

On Error GoTo errormsg
    
    
'error handler
errormsg:
'server not responding
If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Exit Sub
    'table not found
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Exit Sub
    
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If


End Sub



Private Sub Text1_Change()

End Sub

Private Sub cmdUpdate_Click()

If txtCIF.Text = "" Or txtACNO.Text = "" Or txtAddress.Text = "" Or txtName.Text = "" Or txtAmount.Text = "" Or cmbFacility.Text = "Please Select" Then
txtCIF.SetFocus
MsgBox "Please Enter the details...!!"
Exit Sub
End If

'open the database and record sets
 Call basGlobal.opendb
 Set rec = New ADODB.Recordset
 Set rec2 = New ADODB.Recordset
 Set rec3 = New ADODB.Recordset
 rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
  
  rec.MoveFirst
  For counter = 1 To rec.RecordCount
  If Trim(txtCIF.Text) = Trim(rec!vcharCIF) Then
        rec2.Open "select * from Facilities", conn, adOpenStatic, adLockOptimistic
        rec3.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
        rec2.MoveLast
        rec2.AddNew
        rec2!vcharCIF = Trim(txtCIF.Text)
        rec2!vcharOutstanding = Trim(txtAmount.Text)
        rec2!vcharAccNo = Trim(txtACNO.Text)
        rec3.MoveFirst
        Dim counter2 As Integer
        For counter2 = 1 To rec3.RecordCount
        If Trim(cmbFacility.Text) = Trim(rec3!vcharProductName) Then
        rec2!vchaFacilityCode = rec3!vcharProductID
        rec2!intFlag = 1
        rec2.Update
        MsgBox "Records Successfully Updated...!!"
        Exit Sub
        End If
        rec3.MoveNext
        Next counter2
  End If
  rec.MoveNext
  Next counter
  
        
        
        
        rec.MoveLast
        rec.AddNew
        rec!vcharCIF = Trim(txtCIF.Text)
        rec!vcharName = Trim(txtName.Text)
        rec!vcharAddress = Trim(txtAddress.Text)
        rec!vcharNIC = Trim(txtNIC.Text)
        rec!Flag = "1"
        rec.Update
        
        rec2.Open "select * from Facilities", conn, adOpenStatic, adLockOptimistic
        rec3.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
        rec2.MoveLast
        rec2.AddNew
        rec2!vcharCIF = Trim(txtCIF.Text)
        rec2!vcharOutstanding = Trim(txtAmount)
        rec2!vcharAccNo = Trim(txtACNO.Text)
        rec3.MoveFirst
        Dim counter3 As Integer
        For counter3 = 1 To rec3.RecordCount
        If Trim(cmbFacility.Text) = Trim(rec3!vcharProductName) Then
        rec2!vchaFacilityCode = rec3!vcharProductID
        rec2!intFlag = 1
        rec.Update
        MsgBox "Records Successfully Updated...!"
        Exit Sub
        End If
        rec3.MoveNext
        Next counter3



Exit Sub
errormsg:
'server not responding
If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Exit Sub
    'table not found
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Exit Sub
    
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If


End Sub

Private Sub Form_Load()
'open the database and record sets
 Call basGlobal.opendb
 Set rec = New ADODB.Recordset
 Set rec2 = New ADODB.Recordset
 Set rec3 = New ADODB.Recordset
 
 rec.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
 
 For counter = 1 To rec.RecordCount
 cmbFacility.AddItem (rec!vcharProductName)
 rec.MoveNext
 Next counter
 conn.Close
 Exit Sub
 
 
 'error handler
errormsg:
'server not responding
If Err.Number = -2147467259 Then
    MsgBox " Database Server not responding, Please Check the connections / DSN..! "
    Exit Sub
    'table not found
    ElseIf Err.Number = -2147217865 Then
    MsgBox " Table referred not found in database, Please Check..!"
    Exit Sub
    
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If
End Sub

Private Sub txtACNO_KeyPress(KeyAscii As Integer)
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

Private Sub txtCIF_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    End If
    
    KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If KeyAscii < 65 Or KeyAscii > 122 Then
    
    'allow back space
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    
    'allow dot (.)
    ElseIf KeyAscii = 46 Then
    KeyAscii = 46
    Exit Sub
    
    'allow opening parenthesis
    ElseIf KeyAscii = 40 Then
    KeyAscii = 40
    Exit Sub
    
    'allow closing parenthesis
    ElseIf KeyAscii = 41 Then
    KeyAscii = 41
    Exit Sub
    
    'allow space bar
    ElseIf KeyAscii = 32 Then
    KeyAscii = 32
    Exit Sub
    End If
    
    'disable opening square bracket
    If KeyAscii = 91 Then
    KeyAscii = 0
    Exit Sub
    End If
    
    'diable any other
    KeyAscii = 0
    End If
End Sub

Private Sub txtNIC_KeyPress(KeyAscii As Integer)

If KeyAscii < 48 Or KeyAscii > 57 Then
    
    'allow back space
    If KeyAscii = 8 Then
    KeyAscii = 8
    Exit Sub
    
    'allow "V"
    ElseIf KeyAscii = 86 Then
    KeyAscii = 86
    Exit Sub
    
    'allow "X"
    ElseIf KeyAscii = 88 Then
    KeyAscii = 88
    Exit Sub
    
    End If
    
    'disable any other
    KeyAscii = 0
    Exit Sub

End If

End Sub


