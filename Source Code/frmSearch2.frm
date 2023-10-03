VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Customers"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmSearch2.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5940
   Begin VB.TextBox txtReqAmt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   26
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame FrameLarge 
      Height          =   6495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtRequest 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   24
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtFlag 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox txtFacility 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtOutcome 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtresAddr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox txtResName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox txtNIC 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblReqAmt 
         Caption         =   "Amount"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblRequest 
         Caption         =   "Request"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label lblExistingFacilities 
         Caption         =   "Existing Facilities"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label lblFacilityType 
         Caption         =   "Facility "
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label lblOr 
         Caption         =   "OR"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblOutcome 
         Caption         =   "Outcome"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblDate 
         Caption         =   "Last Visit"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblResAddr 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblResName 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblNIC 
         Caption         =   "N.I.C.No"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Caption         =   "       Search for Customers"
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
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
   'Clear the Form
    
    txtResName.Text = ""
    txtresAddr.Text = ""
    txtName.Text = ""
    txtNIC.Text = ""
    txtAmount.Text = ""
    txtFacility.Text = ""
    txtFlag.Text = ""
    txtOutcome.Text = ""
    txtDate.Text = ""
    txtReqAmt.Text = ""
    txtRequest.Text = ""
    
End Sub

Private Sub cmdSearch_Click()

On Error GoTo errormsg
    
    'If nothing entered
    If txtName.Text = "" And txtNIC.Text = "" Then
    txtName.SetFocus
    MsgBox "Please Enter the NAME or NIC Number...!"
    Exit Sub
    End If
    


    'If Both argument Entered at the same time
    If txtName.Text <> "" And txtNIC.Text <> "" Then
    txtName.Text = ""
    txtNIC.Text = ""
    txtName.SetFocus
    MsgBox "Please Enter Only One Argument...!"
    Exit Sub
    End If


    'if only the NIC entered
    If txtName.Text = "" And txtNIC.Text <> "" Then
    
        'open the database and record sets
        Call basGlobal.opendb
        Set rec = New ADODB.Recordset
        Set rec2 = New ADODB.Recordset
        Set rec3 = New ADODB.Recordset
        Set rec4 = New ADODB.Recordset
        rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
    
        If rec.BOF = True Then
        MsgBox "Customer Not Found...!"
        Exit Sub
        End If
     
     
        'start searching
        'goto the first record
         rec.MoveFirst
         For counter = 1 To rec.RecordCount
         If Trim(rec!vcharNIC) = Trim(txtNIC.Text) Then
         'fill name and address
         txtResName.Text = rec!vcharName
         txtresAddr.Text = rec!vcharAddress
            
            'search for existing facilities
                    rec2.Open "select * from Facilities", conn, adOpenStatic, adLockOptimistic
                    If rec2.BOF = True Then
                    MsgBox "No Facilities Found...!"
                    Exit Sub
                    End If
                    
                    Dim counters As Integer
                    rec2.MoveFirst
                    For counters = 1 To rec2.RecordCount
                    If Trim(rec2!vcharCIF) = Trim(rec!vcharCIF) Then
                    'fill outstanding cage
                     txtAmount.Text = rec2!vcharOutstanding
                      
                        'Tracing the flag (Regular,Arrears, Irregular and Past due)
                        Select Case (Str(rec2!intFlag))
                        Case (0)
                        txtFlag.Text = "REG" 'regular
                        
                        Case (1)
                        txtFlag.Text = "ARR" 'Arrears
                        
                        Case (2)
                        txtFlag.Text = "IRR" 'Irregular
                        
                        Case (3)
                        txtFlag.Text = "PD" 'Past due
                        
                        End Select
                    
                                'search for Facility type
                                rec3.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
                                
                                If rec3.RecordCount < 1 Then
                                MsgBox "Facility Type Not Found...!"
                                Exit Sub
                                End If
                                
                                rec3.MoveFirst
                                Dim counter3 As Integer
                                For counter3 = 1 To rec3.RecordCount
                                If Trim(rec2!vchaFacilityCode) = Trim(rec3!vcharProductID) Then
                                txtFacility.Text = rec3!vcharProductName
                                
                                
                                     'search for Last Visit
                                        rec4.Open "select * from CreditInfo", conn, adOpenStatic, adLockOptimistic
                                
                                        If rec4.RecordCount < 1 Then
                                            MsgBox "Last Visit Details Not Found...!"
                                            Exit Sub
                                            End If
                                
                                            rec4.MoveFirst
                                            Dim counter6 As Integer
                                            For counter6 = 1 To rec4.RecordCount
                                            If Trim(rec!vcharCIF) = Trim(rec4!vcharCIF) Then
                                            txtDate.Text = rec4!dtDateVisited
                                            txtOutcome.Text = rec4!vcharResult
                                            txtReqAmt.Text = rec4!vcharAmount
                                            
                                            
                                                'search for Facility type
                                                
                                
                                                If rec3.RecordCount < 1 Then
                                                MsgBox "Facility Type Not Found...!"
                                                Exit Sub
                                                End If
                                
                                                rec3.MoveFirst
                                                Dim counter7 As Integer
                                                For counter7 = 1 To rec3.RecordCount
                                                If Trim(rec4!vcharFacilityCode) = Trim(rec3!vcharProductID) Then
                                                txtRequest.Text = rec3!vcharProductName
                                                Exit Sub
                                                End If
                                                rec3.MoveNext
                                                Next counter7
                                                txtRequest.Text = "Not Found"
                                                Exit Sub
                                
                                        Else
                                        rec4.MoveNext
                                        End If
                                        Next counter6
                                        txtDate.Text = "Not Found"
                                        txtOutcome.Text = "Not Found"
                                        txtReqAmt.Text = "Not Found"
                                        txtRequest.Text = "Not Found"
                                        Exit Sub
                                
                                
                                
                                
                                Else
                                rec3.MoveNext
                                End If
                                Next counter3
                                txtFacility.Text = "Not Found"
                                
                                Exit Sub
                                
                                
                                
                                
                    
                    
                    
                    
                    
                    
                    
                    
                    Else
                    'go to the next record
                     rec2.MoveNext
                    End If
                'repeat the process till EOF
                Next counters
                ' no facilities at all
                 txtAmount.Text = "Not Found"
                 txtFlag.Text = ""
                 txtFacility.Text = "Not Found"
                 txtRequest.Text = "Not Found"
                 txtReqAmt.Text = "Not Found"
                 txtOutcome.Text = "NotFound"
                 txtDate.Text = " Not Found"
                Exit Sub
            Else
            'no customer found, go the the next record
            rec.MoveNext
            End If
         Next counter
         'no customers found at all
          txtNIC.SetFocus
          MsgBox "Customer Not found !! "
    Exit Sub
    End If
    
 
   
    
    
    
    'only the Name entered
    
    If txtNIC.Text = "" And txtName.Text <> "" Then
    
    'open the database and record sets
    Call basGlobal.opendb
    Set rec = New ADODB.Recordset
    Set rec2 = New ADODB.Recordset
    Set rec3 = New ADODB.Recordset
    Set rec4 = New ADODB.Recordset
    rec.Open "select * from CIF", conn, adOpenStatic, adLockOptimistic
    
        If rec.BOF = True Then
        MsgBox "Customer Not Found...!"
        Exit Sub
        End If
     
     
     'start searching
      'goto the first record
       rec.MoveFirst
       For counter = 1 To rec.RecordCount
        If Trim(rec!vcharName) = Trim(txtName.Text) Then
        'fill name and address
         txtResName.Text = rec!vcharName
         txtresAddr.Text = rec!vcharAddress
            
            'search for existing facilities
                    rec2.Open "select * from Facilities", conn, adOpenStatic, adLockOptimistic
                    
                    If rec2.BOF = True Then
                    MsgBox "No Facilities Found...!"
                    Exit Sub
                    End If
                    
                    Dim counters1 As Integer
                    rec2.MoveFirst
                    For counters1 = 1 To rec2.RecordCount
                    If Trim(rec2!vcharCIF) = Trim(rec!vcharCIF) Then
                    'fill outstanding cage
                     txtAmount.Text = rec2!vcharOutstanding
                      
                        'Tracing the flag (Regular,Arrears, Irregular and Past due)
                        Select Case (Str(rec2!intFlag))
                        Case (0)
                        txtFlag.Text = "REG" 'regular
                        
                        Case (1)
                        txtFlag.Text = "ARR" 'Arrears
                        
                        Case (2)
                        txtFlag.Text = "IRR" 'Irregular
                        
                        Case (3)
                        txtFlag.Text = "PD" 'Past due
                        
                        End Select
                    
                    
                                'search for Facility type
                                rec3.Open "select * from ProductInfo", conn, adOpenStatic, adLockOptimistic
                                
                                If rec3.RecordCount < 1 Then
                                MsgBox "Facility Type Not Found...!"
                                Exit Sub
                                End If
                                
                                rec3.MoveFirst
                                Dim counter4 As Integer
                                For counter4 = 1 To rec3.RecordCount
                                If Trim(rec2!vchaFacilityCode) = Trim(rec3!vcharProductID) Then
                                txtFacility.Text = rec3!vcharProductName
                                
                                
                                        'search for Last Visit
                                        rec4.Open "select * from CreditInfo", conn, adOpenStatic, adLockOptimistic
                                
                                        If rec4.RecordCount < 1 Then
                                            MsgBox "Last Visit Details Not Found...!"
                                            Exit Sub
                                            End If
                                
                                            rec4.MoveFirst
                                            Dim counter5 As Integer
                                            For counter5 = 1 To rec4.RecordCount
                                            If Trim(rec!vcharCIF) = Trim(rec4!vcharCIF) Then
                                            txtDate.Text = rec4!dtDateVisited
                                            txtOutcome.Text = rec4!vcharResult
                                            txtReqAmt.Text = rec4!vcharAmount
                                            
                                            
                                                'search for Facility type
                                                
                                
                                                If rec3.RecordCount < 1 Then
                                                MsgBox "Facility Type Not Found...!"
                                                Exit Sub
                                                End If
                                
                                                rec3.MoveFirst
                                                Dim counter8 As Integer
                                                For counter8 = 1 To rec3.RecordCount
                                                If Trim(rec4!vcharFacilityCode) = Trim(rec3!vcharProductID) Then
                                                txtRequest.Text = rec3!vcharProductName
                                                Exit Sub
                                                End If
                                                rec3.MoveNext
                                                Next counter8
                                                txtRequest.Text = "Not Found"
                                                Exit Sub
                                
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        Else
                                        rec4.MoveNext
                                        End If
                                        Next counter5
                                        txtDate.Text = "Not Found"
                                        txtOutcome.Text = "Not Found"
                                        txtReqAmt.Text = "Not Found"
                                        txtRequest.Text = "Not Found"
                                        Exit Sub
                                
                                
                                
                                
                                
                                Else
                                rec3.MoveNext
                                End If
                                Next counter4
                                txtFacility.Text = "Not Found"
                                
                                Exit Sub
                                                    
                                                    
                    
                    
                    
                    
                    
                    Else
                    'go to the next record
                     rec2.MoveNext
                    End If
                'repeat the process till EOF
                Next counters1
                ' no facilities at all
                 txtAmount.Text = "Not Found"
                 txtFlag.Text = ""
                 txtFacility.Text = "Not Found"
                 txtRequest.Text = "Not Found"
                 txtReqAmt.Text = "Not Found"
                 txtOutcome.Text = "NotFound"
                 txtDate.Text = " Not Found"
                
                
                Exit Sub
            Else
            'no customer found, go the the next record
            rec.MoveNext
            End If
         Next counter
         'no customers found at all
          txtNIC.SetFocus
          MsgBox "Customer Not found !! "
    Exit Sub
    End If
    
   
    
    
    
    
    
   
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


