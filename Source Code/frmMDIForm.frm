VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMDIForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Automated Credit Evaluator - Commercial Bank of Ceylon Ltd"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   11805
   Icon            =   "frmMDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMDIForm.frx":164A
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crt 
      Left            =   2880
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File       "
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEvaluate 
      Caption         =   "Evaluate"
      Begin VB.Menu mnuPersonal 
         Caption         =   "Personal Borrower"
      End
      Begin VB.Menu mnuProprietorship 
         Caption         =   "Proprietorship"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPartnership 
         Caption         =   "Partnership"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLimitedCompany 
         Caption         =   "Limited Company"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuInput 
      Caption         =   "    Input       "
      Begin VB.Menu mnuNewAdvances 
         Caption         =   "New Advances"
      End
      Begin VB.Menu mnuEditFacilityStatus 
         Caption         =   "Edit Facility Status"
      End
      Begin VB.Menu mnuEditBlackListFlag 
         Caption         =   "Edit Black List Flag"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuCustomer 
         Caption         =   "Search for Customers"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "   Reports       "
      Begin VB.Menu mnuRegularFacilitiesReport 
         Caption         =   "Regular Facilities Report"
      End
      Begin VB.Menu mnuArrearsFacilitiesReport 
         Caption         =   "Arrears Facilities Report"
      End
      Begin VB.Menu mnuIrregularFacilitiesReport 
         Caption         =   "Irregular Facilities Report"
      End
      Begin VB.Menu mnuPastdueFacilitiesReport 
         Caption         =   "Past Due Facilities Report"
      End
      Begin VB.Menu mnuLastVisitReport 
         Caption         =   "Last Visit Report"
      End
      Begin VB.Menu mnuBlackList 
         Caption         =   "Black Listed Customer Report"
      End
   End
   Begin VB.Menu mnuComputations 
      Caption         =   "Computations    "
      Begin VB.Menu mnuLeaseRental 
         Caption         =   "Lease Reantal"
      End
      Begin VB.Menu mnuTermLoanInstallment 
         Caption         =   "Term Loan Installment"
      End
      Begin VB.Menu mnuNivahanaInstalment 
         Caption         =   "Nivahana Instalment"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "Maintenance    "
      Begin VB.Menu mnuAddNewUser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModifyUser 
         Caption         =   "Modify User"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About   "
      Begin VB.Menu mnutHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuVersionInfo 
         Caption         =   "Version Info"
      End
   End
End
Attribute VB_Name = "frmMDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuAddNewUser_Click()
frmOverride2.Show
End Sub



Private Sub mnuArrearsFacilitiesReport_Click()
On Error GoTo errormsg
crt.ReportFileName = App.Path & "/Reports/facilities_with_arrears.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
    
  crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If



End Sub




Private Sub mnuBlackList_Click()
On Error GoTo errormsg


crt.ReportFileName = App.Path & "/Reports/Black_Listed_Customers.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
    crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub

Private Sub mnuCustomer_Click()
frmSearch.Show

End Sub

Private Sub mnuDeleteUser_Click()
frmOverride.Show
End Sub

Private Sub mnuEditBlackListFlag_Click()
frmEditBlackList.Show

End Sub

Private Sub mnuEditFacilityStatus_Click()
frmFacilityStatus.Show

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuIrregularFacilitiesReport_Click()
On Error GoTo errormsg
crt.ReportFileName = App.Path & "/Reports/facilities_in_irregular.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
   
  crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub





Private Sub mnuLastVisitReport_Click()
On Error GoTo errormsg
crt.ReportFileName = App.Path & "/Reports/last_visit_report.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
  crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub

Private Sub mnuLeaseRental_Click()
On Error GoTo errormsg
frmLease.Show
errormsg:
Exit Sub
End Sub

Private Sub mnuLimitedCompany_Click()
frmLtd.Show
End Sub



Private Sub mnuLogout_Click()
Me.Hide
frmLogin.Show
End Sub




Private Sub mnuModifyUser_Click()
frmOverride3.Show
End Sub

Private Sub mnuNewAdvances_Click()
frmInputNewAdvances.Show
End Sub

Private Sub mnuNivahanaInstalment_Click()
frmNivahana.Show

End Sub

Private Sub mnuPartnership_Click()
frmPartner.Show
End Sub


Private Sub mnuPastdueFacilitiesReport_Click()
On Error GoTo errormsg
crt.ReportFileName = App.Path & "/Reports/facilities_past_due.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
  crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub

Private Sub mnuPersonal_Click()
frmPersonal.Show
End Sub

Private Sub mnuProprietorship_Click()
frmProprtr.Show
End Sub



Private Sub mnuRegularFacilitiesReport_Click()
On Error GoTo errormsg
crt.ReportFileName = App.Path & "/Reports/reagular_facilities.rpt"
  crt.DiscardSavedData = True
  crt.WindowState = crptMaximized
    crt.Action = 1

If crt.RecordsRead = 0 Then
MsgBox "NO RECORDS TO DISPLAY !!!"
SendKeys "%{F4}"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub



Private Sub mnuTermLoanInstallment_Click()
frmTermLoan.Show

End Sub

Private Sub mnutHelp_Click()
On Error GoTo errormsg
Shell "explorer Help.pdf"
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
    ElseIf Err.Number = 20507 Then
    MsgBox " Unable to locate the report file..Please check...!!"
    Exit Sub
    Else
    'any other error
    MsgBox Err.Description
    Exit Sub
    
    End If

End Sub

Private Sub mnuVersionInfo_Click()
frmAbout.Show
End Sub
