VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   1440
      Top             =   3600
   End
   Begin VB.Frame Frame 
      Height          =   4176
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Our Interest Is In You"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   480
         TabIndex        =   5
         Top             =   3240
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Commercial Bank of Ceylon Ltd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3120
         TabIndex        =   4
         Top             =   2160
         Width           =   3852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Automated Credit Evaluator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6855
      End
      Begin VB.Image imgLogo 
         Height          =   1170
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright      :   Thushara Widyarathna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3480
         TabIndex        =   1
         Top             =   2760
         Width           =   3012
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version    1.02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Timer_Timer()
On Error GoTo errormsg

frmLogin.Show
Timer.Interval = 0
Unload Me
Exit Sub

errormsg:
Unload Me
Exit Sub



End Sub


