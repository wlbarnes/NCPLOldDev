VERSION 5.00
Begin VB.Form LegislativeForm 
   Caption         =   "Thesis/Dissertation Type"
   ClientHeight    =   11955
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   14955
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblLegislativeType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Text            =   "Legislative Type"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox lblSuDocNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   7920
      TabIndex        =   20
      Text            =   "SuDoc Number"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox lblStateLegislativeSession 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   10080
      TabIndex        =   19
      Text            =   "State Legislative Session"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox lblSessionOfCongress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   2640
      TabIndex        =   18
      Text            =   "Session of Congress"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox lblNumberOfCongress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Text            =   "Number of Congress"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox lblLegislativeHouse 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   6960
      TabIndex        =   16
      Text            =   "Legislative House"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lblReportOrDocumentNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   5040
      TabIndex        =   15
      Text            =   "Report or Document Number"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox lblUSCCANCitation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   10320
      TabIndex        =   14
      Text            =   "USCCAN Citation"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ComboBox cmbLegislativeType 
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtLegislativeHouse 
      Height          =   285
      Left            =   6960
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtNumberOfCongress 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtSessionOfCongress 
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtStateLegislativeSession 
      Height          =   285
      Left            =   10080
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtUSCCANCitation 
      Height          =   285
      Left            =   10320
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtReportOrDocumentNumber 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtSuDocNumber 
      Height          =   285
      Left            =   7920
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Text            =   "Publication Year"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "Check to keep same year"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   11775
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Text            =   "Legislative Work Title"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   12735
   End
End
Attribute VB_Name = "LegislativeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Place_Form()

   With Me.Text20
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      .Height = 195
      .Left = 5680
      .TabIndex = 77
      .Text = "<<<--------->>>"
      .Top = 7320
      .Width = 975
   End With
   
End Sub

Private Sub Form_Load()
    Call Place_Form
End Sub

Private Sub lblSeparateBottom_Click()

End Sub

Private Sub lblThesisDissertationType_Change()

End Sub

Private Sub lblNumberOfCongress_Click()

End Sub
