VERSION 5.00
Begin VB.Form MiscForm 
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
   Begin VB.ComboBox cmbMiscType 
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox lblMiscType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Text            =   "Miscellaneous Type"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox lblLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   5400
      TabIndex        =   11
      Text            =   "Location"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox lblPublicationMonthOrSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Text            =   "Publication Month or Season"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox lblPublicationDay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      TabIndex        =   1
      Text            =   "Publication Day"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "Check to keep same year"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox cmbPublicationMonthOrSeason 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtPublicationDay 
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   11775
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Text            =   "Miscellaneous Work Title"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Text            =   "Publication Year"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   12735
   End
End
Attribute VB_Name = "MiscForm"
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
