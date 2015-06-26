VERSION 5.00
Begin VB.Form UnpublishedForm 
   Caption         =   "Thesis/Dissertation Type"
   ClientHeight    =   11955
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblUnpublishedType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Text            =   "Unpublished Type"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox lblThesisDissertationType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Text            =   "Thesis/Dissertation Type"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox lblLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   6240
      TabIndex        =   13
      Text            =   "Location"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   6240
      TabIndex        =   12
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ComboBox cmbThesisDissertationType 
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ComboBox cmbUnpublishedType 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox lblPublicationMonthOrSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Text            =   "Month or Season"
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
      Left            =   7680
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
      Height          =   285
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
      Left            =   6240
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Text            =   "Unpublished Work Title"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   6240
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
Attribute VB_Name = "UnpublishedForm"
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
