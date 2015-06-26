VERSION 5.00
Begin VB.Form ArticleForm 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   10440
      TabIndex        =   0
      Text            =   "Volume"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox lblPublicationMonthOrSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Text            =   "Publication Month or Season"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox lblPage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   8280
      TabIndex        =   2
      Text            =   "Page Number"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox lblPublicationDay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   6240
      TabIndex        =   3
      Text            =   "Publication Day"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "Check to keep same year"
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox lblJournalTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Text            =   "Journal Title"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox lblArticleDesignation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Text            =   "Article Designation"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   8280
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cmbPublicationMonthOrSeason 
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtVolume 
      Height          =   285
      Left            =   10440
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPublicationDay 
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox cmbJournalTitle 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   6135
   End
   Begin VB.ComboBox cmbArticleDesignation 
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   14
      Top             =   3120
      Width           =   11775
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   8520
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox chkKeepSelected 
      Caption         =   "Check to keep same jourrnal selected for multiple entries"
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Text            =   "Article Title"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   8520
      TabIndex        =   18
      Text            =   "Publication Year"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   0
      TabIndex        =   19
      Top             =   1800
      Width           =   12735
   End
End
Attribute VB_Name = "ArticleForm"
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
