VERSION 5.00
Begin VB.Form ArticleForm 
   Caption         =   "Form1"
   ClientHeight    =   11955
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkKeepSelected 
      Caption         =   "Check to keep same jourrnal selected for multiple entries"
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Text            =   "Article Title"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   5640
      TabIndex        =   25
      Text            =   "Page Number"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   11520
      TabIndex        =   24
      Text            =   "Publiction Year"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   9720
      TabIndex        =   23
      Text            =   "Publiction Day"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   7800
      TabIndex        =   22
      Text            =   "Volume"
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   21
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   20
      Text            =   "Publiction Month (or Season)"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Text            =   "Article Designation"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   11520
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   12135
   End
   Begin VB.ComboBox cmbArticleDesignation 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin VB.ComboBox cmbJournalTitle 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   6135
   End
   Begin VB.TextBox txtPublicationDay 
      Height          =   285
      Left            =   9720
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtVolume 
      Height          =   285
      Left            =   7800
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtArticleID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   5
      Top             =   12960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtChapterID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   12720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtLegislativeID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9600
      TabIndex        =   3
      Top             =   12840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUnpublishedID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   12960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtMiscID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   1
      Top             =   12120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   12735
   End
   Begin VB.Label lblSeriesVolume 
      Caption         =   "Series Volume"
      Height          =   255
      Left            =   9720
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Legis ID"
      Height          =   255
      Left            =   8760
      TabIndex        =   17
      Top             =   12840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Article ID"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   12960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Chapter ID"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   12720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Unpublished ID"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   12960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Misc ID"
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   12120
      Visible         =   0   'False
      Width           =   855
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
