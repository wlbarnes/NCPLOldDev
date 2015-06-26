VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Input and Editing"
   ClientHeight    =   12765
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12765
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox lblWorkingPaper 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   134
      Text            =   "Working Paper Info"
      Top             =   10680
      Width           =   1575
   End
   Begin VB.TextBox txtWorkingPaper 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   133
      Top             =   10920
      Width           =   4815
   End
   Begin VB.CheckBox chkRepublished 
      Caption         =   "Republished?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtJournaTitleShortForm 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   130
      Top             =   10920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview Citation Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   129
      Top             =   10200
      Width           =   2775
   End
   Begin VB.CommandButton lblRecordNumber 
      Appearance      =   0  'Flat
      Caption         =   "Record Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox chkLibraryCollection 
      Caption         =   "In Library Collection?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox lblArrow2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   126
      Text            =   "<<<--------->>>"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox lblDblClicktoAdd2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   125
      Text            =   "frmMain.frx":0000
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   123
      Text            =   "Record Status:"
      Top             =   9765
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "X-Delete Record-X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditJournal 
      Caption         =   "Edit This Journal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCallNumber 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11280
      TabIndex        =   93
      Top             =   12000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewLargerWork 
      Caption         =   "New Larger Work"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   11640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox lblOriginalPublicationDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   91
      Text            =   "Original Publication Date"
      Top             =   10800
      Width           =   1815
   End
   Begin VB.TextBox lblPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9960
      TabIndex        =   90
      Text            =   "Publisher"
      Top             =   10800
      Width           =   1335
   End
   Begin VB.TextBox lblCallNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12960
      TabIndex        =   89
      Text            =   "Call Number"
      Top             =   11520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox lblEditionAndPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   88
      Text            =   "Edition And Printing"
      Top             =   10800
      Width           =   1455
   End
   Begin VB.TextBox lblMiscType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11160
      TabIndex        =   87
      Text            =   "Miscellaneous Type"
      Top             =   11160
      Width           =   1575
   End
   Begin VB.TextBox lblLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11280
      TabIndex        =   86
      Text            =   "Location"
      Top             =   11880
      Width           =   1215
   End
   Begin VB.TextBox lblThesisDissertationType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8520
      TabIndex        =   85
      Text            =   "Thesis/Dissertation Type"
      Top             =   11160
      Width           =   1815
   End
   Begin VB.TextBox lblUnpublishedType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   84
      Text            =   "Unpublished Type"
      Top             =   11160
      Width           =   1575
   End
   Begin VB.TextBox lblUSCCANCitation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11760
      TabIndex        =   83
      Text            =   "USCCAN Citation"
      Top             =   11520
      Width           =   1455
   End
   Begin VB.TextBox lblReportOrDocumentNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6480
      TabIndex        =   82
      Text            =   "Report or Document Number"
      Top             =   11520
      Width           =   2175
   End
   Begin VB.TextBox lblLegislativeHouse 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8400
      TabIndex        =   81
      Text            =   "Legislative House"
      Top             =   10680
      Width           =   1335
   End
   Begin VB.TextBox lblNumberOfCongress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   80
      Text            =   "Number of Congress"
      Top             =   11520
      Width           =   1575
   End
   Begin VB.TextBox lblSessionOfCongress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   79
      Text            =   "Session of Congress"
      Top             =   11520
      Width           =   1575
   End
   Begin VB.TextBox lblStateLegislativeSession 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11520
      TabIndex        =   78
      Text            =   "State Legislative Session"
      Top             =   10680
      Width           =   1815
   End
   Begin VB.TextBox lblSuDocNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9360
      TabIndex        =   77
      Text            =   "SuDoc Number"
      Top             =   11520
      Width           =   1215
   End
   Begin VB.TextBox lblLegislativeType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   76
      Text            =   "Legislative Type"
      Top             =   10680
      Width           =   1215
   End
   Begin VB.TextBox lblSeriesVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   75
      Text            =   "Series Volume"
      Top             =   11880
      Width           =   1215
   End
   Begin VB.TextBox lblTitleOfSeriesIfNotIssuedByAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10800
      TabIndex        =   74
      Text            =   "Title Of Series If Not Issued By Author"
      Top             =   11880
      Width           =   2775
   End
   Begin VB.TextBox lblLargerWorkTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   73
      Text            =   "Larger Work Title"
      Top             =   11160
      Width           =   1335
   End
   Begin VB.ComboBox cmbPagination 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      TabIndex        =   72
      Top             =   9840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox lblVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10440
      TabIndex        =   108
      Text            =   "Volume"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox lblPublicationMonthOrSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   105
      Text            =   "Publication Month or Season"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox lblPage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   107
      Text            =   "Page Number"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox chkSource 
      Caption         =   "Check to keep same type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox lblPublicationDay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6240
      TabIndex        =   106
      Text            =   "Publication Day"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "Check to keep same year"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox lblKeywords 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   70
      Text            =   "Select Keywords"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox lblJournalTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   99
      Text            =   "Journal Title"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox lblSourceType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   94
      Text            =   "Source Type"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox lblArticleDesignation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   104
      Text            =   "Article Designation"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox lblInputInitials 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11280
      TabIndex        =   97
      Text            =   "Input Initials"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox lblDateUpdated 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9600
      TabIndex        =   96
      Text            =   "Date Updated"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox lblPublicationYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8040
      TabIndex        =   95
      Text            =   "Date Added"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   69
      Text            =   "Unchanged"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.ListBox lstNewKeywords 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   10080
      Sorted          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdGetNewKeywords 
      Caption         =   "Suggest New Keywords"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdNewAuthor 
      Caption         =   "New Author"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtMiscID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   58
      Top             =   10800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUnpublishedID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   57
      Top             =   10320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtLegislativeID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   56
      Top             =   10200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTreatiseID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   55
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtChapterID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   54
      Top             =   10080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArticleID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   53
      Top             =   10440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbRecordNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":001E
      Left            =   480
      List            =   "frmMain.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextRecord 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreviousRecord 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   9720
      Width           =   1215
   End
   Begin VB.ListBox lstKeywords 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   6840
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentKeywords 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5640
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4215
   End
   Begin VB.ComboBox cmbAETChoice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtSuDocNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   43
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtLargerWorkID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   41
      Top             =   12000
      Width           =   1455
   End
   Begin VB.ComboBox cmbLargerWorkTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   9960
      Width           =   5295
   End
   Begin VB.TextBox txtReportOrDocumentNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   38
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox txtUSCCANCitation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   37
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtStateLegislativeSession 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   36
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtSessionOfCongress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   35
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtNumberOfCongress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   34
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtLegislativeHouse 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12960
      TabIndex        =   33
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox cmbLegislativeType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12960
      TabIndex        =   32
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ComboBox cmbMiscType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11760
      TabIndex        =   31
      Top             =   10320
      Width           =   2175
   End
   Begin VB.TextBox txtLocation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   30
      Top             =   10440
      Width           =   2415
   End
   Begin VB.ComboBox cmbUnpublishedType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8760
      TabIndex        =   27
      Top             =   10320
      Width           =   1695
   End
   Begin VB.ComboBox cmbThesisDissertationType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9600
      TabIndex        =   26
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CheckBox chkAllChaptersBySameAuthor 
      Caption         =   "All Chapters By Same Author?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   11520
      Width           =   2655
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   24
      Top             =   11400
      Width           =   4095
   End
   Begin VB.TextBox txtSeriesVolume 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      TabIndex        =   23
      Top             =   9960
      Width           =   1095
   End
   Begin VB.TextBox txtOriginalPublicationDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13200
      TabIndex        =   22
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtPublisher 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9600
      TabIndex        =   18
      Top             =   12360
      Width           =   4815
   End
   Begin VB.TextBox txtEditionAndPrinting 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Top             =   11160
      Width           =   1215
   End
   Begin VB.TextBox txtOrganizationIssuingNewsletter 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   12360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtPage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cmbPublicationMonthOrSeason 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtVolume 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPublicationDay 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtJournalID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   11760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbJournalTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   6135
   End
   Begin VB.ComboBox cmbArticleDesignation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   11775
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtInputInitials 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11280
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtDateUpdated 
      BackColor       =   &H80000013&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   9600
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtDateAdded 
      BackColor       =   &H80000013&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbSourceType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":0022
      Left            =   2160
      List            =   "frmMain.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox lstAuthors 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   5160
      Width           =   4215
   End
   Begin VB.ListBox lstTranslators 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentAuthors 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5640
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   5160
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentTranslators 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentEditors 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5640
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstEditors 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox lblT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   195
      Left            =   10680
      TabIndex        =   8
      Text            =   "No Translator"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   195
      Left            =   10680
      TabIndex        =   9
      Text            =   "No Editor"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox lblA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   195
      Left            =   10680
      TabIndex        =   44
      Text            =   "No Author"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox lblAETChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   40
      Text            =   "Select"
      Top             =   4890
      Width           =   735
   End
   Begin VB.CheckBox chkKeepSelected 
      Caption         =   "Check to keep same jourrnal selected for multiple entries"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox lblDoubleClickToAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "frmMain.frx":0026
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox lblArrow 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   42
      Text            =   "<<<--------->>>"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   103
      Text            =   "Article Title"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8520
      TabIndex        =   101
      Text            =   "Publication Year"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frmEntryInfo 
      Caption         =   "Entry Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   112
      Top             =   720
      Width           =   5175
   End
   Begin VB.Frame frmRecordInfo 
      Caption         =   "Record Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   111
      Top             =   720
      Width           =   7455
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   113
      Top             =   1800
      Width           =   12735
   End
   Begin VB.Frame frmAuthorInfo 
      Caption         =   "Author Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   114
      Top             =   4560
      Width           =   12735
   End
   Begin VB.Frame frmKeywordInfo 
      Caption         =   "Keyword Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   115
      Top             =   6360
      Width           =   12735
      Begin VB.CommandButton cmdKeywordEntry 
         Caption         =   "Keyword Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   8280
      Width           =   12255
   End
   Begin VB.Frame frmURL 
      Caption         =   "Parallel URL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   45
      Top             =   8040
      Width           =   12735
   End
   Begin VB.Label lblSeparateBottom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   -3360
      TabIndex        =   66
      Top             =   9360
      Width           =   30000
   End
   Begin VB.Label lblMiscID 
      Caption         =   "Misc ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   65
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTreatiseID 
      Caption         =   "Treatise ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   64
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblUnpublishedID 
      Caption         =   "Unpublished ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   63
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblChapterID 
      Caption         =   "Chapter ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   62
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblArticleID 
      Caption         =   "Article ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   61
      Top             =   10440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLegisID 
      Caption         =   "Legis ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   60
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSForms.ToggleButton tglNewRecords 
      Height          =   375
      Left            =   2040
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "New Entries"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglUpdateRecords 
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "Update Records"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglImportRecords 
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "Filter Records"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblLargerWorkID 
      Caption         =   "Larger Work ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   59
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblSeparateTop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      TabIndex        =   67
      Top             =   0
      Width           =   19995
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Index           =   1
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Index           =   2
      Begin VB.Menu mneNewAuthor 
         Caption         =   "New Author"
         Index           =   3
      End
      Begin VB.Menu mnuNewJournal 
         Caption         =   "New Journal"
         Index           =   4
      End
      Begin VB.Menu mnuNewKeyword 
         Caption         =   "New Keyword"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstJournals As ADODB.Recordset
Public rstAuthors As ADODB.Recordset
Public rstEditors As ADODB.Recordset
Public rstTranslators As ADODB.Recordset
Dim rstArticles As ADODB.Recordset
Dim rstChapters As ADODB.Recordset
Dim rstMisc As ADODB.Recordset
Dim rstLegislativeMaterial As ADODB.Recordset
Public rstRecords As ADODB.Recordset
Dim rstRecordsAET As ADODB.Recordset
Dim rstRecordsKeywords As ADODB.Recordset
Dim rstTreatises As ADODB.Recordset
Dim rstUnpublishedWork As ADODB.Recordset

Public rstLargerWorks As ADODB.Recordset
Public rstKeywords As ADODB.Recordset
Public cnReadDatabase As ADODB.Connection
Public cnWriteDatabase As ADODB.Connection
Public cnRemoteReadDatabase As ADODB.Connection
Public cnRemoteWriteDatabase As ADODB.Connection

Dim iSaveListIndex As Integer
Public cAuthors As Collection
Public cEditors As Collection
Public cTranslators As Collection
Public cKeywords As Collection

Private Sub Position_Article_Form()
  With Me.lblVolume
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 8280
      '.TabIndex = 0
      .Text = "Volume"
      .Top = 3600
      .Width = 735
      .Visible = True
   End With
   With Me.lblPage
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 10440
      '.TabIndex = 2
      .Text = "Page Number"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblPublicationMonthOrSeason
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3240
      '.TabIndex = 1
      .Text = "Publication Month or Season"
      .Top = 3600
      .Width = 2055
      .Visible = True
   End With

   With Me.lblPublicationDay
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 6240
      '.TabIndex = 3
      .Text = "Publication Day"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblJournalTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 5
      .Text = "Journal Title"
      .Top = 2160
      .Width = 975
      .Visible = True
   End With
   With Me.lblArticleDesignation
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 6
      .Text = "Article Designation"
      .Top = 3600
      .Width = 1455
      .Visible = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 17
      .Text = "Article Title"
      .Top = 2880
      .Width = 1335
      .Visible = True
   End With
   With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 8520
      '.TabIndex = 18
      .Text = "Publication Year"
      .Top = 2160
      .Width = 1215
      .Visible = True
   End With
   
   With Me.chkYear
      .Caption = "Check to keep same year"
      '.height = 255
      .Left = 9720
      .TabIndex = 999
      .Top = 2400
      .Width = 2175
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   
   With Me.cmbJournalTitle
      '.height = 315
      .Left = 480
      '.Sorted = -1             'True
      .TabIndex = 1
      .Top = 2400
      .Width = 6135
      .Visible = True
      .Enabled = True
   End With
   With Me.txtYear
      '.height = 315
      .Left = 8520
      .TabIndex = 2
      .Top = 2400
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 3
      .Top = 3120
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbArticleDesignation
      '.height = 315
      .Left = 480
      .TabIndex = 4
      .Top = 3840
      .Width = 2175
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbPublicationMonthOrSeason
      '.height = 315
      .Left = 3240
      If Me.cmbPagination = "Nonconsecutive" Then .TabIndex = 5 Else .TabIndex = 9
      .Top = 3840
      .Width = 2055
      .Visible = True
      .Enabled = True
      .BackColor = &H80000005
      If Me.cmbPagination = "Consecutive" Then .BackColor = &H8000000F
   End With
   With Me.txtPublicationDay
      '.height = 285
      .Left = 6240
      If Me.cmbPagination = "Nonconsecutive" Then .TabIndex = 6 Else .TabIndex = 10
      .TabIndex = 6
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
      .BackColor = &H80000005
      If Me.cmbPagination = "Consecutive" Then .BackColor = &H8000000F
   End With
   With Me.txtVolume
      '.height = 285
      .Left = 8280
      If Me.cmbPagination = "Consecutive" Then .TabIndex = 5 Else .TabIndex = 9
      '.TabIndex = 7
      .Top = 3840
      .Width = 1215
      .Visible = True
      .Enabled = True
      .BackColor = &H80000005
      If Me.cmbPagination = "Nonconsecutive" Then .BackColor = &H8000000F

   End With
   With Me.txtPage
      '.height = 285
      .Left = 10440
      .TabIndex = 8
      .Top = 3840
      .Width = 1215
      .Visible = True
      .Enabled = True
   End With
   
   With Me.cmdNewJournal
      .Caption = "New Journal"
      '.Enabled = 0             'False
      '.height = 315
      .Left = 6720
      .TabIndex = 999
      .Top = 2520
      .Visible = 1             'False
      .Width = 1455
      .Visible = True
      .Enabled = True
   End With
      With Me.cmdEditJournal
      .Caption = "Edit This Journal"
      '.Enabled = 0             'False
      '.height = 315
      .Left = 6720
      .TabIndex = 999
      .Top = 2160
      .Visible = 1             'False
      .Width = 1455
      .Visible = True
      .Enabled = True
   End With
   With Me.chkKeepSelected
      .Caption = "Check to keep same jourrnal selected for multiple entries"
      '.height = 195
      .Left = 1440
      .TabIndex = 999
      .Top = 2160
      .Width = 5895
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
End Sub
   
Private Sub Position_Treatise_Form()
  With Me.lblEditionandPrinting
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3000
      '.TabIndex = 17
      .Text = "Edition And Printing"
      .Top = 2880
      .Width = 1455
      .Visible = True
   End With
   With Me.lblCallNumber
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 10440
      '.TabIndex = 16
      .Text = "Call Number"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblTitleOfSeriesIfNotIssuedByAuthor
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 1800
      '.TabIndex = 15
      .Text = "Title Of Series If Not Issued By Author"
      .Top = 3600
      .Width = 2775
      .Visible = True
   End With
   With Me.lblPublisher
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 7440
      '.TabIndex = 14
      .Text = "Publisher"
      .Top = 2880
      .Width = 1335
      .Visible = True
   End With
   With Me.lblSeriesVolume
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 13
      .Text = "Series Volume"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblOriginalPublicationDate
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 5160
      '.TabIndex = 12
      .Text = "Original Publication Date"
      .Top = 2880
      .Width = 1815
      .Visible = True
   End With
   With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 10
      .Text = "Publication Year"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 9
      .Text = "Title"
      .Top = 2160
      .Width = 1335
      .Visible = True
   End With
   
   
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 1
      .Top = 2400
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.txtYear
      '.height = 315
      .Left = 480
      .TabIndex = 2
      .Top = 3120
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   With Me.txtEditionandPrinting
      '.height = 285
      .Left = 3000
      .TabIndex = 3
      .Top = 3120
      .Width = 1695
      .Visible = True
      .Enabled = True
   End With
   With Me.txtOriginalPublicationDate
      '.height = 285
      .Left = 5160
      .TabIndex = 4
      .Top = 3120
      .Width = 1815
      .Visible = True
      .Enabled = True
   End With
   With Me.txtPublisher
      '.height = 285
      .Left = 7440
      .TabIndex = 5
      .Top = 3120
      .Width = 4815
      .Visible = True
      .Enabled = True
   End With
   With Me.txtSeriesVolume
      '.height = 285
      .Left = 480
      .TabIndex = 6
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   With Me.txtTitleOfSeriesIfNotIssuedByAuthor
      '.height = 285
      .Left = 1800
      .TabIndex = 7
      .Top = 3840
      .Width = 8415
      .Visible = True
      .Enabled = True
   End With
   With Me.txtCallNumber
      '.height = 285
      .Left = 10440
      .TabIndex = 8
      .Top = 3840
      .Width = 1815
      .Visible = True
      .Enabled = True
   End With

   With Me.chkYear
      .Caption = "Keep year"
      '.height = 495
      .Left = 1680
      '.TabIndex = 999
      .Top = 2860
      .Width = 1215
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   
End Sub
Private Sub Position_Chapter_Form()
   With Me.lblPage
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 9000
      '.TabIndex = 2
      .Text = "Page Number"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
      With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 10920
      '.TabIndex = 18
      .Text = "Publication Year"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.chkKeepSelected
      .Caption = "Check to keep same Larger Work selected for multiple entries"
      '.height = 195
      .Left = 2040
      '.TabIndex = 10
      .Top = 2880
      .Width = 4815
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   With Me.cmdNewLargerWork
      .Caption = "New Larger Work"
      '.height = 315
      .Left = 10320
      '.TabIndex = 9
      .Top = 3100
      .Width = 1935
      .Visible = True
      .Enabled = True
   End With
   '   With Me.cmdEditLargerWork
   '   .Caption = "Edit This Larger Work"
   '   '.height = 315
   '   .Left = 10320
   '   '.TabIndex = 9
   '   .Top = 2880
   '   .Width = 1935
   '   .Visible = True
   '   .Enabled = True
   'End With
   With Me.lblLargerWorkTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 8
      .Text = "Larger Work Title"
      .Top = 2880
      .Width = 1335
      .Visible = True
   End With
   With Me.lblTitleOfSeriesIfNotIssuedByAuthor
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 2040
      '.TabIndex = 6
      .Text = "Title Of Series If Not Issued By Author"
      .Top = 3600
      .Width = 2775
      .Visible = True
   End With
   With Me.lblSeriesVolume
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 5
      .Text = "Series Volume"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 3
      .Text = "Chapter Title"
      .Top = 2160
      .Width = 1335
      .Visible = True
   End With
   
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 1
      .Top = 2400
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbLargerWorkTitle
      '.height = 315
      .Left = 480
      .TabIndex = 2
      .Top = 3120
      .Width = 9615
      .Visible = True
      .Enabled = True
   End With
   With Me.txtSeriesVolume
      '.height = 285
      .Left = 480
      .TabIndex = 3
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   With Me.txtTitleOfSeriesIfNotIssuedByAuthor
      '.height = 285
      .Left = 2040
      .TabIndex = 4
      .Top = 3840
      .Width = 6374
      .Visible = True
      .Enabled = True
   End With
   With Me.txtPage
      '.height = 285
      .Left = 9000
      .TabIndex = 5
      .Top = 3840
      .Width = 1215
      .Visible = True
      .Enabled = True
   End With
   With Me.txtYear
      '.height = 315
      .Left = 10920
      .TabIndex = 6
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
      With Me.chkYear
      .Caption = "Keep year"
      '.height = 495
      .Left = 10920
      '.TabIndex = 999
      .Top = 4200
      .Width = 1215
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
End Sub
Private Sub Position_Legislative_Form()
  With Me.lblLegislativeType
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 21
      .Text = "Legislative Type"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   With Me.lblSuDocNumber
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 7920
      '.TabIndex = 20
      .Text = "SuDoc Number"
      .Top = 3720
      .Width = 1215
      .Visible = True
   End With
   With Me.lblStateLegislativeSession
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 10080
      '.TabIndex = 19
      .Text = "State Legislative Session"
      .Top = 2880
      .Width = 1815
      .Visible = True
   End With
   With Me.lblSessionOfCongress
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 2640
      '.TabIndex = 18
      .Text = "Session of Congress"
      .Top = 3720
      .Width = 1575
      .Visible = True
   End With
   With Me.lblNumberOfCongress
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 17
      .Text = "Number of Congress"
      .Top = 3720
      .Width = 1575
      .Visible = True
   End With
   With Me.lblLegislativeHouse
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 6960
      '.TabIndex = 16
      .Text = "Legislative House"
      .Top = 2880
      .Width = 1335
      .Visible = True
   End With
   With Me.lblReportOrDocumentNumber
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 5040
      '.TabIndex = 15
      .Text = "Report or Document Number"
      .Top = 3720
      .Width = 2175
      .Visible = True
   End With
   With Me.lblUSCCANCitation
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 10320
      '.TabIndex = 14
      .Text = "USCCAN Citation"
      .Top = 3720
      .Width = 1455
      .Visible = True
   End With
   With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 2880
      '.TabIndex = 5
      .Text = "Publication Year"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      .TabIndex = 1
      .Text = "Legislative Work Title"
      .Top = 2160
      .Width = 1935
      .Visible = True
      '.Enabled = True
   End With
   
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 1
      .Top = 2400
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbLegislativeType
      '.height = 315
      .Left = 480
      .TabIndex = 2
      .Top = 3120
      .Width = 2055
      .Visible = True
      .Enabled = True
   End With
      With Me.txtYear
      '.height = 315
      .Left = 2880
      .TabIndex = 3
      .Top = 3120
      .Width = 1095
      .Visible = True
      .Enabled = True
      
   End With
   With Me.txtLegislativeHouse
      '.height = 285
      .Left = 6960
      .TabIndex = 4
      .Top = 3120
      .Width = 2175
      .Visible = True
      .Enabled = True
   End With
   With Me.txtStateLegislativeSession
      '.height = 285
      .Left = 10080
      .TabIndex = 5
      .Top = 3120
      .Width = 2175
      .Visible = True
      .Enabled = True
   End With
   With Me.txtNumberOfCongress
      '.height = 285
      .Left = 480
      .TabIndex = 6
      .Top = 3960
      .Width = 1575
      .Visible = True
      .Enabled = True
   End With
   With Me.txtSessionOfCongress
      '.height = 285
      .Left = 2640
      .TabIndex = 7
      .Top = 3960
      .Width = 1695
      .Visible = True
      .Enabled = True
   End With
   With Me.txtReportOrDocumentNumber
      '.height = 285
      .Left = 5040
      .TabIndex = 8
      .Top = 3960
      .Width = 2175
      .Visible = True
      .Enabled = True
   End With
   With Me.txtSuDocNumber
      '.height = 285
      .Left = 7920
      .TabIndex = 9
      .Top = 3960
      .Width = 1695
      .Visible = True
      .Enabled = True
   End With

   With Me.txtUSCCANCitation
      '.height = 285
      .Left = 10320
      .TabIndex = 10
      .Top = 3960
      .Width = 1935
      .Visible = True
      .Enabled = True
   End With
   
   
   With Me.chkYear
      .Caption = "Check to keep same year"
      '.height = 255
      .Left = 4080
      '.TabIndex = 3
      .Top = 3120
      .Width = 2175
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   
End Sub
Private Sub Position_Misc_Form()
   
   With Me.lblMiscType
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 12
      .Text = "Miscellaneous Type"
      .Top = 2880
      .Width = 1575
      .Visible = True
   End With
   With Me.lblLocation
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 7500
      '.TabIndex = 11
      .Text = "Location"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   With Me.lblWorkingPaper
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 5400
      '.TabIndex = 11
      .Text = "Working Paper Info"
      .Top = 3600
      .Width = 1600
      .Visible = True
   End With
   With Me.lblPublicationMonthOrSeason
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 0
      .Text = "Publication Month or Season"
      .Top = 3600
      .Width = 2055
      .Visible = True
   End With
   With Me.lblPublicationDay
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3480
      '.TabIndex = 1
      .Text = "Publication Day"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.chkYear
      .Caption = "Check to keep same year"
      '.height = 255
      .Left = 4680
      '.TabIndex = 2
      .Top = 3120
      .Width = 2175
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 7
      .Text = "Miscellaneous Work Title"
      .Top = 2160
      .Width = 1935
      .Visible = True
   End With
   With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3480
      '.TabIndex = 8
      .Text = "Publication Year"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 1
      .Top = 2400
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbMiscType
      '.height = 315
      .Left = 480
      .TabIndex = 2
      .Top = 3120
      .Width = 2295
      .Visible = True
      .Enabled = True
   End With
   With Me.txtYear
      '.height = 315
      .Left = 3480
      .TabIndex = 3
      .Top = 3120
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   
   With Me.txtLocation
      '.height = 285
      .Left = 7500
      .TabIndex = 4
      .Top = 3120
      .Width = 4000
      .Visible = True
      .Enabled = True
   End With
   
   With Me.cmbPublicationMonthOrSeason
      '.height = 315
      .Left = 480
      .TabIndex = 5
      .Top = 3840
      .Width = 2295
      .Visible = True
      .Enabled = True
   End With
   With Me.txtPublicationDay
      '.height = 315
      .Left = 3480
      .TabIndex = 6
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   
    
   With Me.txtWorkingPaper
      '.height = 285
      .Left = 5400
      .TabIndex = 7
      .Top = 3840
      .Width = 6895
      .Visible = True
      .Enabled = True
   End With
      
      
      
End Sub
Private Sub Position_Unpublished_Form()
    With Me.lblUnpublishedType
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 15
      .Text = "Unpublished Type"
      .Top = 2880
      .Width = 1575
      .Visible = True
   End With
   With Me.lblThesisDissertationType
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3480
      '.TabIndex = 14
      .Text = "Thesis/Dissertation Type"
      .Top = 2880
      .Width = 1815
      .Visible = True
   End With
   With Me.lblLocation
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 6240
      '.TabIndex = 13
      .Text = "Location"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.lblPublicationMonthOrSeason
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 0
      .Text = "Month or Season"
      .Top = 3600
      .Width = 2055
      .Visible = True
   End With
   With Me.lblPublicationDay
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 3480
      '.TabIndex = 1
      .Text = "Publication Day"
      .Top = 3600
      .Width = 1215
      .Visible = True
   End With
   With Me.chkYear
      .Caption = "Check to keep same year"
      '.height = 255
      .Left = 7680
      '.TabIndex = 2
      .Top = 3120
      .Width = 2175
      .Visible = True
      If Me.tglNewRecords.Value = True Then .Enabled = True
   End With
   With Me.lblTitle
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      '.TabIndex = 7
      .Text = "Unpublished Work Title"
      .Top = 2160
      .Width = 1935
      .Visible = True
   End With
   With Me.lblYear
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .BorderStyle = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 6240
      '.TabIndex = 8
      .Text = "Publication Year"
      .Top = 2880
      .Width = 1215
      .Visible = True
   End With
   
   With Me.txtTitle
      '.height = 315
      .Left = 480
      .TabIndex = 1
      .Top = 2400
      .Width = 11775
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbUnpublishedType
      '.height = 315
      .Left = 480
      .TabIndex = 2
      .Top = 3120
      .Width = 2295
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbThesisDissertationType
      ''.height = 315
      .Left = 3480
      .TabIndex = 3
      .Top = 3120
      .Width = 1935
      .Visible = True
      .Enabled = True
   End With
   With Me.txtYear
      '.height = 315
      .Left = 6240
      .TabIndex = 4
      .Top = 3120
      .Width = 1215
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbPublicationMonthOrSeason
      '.height = 315
      .Left = 480
      .TabIndex = 5
      .Top = 3840
      .Width = 2295
      .Visible = True
      .Enabled = True
   End With
   With Me.txtPublicationDay
      '.height = 285
      .Left = 3480
      .TabIndex = 6
      .Top = 3840
      .Width = 1095
      .Visible = True
      .Enabled = True
   End With
   With Me.txtLocation
      '.height = 285
      .Left = 6240
      .TabIndex = 7
      .Top = 3840
      .Width = 2415
      .Visible = True
      .Enabled = True
   End With
   
End Sub


Private Sub Position_Initial_Form()
   With Me.chkSource
      .Caption = "Check to keep same type"
      If Me.tglNewRecords = True Then .Enabled = True Else .Enabled = False
      '.height = 255
      .Left = 3240
      .TabIndex = 119
      .Top = 960
      .Width = 2175
      .Visible = True
      .Enabled = True
   End With


   With Me.lblKeywords
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      .TabIndex = 116
      .Text = "Select Keywords"
      .Top = 7080
      .Width = 1335
      .Visible = True
   End With

   With Me.lblSourceType
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 2160
      .TabIndex = 113
      .Text = "Source Type"
      .Top = 960
      .Width = 1335
      .Visible = True
   End With
   With Me.lblInputInitials
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 11280
      .TabIndex = 105
      .Text = "Input Initials"
      .Top = 960
      .Width = 1215
      .Visible = True
   End With
   With Me.lblDateUpdated
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 9600
      .TabIndex = 104
      .Text = "Date Updated"
      .Top = 960
      .Width = 1215
      .Visible = True
   End With
      With Me.txtStatus
      .BackColor = &H80000011
      .Enabled = 0             'False
      '.height = 285
      .Left = 6720
      .TabIndex = 101
      .Text = "Status:Not Saved"
      .Top = 11520
      .Visible = 0             'False
      .Width = 1695
      .Enabled = True
   End With
   With Me.lstNewKeywords
      '.height = 840
      .Left = 10080
      .Sorted = -1             'True
      .TabIndex = 1
      .Top = 7320
      .Width = 2535
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdGetNewKeywords
      .Caption = "Suggest New Keywords"
      '.height = 255
      .Left = 10080
      .TabIndex = 2
      .Top = 7080
      .Width = 2535
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdNewKeyword
      .Caption = "New Keyword"
      '.height = 255
      .Left = 3120
      .TabIndex = 3
      .Top = 7080
      .Width = 1455
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdNewAuthor
      .Caption = "New Author"
      '.height = 255
      .Left = 3240
      .TabIndex = 4
      .Top = 5160
      .Width = 1455
      .Visible = True
      .Enabled = True
   End With
   With Me.txtMiscID
      .Enabled = 0             'False
      '.height = 285
      .Left = 1440
      .TabIndex = 86
      .Top = 10800
      .Visible = 0             'False
      .Width = 1095
   End With
   With Me.txtUnpublishedID
      .Enabled = 0             'False
      '.height = 285
      .Left = 1440
      .TabIndex = 85
      .Top = 10320
      .Visible = 0             'False
      .Width = 1095
      
   End With
   With Me.txtLegislativeID
      .Enabled = 0             'False
      '.height = 285
      .Left = 3600
      .TabIndex = 84
      .Top = 10200
      .Visible = 0             'False
      .Width = 1095
      
   End With
   With Me.txtTreatiseID
      .Enabled = 0             'False
      '.height = 285
      .Left = 1440
      .TabIndex = 83
      .Top = 10560
      .Visible = 0             'False
      .Width = 1095
   End With
   With Me.txtChapterID
      .Enabled = 0             'False
      '.height = 285
      .Left = 1440
      .TabIndex = 82
      .Top = 10080
      .Visible = 0             'False
      .Width = 1095
   End With
   With Me.txtArticleID
      .Enabled = 0             'False
      '.height = 285
      .Left = 3600
      .TabIndex = 81
      .Top = 10440
      .Visible = 0             'False
      .Width = 1095
   End With
   With Me.cmbRecordNumber
      '.height = 315
      '.itemdata        =   "frmMain.frx":0000
      .Left = 480
      '.list            =   "frmMain.frx":0002
      .Style = 2              'Dropdown .list
      .TabIndex = 80
      .Top = 1200
      .Width = 1215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdNextRecord
      .Caption = "-->"
      '.height = 495
      .Left = 9480
      .TabIndex = 75
      .Top = 10680
      .Width = 1215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdPreviousRecord
      .Caption = "<--"
      '.height = 495
      .Left = 4560
      .TabIndex = 74
      .Top = 10680
      .Width = 1215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.cmdSave
      .Caption = "Save"
      '.height = 495
      .Left = 6960
      .TabIndex = 73
      .Top = 10680
      .Width = 1215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.lstKeywords
      '.height = 840
      .Left = 480
      .Sorted = -1             'True
      .TabIndex = 13
      .Top = 7320
      .Width = 4215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.lstCurrentKeywords
      '.height = 840
      .Left = 5640
      .TabIndex = 12
      .Top = 7320
      .Width = 4215
      
      .Visible = True
      .Enabled = True
   End With
   With Me.cmbAETChoice
      '.height = 315
      .Left = 1440
      .TabIndex = 65
      .Top = 5040
      .Width = 1695
      .Visible = True
      .Enabled = True
   End With

   With Me.txtLargerWorkID
      '.height = 285
      .Left = 1920
      .TabIndex = 56
      .Top = 12000
      .Width = 1455
      .Enabled = False
      .Visible = False
   End With
   'With Me.txtNotes
      '.height = 735
   '   .Left = 480
   '   .MultiLine = -1          'True
   '   .TabIndex = 18
   '   .Top = 8880
   '   .Width = 12255
   '   .Visible = True
   '   .Enabled = True
   'End With
   
   With Me.txtURL
      '.height = 735
      .Left = 480
      '.MultiLine = -1          'True
      .TabIndex = 18
      .Top = 8880
      .Width = 12255
      .Visible = True
      .Enabled = True
   End With
   
   With Me.txtJournalID
      .BackColor = &H80000013
      .Enabled = 0             'False
      '.height = 315
      .Left = 3360
      .TabIndex = 11
      .Top = 11760
      .Visible = 0             'False
      .Width = 1215
   End With
   With Me.txtInputInitials
      .BackColor = &H80000013
      '.height = 315
      .Left = 11280
      .TabIndex = 68
      If Me.tglNewRecords.Enabled = True Then .Enabled = True Else .Enabled = False
      .Top = 1200
      .Width = 1335
      
      .Visible = True
      .Enabled = True
   End With
   With Me.txtDateUpdated
      .BackColor = &H80000013
      .Enabled = 0             'False
      .ForeColor = &H80000012
      '.height = 315
      .Left = 9600
      .TabIndex = 69
      .Top = 1200
      .Width = 1335
      .Visible = True
   End With
   With Me.txtDateAdded
      .BackColor = &H80000013
      .Enabled = 0             'False
      '.height = 315
      .Left = 8040
      .TabIndex = 70
      .Top = 1200
      .Width = 1335
      .Visible = True
   End With
   With Me.cmbSourceType
      '.height = 315
      '.itemdata        =   "frmMain.frx":0004
      .Left = 2160
      '.list            =   "frmMain.frx":0006
      .Style = 2              'Dropdown .list
      .TabIndex = 0
      .Top = 1200
      .Width = 3855
      .Visible = True
      .Enabled = True
   End With
   With Me.lstAuthors
      '.height = 840
      .Left = 480
      .Sorted = -1             'True
      .TabIndex = 71
      .Top = 5400
      .Width = 4215
      .Visible = True
      .Enabled = True
   End With
   With Me.lstTranslators
      .Enabled = 0             'False
      '.height = 840
      .Left = 480
      .Sorted = -1             'True
      .TabIndex = 25
      .Top = 5400
      .Visible = 0             'False
      .Width = 4215
      .Visible = True
   End With
   With Me.lstCurrentAuthors
      '.height = 840
      .Left = 5640
      .TabIndex = 77
      .Top = 5400
      .Width = 4215
      .Visible = True
      .Enabled = True
   End With
   With Me.lstCurrentTranslators
      .Enabled = 0             'False
      '.height = 840
      .Left = 5640
      .Sorted = -1             'True
      .TabIndex = 23
      .Top = 5400
      .Visible = 0             'False
      .Width = 4215
   End With
   With Me.lstCurrentEditors
      .Enabled = 0             'False
      '.height = 840
      .Left = 5640
      .Sorted = -1             'True
      .TabIndex = 78
      .Top = 5400
      .Visible = 0             'False
      .Width = 4215
   End With
   With Me.lstEditors
      .Enabled = 0             'False
      '.height = 840
      .Left = 480
      .Sorted = -1             'True
      .TabIndex = 21
      .Top = 5400
      .Visible = 0             'False
      .Width = 4215
   End With
   With Me.lblT
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      .ForeColor = &H80000003
      '.height = 195
      .Left = 10680
      .TabIndex = 6
      .Text = "No Translator"
      .Top = 6120
      .Width = 1095
      .Visible = True
   End With
   With Me.lblE
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      .ForeColor = &H80000003
      '.height = 195
      .Left = 10680
      .TabIndex = 8
      .Text = "No Editor"
      .Top = 5640
      .Width = 1095
      .Visible = True
   End With
   With Me.lblA
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      .ForeColor = &H80000003
      '.height = 195
      .Left = 10680
      .TabIndex = 62
      .Text = "No Author"
      .Top = 5160
      .Width = 1095
      .Visible = True
   End With
   With Me.lblRecordNumber
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 480
      .TabIndex = 63
      .Text = "Record Number"
      .Top = 960
      .Width = 1335
   End With
   With Me.lblAETChoice
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 720
      .TabIndex = 55
      .Text = "Select"
      .Top = 5040
      .Width = 735
      
      .Visible = True
   End With
      With Me.lblDoubleClickToAdd
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 675
      .Left = 4680
      .MultiLine = -1          'True
      .TabIndex = 72
      '.Text            =   "frmMain.frx":0008
      .Top = 7560
      .Width = 975
      
      .Visible = True
   End With
   With Me.lblArrow
      .Appearance = 0         'Flat
      .BackColor = &H8000000F
      .Border.Style = 0        'None
      .Enabled = 0             'False
      '.height = 195
      .Left = 4680
      .TabIndex = 57
      .Text = "<<<--------->>>"
      .Top = 5400
      .Width = 975
      .Visible = True
   End With
   With Me.frmEntryInfo
      .Caption = "Entry Information"
      '.height = 975
      .Left = 7800
      .TabIndex = 42
      .Top = 720
      .Width = 5175
      .Visible = True
   End With
   With Me.frmRecordInfo
      .Caption = "Record Information"
      '.height = 975
      .Left = 240
      .TabIndex = 79
      .Top = 720
      .Width = 6615
      .Visible = True
   End With
   With Me.frmCitationInfo
      .Caption = "Citation Information"
      '.height = 2775
      .Left = 240
      .TabIndex = 96
      .Top = 1800
      .Width = 12735
      .Visible = True
   End With
   With Me.frmAuthorInfo
      .Caption = "Author Information"
      '.height = 1935
      .Left = 240
      .TabIndex = 97
      .Top = 4680
      .Width = 12735
      .Visible = True
   End With
   With Me.frmKeywordInfo
      .Caption = "Keyword Information"
      '.height = 1815
      .Left = 240
      .TabIndex = 66
      .Top = 6720
      .Width = 12735
      .Visible = True
   End With
   'With Me.frmNotes
   '   .Caption = "Notes"
   '   '.height = 1095
   '   .Left = 240
   '   .TabIndex = 67
   '   .Top = 8640
   '   .Width = 12735
   '   .Visible = True
   'End With
    With Me.frmURL
      .Caption = "Parallel URL"
      '.height = 1095
      .Left = 240
      .TabIndex = 67
      .Top = 8640
      .Width = 12735
      .Visible = True
   End With
   With Me.lblSeparateBottom
      Back.Style = 0          'Transparent
      .Border.Style = 1        'Fixed Single
      '.height = 5175
      .Left = -3600
      .TabIndex = 98
      .Top = 12120
      .Width = 30000
      .Visible = True
   End With
   With Me.lblSeparate.Top
      Back.Style = 0          'Transparent
      .Border.Style = 1        'Fixed Single
      '.height = 615
      .Left = -480
      .TabIndex = 99
      .Top = 0
      .Width = 19995
      .Visible = True
   End With
   With Me.lblMiscID
      .Caption = "Misc ID"
      '.height = 255
      .Left = 600
      .TabIndex = 95
      .Top = 10800
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.lblTreatiseID
      .Caption = "Treatise ID"
      '.height = 255
      .Left = 600
      .TabIndex = 94
      .Top = 10560
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.lblUnpublishedID
      .Caption = "Unpublished ID"
      '.height = 255
      .Left = 600
      .TabIndex = 93
      .Top = 10320
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.lblChapterID
      .Caption = "Chapter ID"
      '.height = 255
      .Left = 600
      .TabIndex = 92
      .Top = 10080
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.lblArticleID
      .Caption = "Article ID"
      '.height = 255
      .Left = 2760
      .TabIndex = 91
      .Top = 10440
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.lblLegisID
      .Caption = "Legis ID"
      '.height = 255
      .Left = 2760
      .TabIndex = 90
      .Top = 10200
      .Visible = 0             'False
      .Width = 855
   End With
   With Me.tglNewRecords
      '.height = 375
      .Left = 2040
      .TabIndex = 100
      .Top = 120
      .Width = 1455
      .BackColor = -2147483633
      .ForeColor = -2147483630
      .Display.Style = 6
      .Size = "2566;661"
      .Value = "0"
      .Caption = "New Entries"
      .Font Height = 165
      .FontCharSet = 204
      .FontPitchAndFamily = 2
      .ParagraphAlign = 3
      .Visible = True
   End With
   With Me.tglUpdateRecords
      '.height = 375
      .Left = 6240
      .TabIndex = 36
      .Top = 120
      .Width = 1455
      .BackColor = -2147483633
      .ForeColor = -2147483630
      .Display.Style = 6
      .Size = "2566;661"
      .Value = "0"
      .Caption = "Update Records"
      .Font Height = 165
      .FontCharSet = 204
      .FontPitchAndFamily = 2
      .ParagraphAlign = 3
      .Visible = True
   End With
   With Me.tglImportRecords
      '.height = 375
      .Left = 10320
      .TabIndex = 35
      .Top = 120
      .Width = 1455
      .BackColor = -2147483633
      .ForeColor = -2147483630
      .Display.Style = 6
      .Size = "2566;661"
      .Value = "0"
      .Caption = "Import Records"
      .Font Height = 165
      .FontCharSet = 204
      .FontPitchAndFamily = 2
      .ParagraphAlign = 3
      .Visible = True
   End With
   With Me.lblLargerWorkID
      .Caption = "Larger Work ID"
      '.height = 255
      .Left = 7320
      .TabIndex = 87
      .Top = 10080
      .Width = 1335
   End With
End Sub





Private Sub Populate_Comboboxes()
    Dim icounter As Integer

    Call Populate_Journal_Combobox

    Call Populate_LargerWork_Combobox
    
    Call Populate_AET_Lists
    
    Call Populate_Keyword_List

    Call populate_RecordID_List
    
    cmbAETChoice.AddItem "Authors"
    cmbAETChoice.AddItem "Editors"
    cmbAETChoice.AddItem "Translators"
    
    
    cmbSourceType.AddItem "Journal Article"
    cmbSourceType.AddItem "Treatise"
    cmbSourceType.AddItem "Chapter in Treatise"
    cmbSourceType.AddItem "Unpublished Work"
    cmbSourceType.AddItem "Legislative Material"
    cmbSourceType.AddItem "Nonprint Material"
    
    
    cmbArticleDesignation.AddItem "Abstract"
    cmbArticleDesignation.AddItem "Annotation"
    cmbArticleDesignation.AddItem "Book Note"
    cmbArticleDesignation.AddItem "Book Review"
    cmbArticleDesignation.AddItem "Case Comment"
    cmbArticleDesignation.AddItem "Case Note"
    cmbArticleDesignation.AddItem "Comment"
    cmbArticleDesignation.AddItem "Note"
    cmbArticleDesignation.AddItem "Recent Case"
    cmbArticleDesignation.AddItem "Recent Decision"
    cmbArticleDesignation.AddItem "Recent Development"
    cmbArticleDesignation.AddItem "Recent Statute"
    cmbArticleDesignation.AddItem "Symposium"
    
    cmbPublicationMonthOrSeason.AddItem "Jan."
    cmbPublicationMonthOrSeason.AddItem "Jan./Feb."
    cmbPublicationMonthOrSeason.AddItem "Feb."
    cmbPublicationMonthOrSeason.AddItem "Feb./Mar."
    cmbPublicationMonthOrSeason.AddItem "Mar."
    cmbPublicationMonthOrSeason.AddItem "Mar./Apr."
    cmbPublicationMonthOrSeason.AddItem "Apr."
    cmbPublicationMonthOrSeason.AddItem "Apr./May"
    cmbPublicationMonthOrSeason.AddItem "May"
    cmbPublicationMonthOrSeason.AddItem "May/June"
    cmbPublicationMonthOrSeason.AddItem "June"
    cmbPublicationMonthOrSeason.AddItem "June/July"
    cmbPublicationMonthOrSeason.AddItem "July"
    cmbPublicationMonthOrSeason.AddItem "July/Aug."
    cmbPublicationMonthOrSeason.AddItem "Aug."
    cmbPublicationMonthOrSeason.AddItem "Aug./Sept."
    cmbPublicationMonthOrSeason.AddItem "Sept."
    cmbPublicationMonthOrSeason.AddItem "Sept./Oct."
    cmbPublicationMonthOrSeason.AddItem "Oct."
    cmbPublicationMonthOrSeason.AddItem "Oct./Nov."
    cmbPublicationMonthOrSeason.AddItem "Nov."
    cmbPublicationMonthOrSeason.AddItem "Nov./Dec."
    cmbPublicationMonthOrSeason.AddItem "Dec."
    cmbPublicationMonthOrSeason.AddItem "Dec./Jan."
    cmbPublicationMonthOrSeason.AddItem "Spring"
    cmbPublicationMonthOrSeason.AddItem "Summer"
    cmbPublicationMonthOrSeason.AddItem "Fall"
    cmbPublicationMonthOrSeason.AddItem "Winter"
    
    cmbPagination.AddItem "Consecutive"
    cmbPagination.AddItem "Nonconsecutive"
        
    cmbMiscType.AddItem "Electronic Paginated"
    cmbMiscType.AddItem "Internet Site"
    cmbMiscType.AddItem "Film"
    cmbMiscType.AddItem "Audio Recording"
    cmbMiscType.AddItem "Electronic Material"
                
    cmbLegislativeType.AddItem "Committee Hearing"
    cmbLegislativeType.AddItem "Report"
    cmbLegislativeType.AddItem "Conference Report"
    cmbLegislativeType.AddItem "Committee Print"
    cmbLegislativeType.AddItem "Executive Document"
    cmbLegislativeType.AddItem "Miscellaneous Document"
    cmbLegislativeType.AddItem "State Material"
        
    cmbUnpublishedType.AddItem "Manuscript"
    cmbUnpublishedType.AddItem "Dissertation"
    cmbUnpublishedType.AddItem "Thesis"
    
    cmbThesisDissertationType.AddItem "Ph.D"
    cmbThesisDissertationType.AddItem "M.A."
    cmbThesisDissertationType.AddItem "M.S."
    cmbThesisDissertationType.AddItem "A.B."
    cmbThesisDissertationType.AddItem "B.A."
    cmbThesisDissertationType.AddItem "B.S."
    cmbThesisDissertationType.AddItem "M.B.A."
    cmbThesisDissertationType.AddItem "B.B.A."
End Sub

Public Sub Populate_Journal_Combobox()
    Dim rstTempJournals As ADODB.Recordset
    Dim sTempJournalSource As String
    cmbJournalTitle.Clear
    Set rstTempJournals = New ADODB.Recordset
    sTempJournalSource = "SELECT * FROM tblJournals"
    rstTempJournals.CursorLocation = adUseClient
    rstTempJournals.Open sTempJournalSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    Do While Not rstTempJournals.EOF
        If rstTempJournals.Fields("JournalTitle").Value <> "" Then
            cmbJournalTitle.AddItem rstTempJournals.Fields("JournalTitle").Value
        End If
        rstTempJournals.MoveNext
    Loop
    rstTempJournals.Close
    Set rstTempJournals = Nothing
End Sub
Public Sub Populate_LargerWork_Combobox()
    Dim rstTempLargerWorks As ADODB.Recordset
    Dim sTempLargerWorkSource As String
    cmbLargerWorkTitle.Clear
    Set rstTempLargerWorks = New ADODB.Recordset
    sTempLargerWorkSource = "SELECT * FROM tblLargerWorks"
    rstTempLargerWorks.CursorLocation = adUseClient
    rstTempLargerWorks.Open sTempLargerWorkSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    Do While Not rstTempLargerWorks.EOF
        If rstTempLargerWorks.Fields("LargerWorkTitle").Value <> "" Then
            cmbLargerWorkTitle.AddItem rstTempLargerWorks.Fields("LargerWorkTitle").Value
        End If
        rstTempLargerWorks.MoveNext
    Loop
    rstTempLargerWorks.Close
    Set rstTempLargerWorks = Nothing

    
End Sub

Public Sub Populate_AET_Lists()
    Dim rstTempAET As ADODB.Recordset
    Dim sTempAETSource As String
    Dim sTempAET As String
    
    lstAuthors.Clear
    lstEditors.Clear
    lstTranslators.Clear
        'With rstAuthors
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from qryAETRecords WHERE AETType='Author'")
    'End With
    Set rstTempAET = New ADODB.Recordset
    rstTempAET.CursorLocation = adUseClient
    
    sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Author'"
    rstTempAET.Open sTempAETSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    Do While Not rstTempAET.EOF
        'rstAETLMFRecords.Open "SELECT * FROM qryAETRecords WHERE RecordID=" & iRecNum, cnDatabase, adOpenStatic, adLockPessimistic
            
        sTempAET = Full_AET_Name(rstTempAET)
        'If rstAuthors.Fields("FullName").Value <> "" Then
            'lstAuthors.AddItem rstAuthors.Fields("FullName").Value & " (ID: " & rstAuthors!AETID & ")"
        
        'End If
        lstAuthors.AddItem sTempAET & " (ID: " & rstTempAET!AETID & ")"
        
        rstTempAET.MoveNext
    Loop
    
    rstTempAET.Close
    Set rstTempAET = Nothing
    Set rstTempAET = New ADODB.Recordset
    rstTempAET.CursorLocation = adUseClient
    sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Editor'"
    rstTempAET.Open sTempAETSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rstTempAET.EOF
        sTempAET = Full_AET_Name(rstTempAET)
        'If rstEditors.Fields("FullName").Value <> "" Then
            'lstEditors.AddItem rstEditors.Fields("FullName").Value & " (ID: " & rstEditors!AETID & ")"
        'End If
        lstEditors.AddItem sTempAET & " (ID: " & rstTempAET!AETID & ")"
        
        rstTempAET.MoveNext
    Loop
    
    rstTempAET.Close
    Set rstTempAET = Nothing
    Set rstTempAET = New ADODB.Recordset
    rstTempAET.CursorLocation = adUseClient
    sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Translator'"
    rstTempAET.Open sTempAETSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    
    
    Do While Not rstTempAET.EOF
        sTempAET = Full_AET_Name(rstTempAET)
        
        'If rstTranslators.Fields("FullName").Value <> "" Then
            'lstTranslators.AddItem rstTranslators.Fields("FullName").Value & " (ID: " & rstTranslators!AETID & ")"
        'End If
        lstTranslators.AddItem sTempAET & " (ID: " & rstTempAET!AETID & ")"
        
        rstTempAET.MoveNext
    Loop
    rstTempAET.Close
    Set rstTempAET = Nothing
    
End Sub

Public Function Full_AET_Name(rstRecordset As ADODB.Recordset) As String
    Dim sCurrentAET As String
        sCurrentAET = ""
        If rstRecordset!InstitutionalEntity <> "" Then sCurrentAET = sCurrentAET & rstRecordset!InstitutionalEntity
        If rstRecordset!LastName <> "" Then
            If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
            sCurrentAET = sCurrentAET & rstRecordset!LastName
        End If
        If rstRecordset!FirstName <> "" Then
            If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
            sCurrentAET = sCurrentAET & rstRecordset!FirstName
        End If
        If rstRecordset!MiddleName <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset!MiddleName
        If rstRecordset!Suffix <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset!Suffix
     Full_AET_Name = sCurrentAET
End Function
Private Function Full_AET(rstRecordset As ADODB.Recordset, sType As String) As String
    Dim sCurrentAET As String
        sCurrentAET = ""
        Select Case sType
            Case "FMLS", "FML", "FL"
                If rstRecordset!FirstName <> "" Then
                    'If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
                    sCurrentAET = sCurrentAET & rstRecordset!FirstName
                End If
                If (sType = "FMLS") Or (sType = "FML") Then
                    If rstRecordset!MiddleName <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset!MiddleName
                End If
                If rstRecordset!LastName <> "" Then
                    If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & " "
                    sCurrentAET = sCurrentAET & rstRecordset!LastName
                End If
                If (rstRecordset!Suffix <> "") And (sType = "FMLS") Then
                    If rstRecordset!Suffix = "Jr." Then sCurrentAET = sCurrentAET & ","
                    sCurrentAET = sCurrentAET & " " & rstRecordset!Suffix
                End If
            Case "LFM"
                If rstRecordset!LastName <> "" Then
                    sCurrentAET = sCurrentAET & rstRecordset!LastName
                End If
                If rstRecordset!FirstName <> "" Then
                    If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
                    sCurrentAET = sCurrentAET & rstRecordset!FirstName
                End If
                If rstRecordset!MiddleName <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset!MiddleName
        End Select
     Full_AET = sCurrentAET
End Function

Public Sub Populate_Keyword_List()
    Dim rstTempKeywords As ADODB.Recordset
    Dim sTempSource As String
    
    lstKeywords.Clear
    sTempSource = "SELECT * FROM tblKeywords"
    Set rstTempKeywords = New ADODB.Recordset
    rstTempKeywords.CursorLocation = adUseClient
    rstTempKeywords.Open sTempSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    Do While Not rstTempKeywords.EOF
        If rstTempKeywords.Fields("KeywordOrCodeSection").Value <> "" Then
            lstKeywords.AddItem rstTempKeywords.Fields("KeywordOrCodeSection").Value & " (ID: " & rstTempKeywords!KeywordID & ")"
        End If
        rstTempKeywords.MoveNext
    Loop
    rstTempKeywords.Close
    Set rstTempKeywords = Nothing
End Sub

Public Sub populate_RecordID_List()
    Dim iRecNum As Integer
    Me.cmbRecordNumber.Clear
    If Not rstRecords.EOF Then rstRecords.MoveFirst
    Do While Not rstRecords.EOF
        iRecNum = rstRecords!RecordID
        If cmbRecordNumber.List(Me.cmbRecordNumber.ListCount - 1) = "" Then
             cmbRecordNumber.AddItem iRecNum
        Else
            If Str(cmbRecordNumber.List(cmbRecordNumber.ListCount - 1)) <> Str(iRecNum) Then
                cmbRecordNumber.AddItem iRecNum
            End If
        End If
        rstRecords.MoveNext
    Loop
    cmbRecordNumber.AddItem ("New Record")
    If Me.tglImportRecords.Value = True Then MsgBox "Query executed. " & Me.cmbRecordNumber.ListCount - 1 & " records found."

End Sub
    

Private Sub chkLibraryCollection_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub chkRepublished_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbAETChoice_Click()
    Dim sChoice As String
    sChoice = Me.cmbAETChoice.Text
    Me.lblArrow.Visible = True
    Me.lblAETChoice.Visible = True
    Select Case sChoice
        Case "Authors"
            Erase_Object lstEditors
            Erase_Object lstTranslators
            Erase_Object lstCurrentEditors
            Erase_Object lstCurrentTranslators
            lstAuthors.Visible = True
            lstAuthors.Enabled = True
            lstCurrentAuthors.Visible = True
            lstCurrentAuthors.Enabled = True
            Me.cmdNewAuthor.Caption = "New Author"
        Case "Editors"
            
            Erase_Object lstAuthors
            Erase_Object lstTranslators
            Erase_Object lstCurrentAuthors
            Erase_Object lstCurrentTranslators
            lstEditors.Visible = True
            lstEditors.Enabled = True
            lstCurrentEditors.Visible = True
            lstCurrentEditors.Enabled = True
            Me.cmdNewAuthor.Caption = "New Editor"
            
        Case "Translators"
        
            Erase_Object lstEditors
            Erase_Object lstAuthors
            Erase_Object lstCurrentEditors
            Erase_Object lstCurrentAuthors
            lstCurrentTranslators.Visible = True
            lstCurrentTranslators.Enabled = True
            lstTranslators.Visible = True
            lstTranslators.Enabled = True
            Me.cmdNewAuthor.Caption = "New Translator"

    End Select
End Sub



Private Sub cmbArticleDesignation_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

'Private Sub cmbJournalTitle_click()
    
'    Call Lookup_Journal
'End Sub

Private Sub cmbJournalTitle_click()
    Call Lookup_Journal
    Call Position_Article_Form
    Me.txtStatus.Text = "Not Saved"
End Sub
Private Sub Lookup_Journal()
    Dim rstJournalLookup As ADODB.Recordset
    Dim sJournalSource As String
    Dim sJournalTitle As String
    Dim sJournalTitleShortForm As String
    
    sJournalTitle = cmbJournalTitle.Text
    'sJournalTitle = Replace(sJournalTitle, "'", "*")
    If Not sJournalTitle = "" Then
        sJournalSource = "SELECT * from tblJournals"
        Set rstJournalLookup = New ADODB.Recordset
        rstJournalLookup.CursorLocation = adUseClient
        rstJournalLookup.Open sJournalSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
        rstJournalLookup.MoveFirst
        On Error GoTo Lookup_Journal_Error
        Do Until (rstJournalLookup.EOF) Or (rstJournalLookup!JournalTitle = sJournalTitle)
            rstJournalLookup.MoveNext
        Loop
        
        'rstjournallookup.Find "JournalTitle LIKE '" & sJournalTitle & "'"
Lookup_EOF:
        If Not rstJournalLookup.EOF Then
            Me.txtJournalID = rstJournalLookup!JournalID
            Me.txtJournaTitleShortForm = rstJournalLookup!JournalTitleShortFOrm
            'If rstjournallookup!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstjournallookup!JournalTitleShortForm
            If rstJournalLookup!Pagination <> "" Then Me.cmbPagination = rstJournalLookup!Pagination
        '    If rstjournallookup!CallNumber <> "" Then Me.txtCallNumber = rstjournallookup!CallNumber
            'If rstjournallookup!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstjournallookup!PLaceOfPublication
            frmNewJournal.txtJournalID = rstJournalLookup!JournalID
            frmNewJournal.txtNewJournal = rstJournalLookup!JournalTitle
            frmNewJournal.txtNewJournalShortForm = rstJournalLookup!JournalTitleShortFOrm
            frmNewJournal.cmbPagination.Text = rstJournalLookup!Pagination
            If rstJournalLookup!CallNumber <> Null Then frmNewJournal.txtCallNumber = rstJournalLookup!CallNumber
            If rstJournalLookup!PlaceOfPublication <> Null Then frmNewJournal.txtPlaceOfPublication = rstJournalLookup!PlaceOfPublication
                    
        End If
        rstJournalLookup.Close
        Set rstJournalLookup = Nothing
    End If
    
Lookup_Journal_Error:
    Select Case Err
        Case 0
        Case 3021
            Resume Lookup_EOF
        'Case Else
        '    MsgBox Err.Number & " " & Err.Description
    End Select
    
        
End Sub

Private Sub cmbLargerWorkTitle_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbLargerWorkTitle_Click()
    Dim sTempSource As String
    Dim sLargerWorkTitle As String
    sLargerWorkTitle = cmbLargerWorkTitle.Text
    'sJournalTitle = Replace(sJournalTitle, "'", "*")
    Set rstLargerWorks = New ADODB.Recordset
    sTempSource = "SELECT * FROM tblLargerWorks"
    rstLargerWorks.CursorLocation = adUseClient
    rstLargerWorks.Open sTempSource, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    rstLargerWorks.MoveFirst
    Do Until (rstLargerWorks!LargerWorkTitle = sLargerWorkTitle) Or rstLargerWorks.EOF
        rstLargerWorks.MoveNext
    Loop
    
    'rstJournals.Find "JournalTitle LIKE '" & sJournalTitle & "'"
    If Not rstLargerWorks.EOF Then
        Me.txtLargerWorkID = rstLargerWorks!LargerWorkID
        If rstLargerWorks!CallNumber <> "" Then Me.txtCallNumber = rstLargerWorks!CallNumber
        If rstLargerWorks!EditionAndPrinting <> "" Then Me.txtEditionandPrinting = rstLargerWorks!EditionAndPrinting
        If rstLargerWorks!Publisher <> "" Then Me.txtPublisher = rstLargerWorks!Publisher
        If rstLargerWorks!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate = rstLargerWorks!OriginalPublicationDate
        If rstLargerWorks!SeriesVolume <> "" Then Me.txtSeriesVolume = rstLargerWorks!SeriesVolume
        If rstLargerWorks!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor = rstLargerWorks!TitleOfSeriesIfNotIssuedByAuthor
        
        'TitleOfSeriesIfNotIssuedByAuthor
        
        
    End If
End Sub



Private Sub cmbLegislativeType_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbMiscType_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbPagination_Change()
    Call Position_Article_Form
End Sub


Private Sub cmbPublicationMonthOrSeason_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbRecordNumber_Click()
    Dim iRecNum As Integer
    If Me.cmbRecordNumber.ListIndex <> cmbRecordNumber.ListCount - 1 Then
        'Me.tglUpdateRecords.Value = True
        If IsNumeric(Me.cmbRecordNumber.Text) Then iRecNum = Me.cmbRecordNumber.Text
        rstRecords.MoveFirst
        Do Until rstRecords!RecordID = iRecNum
            rstRecords.MoveNext
        Loop
        'rstRecords.Find "RecordID=" & iRecNum
        Call Erase_Form
        Call Clear_Form
        Me.Refresh
        
        Call Change_Record_Lists
        Call Fill_Form
    End If
    If Me.cmbRecordNumber.Text = "New Record" Then
        Me.tglNewRecords.Value = True
        
    End If
End Sub

Private Sub cmbSourceType_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbSourceType_Click()
    Dim sSourceType As String
    
    
    'Call Fill_Form
    sSourceType = cmbSourceType.Text
    Select Case sSourceType
        Case "Journal Article"
            'Call Article_Form
            Call Erase_Form
            Call Position_Article_Form
        
        Case "Treatise"
            'Call Treatise_Form
            Call Erase_Form
            Call Position_Treatise_Form
        
        Case "Chapter in Treatise"
            'Call Chapter_Form
            Call Erase_Form
            Call Position_Chapter_Form
        
        Case "Unpublished Work"
            'Call Unpublished_Form
            Call Erase_Form
            Call Position_Unpublished_Form
        
        Case "Legislative Material"
            'Call Legislative_Form
            Call Erase_Form
            Call Position_Legislative_Form
        
        Case "Nonprint Material"
            'Call Misc_Form
            Call Erase_Form
            Call Position_Misc_Form
        
    End Select
End Sub

Private Sub cmbSourceType_Validate(Cancel As Boolean)
    
    If cmbSourceType.Text = "" Then
        MsgBox "Please Enter a Source Type."
        Cancel = True
    End If
End Sub

Private Sub cmbThesisDissertationType_Click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmbUnpublishedType_click()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub cmdDelete_Click()
    Dim bYN As String
    Dim iRecordNumber As Integer
    Dim iListindex As Integer
    bYN = MsgBox("Permanently Delete Record", vbOKCancel, "Confirm Deletion")
    Select Case bYN
        Case vbOK
            'MsgBox "Yes"
                iRecordNumber = Me.cmbRecordNumber.Text
                iListindex = Me.cmbRecordNumber.ListIndex
                rstRecords.MoveFirst
                Do Until rstRecords!RecordID = iRecordNumber
                    rstRecords.MoveNext
                Loop
                If Not rstRecords.EOF Then
                    On Error GoTo CancelErr
                    cnWriteDatabase.BeginTrans
                        rstRecords.Delete
                        rstRecords.Update
                    cnWriteDatabase.CommitTrans
                    rstRecords.Requery
                    Me.cmbRecordNumber.RemoveItem (iListindex)
                End If
                Call cmdNextRecord_Click
        Case vbCancel
            'MsgBox "No"
    End Select
CancelErr:
    Select Case Err
        Case 0
        Case Else
            cnWriteDatabase.RollbackTrans
    End Select
End Sub

Private Sub cmdEditJournal_Click()
    Unload frmNewJournal
    frmNewJournal.bEdit = True
    frmNewJournal.Show
    frmNewJournal.Caption = "Edit Journal"
    
    frmNewJournal.bEdit = False
End Sub

Private Sub cmdGetNewKeywords_Click()
    Call suggest_keywords
End Sub
Private Sub suggest_keywords()
    Dim rstOldKeywords As Recordset
    Dim rstKeywordCheck As Recordset
    Dim rstThesaurusCheck As Recordset
    Dim sKeywordText As String
    Dim sThesaurusText As String
    Dim sTitleText As String
    Dim cSuggestedKeywords As Collection
    Dim i As Integer
    Dim sOldKeywordString As String
    Dim iCurrentRecnum As Integer
    Dim bDuplicate As Boolean
    Dim rstBigCategory As Recordset
    Dim sTempText As String
    Dim rstExistingKeywordBigCat As Recordset
    Dim rstJournalKeyword As Recordset
    Dim sJournalLocation As String
    
'the following part will be removed later
    'iCurrentRecnum = Me.cmbRecordNumber.Text
    'Set rstOldKeywords = New Recordset
    'With rstOldKeywords
    '    .ActiveConnection = cnDatabase
    '    .CursorType = adOpenForwardOnly
    '    .LockType = adLockReadOnly
    '    .Open ("SELECT AllKeywords from tblRecordsAllKeywords WHERE RecordID=" & iCurrentRecnum)
    'End With
    'If Not rstOldKeywords.EOF Then sOldKeywordString = rstOldKeywords!AllKeywords
    
    'Set rstOldKeywords = Nothing
    
    Set cSuggestedKeywords = New Collection
    sTitleText = Me.txtTitle.Text
    'next line taken out later
    'sTitleText = sTitleText & " " & sOldKeywordString
    Me.lstNewKeywords.Clear
    Set rstKeywordCheck = New Recordset
    rstKeywordCheck.CursorLocation = adUseClient
    With rstKeywordCheck
        .ActiveConnection = cnReadDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from tblKeywords")
    End With
    
    Set rstThesaurusCheck = New Recordset
    
    Do While Not rstKeywordCheck.EOF
        sKeywordText = rstKeywordCheck!KeywordOrCodeSection
        If InStr(1, sTitleText, sKeywordText) Then
            sKeywordText = sKeywordText & " (ID: " & rstKeywordCheck!KeywordID & ")"
            bDuplicate = False
            For i = 0 To (Me.lstCurrentKeywords.ListCount - 1)
                If Me.lstCurrentKeywords.List(i) = sKeywordText Then bDuplicate = True
            Next
            If Not bDuplicate Then cSuggestedKeywords.Add sKeywordText
            'Me.lstNewKeywords.AddItem sKeywordText
        End If
        rstKeywordCheck.MoveNext
    Loop
    
    Set rstExistingKeywordBigCat = New Recordset
    rstExistingKeywordBigCat.CursorLocation = adUseClient
    With rstExistingKeywordBigCat
        .ActiveConnection = cnReadDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryRecordsKeywordsThesaurus WHERE LargerCategory=1 AND RecordID=" & iCurrentRecnum)
    End With
    
    Do While Not rstExistingKeywordBigCat.EOF
        Set rstBigCategory = New Recordset
        sTempText = rstExistingKeywordBigCat!KeywordOrCodeSection
        rstBigCategory.CursorLocation = adUseClient
        With rstBigCategory
                .ActiveConnection = cnReadDatabase
                .CursorType = adOpenForwardOnly
                .LockType = adLockReadOnly
                .Open ("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'")
        End With
        If Not rstBigCategory.EOF Then
                sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
                sKeywordText = rstThesaurusCheck!KeywordOrCodeSection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
                If InStr(1, sTitleText, sThesaurusText) Then
                    For i = 0 To (Me.lstCurrentKeywords.ListCount - 1)
                            If Me.lstCurrentKeywords.List(i) = sKeywordText Then bDuplicate = True
                    Next
                    For i = 1 To cSuggestedKeywords.Count
                        If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
                    Next
                    If Not bDuplicate Then cSuggestedKeywords.Add sKeywordText
                End If
            End If
        
        Set rstBigCategory = Nothing
    Loop
    
    Set rstExistingKeywordBigCat = Nothing
    'sJournalLocation = Me.txtPlaceOfPublication
    If Not (sJournalLocation = "") Then
    
    
        Set rstJournalKeyword = New Recordset
            With rstJournalKeyword
                    .ActiveConnection = cnReadDatabase
                    .CursorType = adOpenForwardOnly
                    .CursorLocation = adUseClient
                    .LockType = adLockReadOnly
                    .Open ("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sJournalLocation & "'")
            End With
            sKeywordText = sJournalLocation
            If Not rstJournalKeyword.EOF Then sKeywordText = sKeywordText & " (ID: " & rstJournalKeyword!KeywordID & ")"
            bDuplicate = False
            For i = 0 To (Me.lstCurrentKeywords.ListCount - 1)
                If Me.lstCurrentKeywords.List(i) = sKeywordText Then bDuplicate = True
            Next
            If Not bDuplicate Then cSuggestedKeywords.Add sKeywordText
        
        Set rstJournalKeyword = Nothing
    End If
        
    With rstThesaurusCheck
        .ActiveConnection = cnReadDatabase
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryKeywordThesaurus where (NOT (ThesaurusEquivalent IS NULL))")
    End With
    
    Do While Not rstThesaurusCheck.EOF
        If rstThesaurusCheck!LargerCategory = 1 Then
            Set rstBigCategory = New Recordset
            sTempText = rstThesaurusCheck!KeywordOrCodeSection
            With rstBigCategory
                .ActiveConnection = cnReadDatabase
                .CursorType = adOpenForwardOnly
                .CursorLocation = adUseClient
                .LockType = adLockReadOnly
                .Open ("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'")
            End With
            If Not rstBigCategory.EOF Then
                sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
                sKeywordText = rstThesaurusCheck!KeywordOrCodeSection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
                If InStr(1, sTitleText, sThesaurusText) Then
                    For i = 0 To (Me.lstCurrentKeywords.ListCount - 1)
                            If Me.lstCurrentKeywords.List(i) = sKeywordText Then bDuplicate = True
                    Next
                    For i = 1 To cSuggestedKeywords.Count
                        If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
                    Next
                    If Not bDuplicate Then cSuggestedKeywords.Add sKeywordText
                End If
            End If
            Set rstBigCategory = Nothing
        End If
        If rstThesaurusCheck!LargerCategory = 0 Then
            bDuplicate = False
            sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
            sKeywordText = rstThesaurusCheck!KeywordOrCodeSection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
            If InStr(1, sTitleText, sThesaurusText) Then
                For i = 0 To (Me.lstCurrentKeywords.ListCount - 1)
                        If Me.lstCurrentKeywords.List(i) = sKeywordText Then bDuplicate = True
                Next
                For i = 1 To cSuggestedKeywords.Count
                    If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
                Next
                If Not bDuplicate Then cSuggestedKeywords.Add sKeywordText
                
                'Me.lstNewKeywords.AddItem sKeywordText
            End If
        End If
        rstThesaurusCheck.MoveNext
    Loop
        

        
    For i = 1 To cSuggestedKeywords.Count
        Me.lstNewKeywords.AddItem cSuggestedKeywords.Item(i)
    Next
    Set rstThesaurusCheck = Nothing
    Set rstKeywordCheck = Nothing
    Set cSuggestedKeywords = Nothing
End Sub

Private Sub cmdKeywordEntry_Click()
    Unload frmKeywordChange
    frmKeywordChange.Show
End Sub

Private Sub cmdNewAuthor_Click()
    Unload frmNewAuthor
    frmNewAuthor.Show
End Sub

Private Sub cmdNewJournal_Click()
    Unload frmNewJournal
    frmNewJournal.bEdit = False
    frmNewJournal.Caption = "New Journal"
    frmNewJournal.Show
End Sub

Private Sub cmdNewLargerWork_Click()
    frmNewLargerWork.Show
End Sub

Private Sub cmdNextRecord_Click()
    Dim icounter As Integer
    Dim iYN As Integer
    iYN = vbOK
    If Me.txtStatus.Text = "Not Saved" Then
        iYN = (MsgBox("You have made changes without saving record. Continue without saving?", vbOKCancel, "Confirm No Save"))
    End If
    If iYN = vbOK Then
        icounter = (Me.cmbRecordNumber.ListIndex) + 1
        If icounter < Me.cmbRecordNumber.ListCount Then
            Me.cmbRecordNumber.ListIndex = icounter
        End If
    End If
    'rstRecords.MoveNext
    'If rstRecords.EOF Then rstRecords.MoveLast
    'Call Erase_Form
    'Call Change_Record_Lists
    'Call Fill_Form
End Sub

Private Sub cmdPreview_Click()

    Dim report As procs
    Dim iSourceType As Integer
    Dim sAuthor As String
    Dim sEditor As String
    Dim iRecordID As Integer
    Dim sArticleDesignation As String
    Dim sTitle As String
    Dim sJournalTitle As String
    Dim sPage As String
    Dim sMonth As String
    Dim sDay As String
    Dim sYear As String
    Dim sVolume As String
    Dim sSeriesTitle As String
    Dim sEdition As String
    Dim sLegislativeMaterialType As String
    Dim sNameOfHouse As String
    Dim sNumberOfCongress As String
    Dim SessionOfCongress As String
    Dim sStateLegislativeSession As String
    Dim sUSCCANCitation As String
    Dim sReportOrDocumentNumber As String
    Dim sSuDocNumber As String
    Dim sURL As String
    Dim sWorkingPaperInfo As String
    
    Dim wDocument As Word.Application
    Dim i As Integer
    Dim cAETIDs As Collection
    
    If IsNumeric(Me.cmbRecordNumber.Text) Then iRecordID = Me.cmbRecordNumber.Text _
        Else iRecordID = 0
       
    If Me.cmbSourceType = "Journal Article" Then
        If Me.cmbPagination = "Consecutive" Then iSourceType = 1
        If Me.cmbPagination = "Nonconsecutive" Then iSourceType = 2
        'sJournalTitle = frmNewJournal.txtNewJournalShortForm.Text
        sJournalTitle = Me.txtJournaTitleShortForm.Text

        sVolume = Me.txtVolume.Text
    End If
    If Me.cmbSourceType = "Treatise" Then iSourceType = 3
    If Me.cmbSourceType = "Chapter in Treatise" Then
        iSourceType = 4
        sVolume = Me.txtSeriesVolume.Text
        sJournalTitle = Me.cmbLargerWorkTitle.Text
    End If
    If Me.cmbSourceType = "Legislative Material" Then iSourceType = 5
    If Me.cmbSourceType = "Unpublished Work" Then iSourceType = 7
    If Me.cmbSourceType = "Nonprint Material" Then
        iSourceType = 6
        sJournalTitle = Me.txtLocation.Text
        'add something here for workingpaperinfo
    End If
    Set report = New procs
    If iRecordID = 0 Then 'build a collection of AETIDs
        Set cAETIDs = New Collection
        For i = 1 To cAuthors.Count
            cAETIDs.Add cAuthors.Item(i)
        Next
        For i = 1 To cEditors.Count
            cAETIDs.Add cEditors.Item(i)
        Next
    End If
    If iRecordID <> 0 Then Call report.Get_AET_String(iRecordID, Me.cnReadDatabase, sAuthor, sEditor, cAuthors.Count, cEditors.Count) _
        Else If (cAETIDs.Count > 0) Then Call report.Get_AET_String(iRecordID, Me.cnReadDatabase, sAuthor, sEditor, cAuthors.Count, cEditors.Count, cAETIDs)
    ' The Else if in above statement leads to author string error for those not yet saved
    
    sArticleDesignation = Me.cmbArticleDesignation.Text
    sTitle = Me.txtTitle.Text
    sDay = Me.txtPublicationDay.Text
    sPage = Me.txtPage.Text
    sMonth = Me.cmbPublicationMonthOrSeason.Text
    sYear = Me.txtYear.Text
    sSeriesTitle = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
    sEdition = Me.txtEditionandPrinting.Text
    sURL = Me.txtURL.Text
    sWorkingPaperInfo = Me.txtWorkingPaper.Text
    
    Set wDocument = New Word.Application
    wDocument.Documents.Add
    
        
    sLegislativeMaterialType = Me.cmbLegislativeType.Text
    
    
    sNameOfHouse = Me.txtLegislativeHouse.Text
    sNumberOfCongress = Me.txtNumberOfCongress.Text
    SessionOfCongress = Me.txtNumberOfCongress.Text
    sStateLegislativeSession = Me.txtStateLegislativeSession.Text
    sUSCCANCitation = Me.txtUSCCANCitation.Text
    sReportOrDocumentNumber = Me.txtReportOrDocumentNumber.Text
    sSuDocNumber = Me.txtSuDocNumber.Text
                
    
'    Call report.Process_Word_Line(iSourceType, sAuthor, iRecordID, sArticleDesignation, sTitle, sVolume, sJournalTitle, _
                        sPage, sMonth, sDay, sYear, sSeriesTitle, sEditor, sEdition, cEditors.Count, False, frmWordPreview.OLEWord.object.Application _
                        , sLegislativeMaterialType, sNameOfHouse, sNumberOfCongress, SessionOfCongress, sStateLegislativeSession, _
                        sUSCCANCitation, sReportOrDocumentNumber, sSuDocNumber)
                        
     Call report.Process_Word_Line(iSourceType, sAuthor, iRecordID, sArticleDesignation, sTitle, sVolume, sJournalTitle, _
                        sPage, sMonth, sDay, sYear, sSeriesTitle, sEditor, sEdition, cEditors.Count, False, wDocument _
                        , sLegislativeMaterialType, sNameOfHouse, sNumberOfCongress, SessionOfCongress, sStateLegislativeSession, _
                        sUSCCANCitation, sReportOrDocumentNumber, sSuDocNumber, "", sURL, sWorkingPaperInfo)
                        
                        
    wDocument.Visible = True
    
    'frmWordPreview.Show
    'frmWordPreview.OLEWord.object.Application.Documents(1).Close
    'frmWordPreview.OLEWord.object.Application.Application.Quit
    'Set frmWordPreview.OLEWord.object.Application = Nothing
    
    'wDocument.Documents(1).Close
    'wDocument.Application.Quit
    'Set wDocument = Nothing
    'frmWordPreview.OLEWord.object.Application.Documents(1).Close
    'frmWordPreview.OLEWord.object.Application.Application.Quit
    'Set frmWordPreview.OLEWord.object.Application = Nothing
End Sub

Private Sub cmdPreviousRecord_Click()
    Dim icounter As Integer
    Dim iYN As Integer
    iYN = vbOK
    If Me.txtStatus.Text = "Not Saved" Then
        iYN = (MsgBox("You have made changes without saving record. Continue without saving?", vbOKCancel, "Confirm No Save"))
    End If
    If iYN = vbOK Then
        icounter = (Me.cmbRecordNumber.ListIndex) - 1
        If icounter > -1 Then
            Me.cmbRecordNumber.ListIndex = icounter
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Dim sConnectionString As String
    Dim sSourceType As String
    Dim sDateAdded As String
    Dim sDateUpdated As String
    Dim sInputInitials As String
    Dim sTitle As String
    Dim sYear As String
    Dim iLargerWorkID As Integer
    Dim sArticleDesignation As String
    Dim iJournalID As Integer
    Dim sPublicationDay As String
    Dim sPageNumber As String
    Dim sVolume As String
    Dim sCallNumber As String
    Dim sPublicationMonth As String
    'Dim sNotes As String
    Dim sURL As String
    Dim sEditionAndPrinting As String
    Dim sPublisher As String
    Dim sOriginalPublicationDate As String
    Dim sSeriesVolume As String
    Dim sTitleOfSeriesIfNotIssuedByAuthor As String
    Dim bAllChaptersBySameAuthor As String
    Dim sUnpublishedWorkType As String
    Dim sThesisDissertationType As String
    Dim sLocation As String
    Dim sWorkingPaperInfo As String
    Dim sMiscellaneousType As String
    Dim sLegislativeType As String
    Dim sNameOfHouse As String
    Dim sNumberOfCongress As String
    Dim sSessionOfCongress As String
    Dim sStateLegislativeSession As String
    Dim sUSCCANCitation As String
    Dim sReportDocumentNumber As String
    Dim sSuDocNumber As String
    Dim rstAETDelete As ADODB.Recordset
    Dim rstKeywordDelete As ADODB.Recordset
    Dim rstBigRecordIndex As ADODB.Recordset
    Dim iRecordID As Integer
    Dim icounter As Integer
    'Dim lAuthorID As Long
    'Dim rstJournalCheck As ADODB.Recordset
    Dim rstCheck As ADODB.Recordset
    Dim sCheckString As String
    Dim sCheckTitle As String
    Dim bDuplicate As Boolean
    'Dim sSQL As String
    'Dim sArticleDesignation As String
    Dim iArticleID As Integer
    Dim iLegislativeID As Integer
    Dim iChapterID As Integer
    Dim iTreatiseID As Integer
    Dim iUnpublishedID As Integer
    Dim iMiscID As Integer
    Dim sSQLString As String
    Dim rstRecordsKeywordsThesaurus As ADODB.Recordset
    Dim rstAllKeyword As ADODB.Recordset
    Dim sAllKeywordString As String
    Dim cAllKeywords As Collection
    Dim sCurrentKeyword As String
    Dim i As Integer
    Dim rstRecordsAuthors As ADODB.Recordset
    Dim rstAllAuthor As ADODB.Recordset
    Dim sFullAuthorString As String
    Dim rstAuthorCiteForm As Recordset
    Dim iAuthorCount As Integer
    Dim sAuthorString As String
    Dim rstAuthorLast As Recordset
    Dim sAuthorLastString As String
    Dim bLibraryCollection As Boolean
    Dim iRecordsAETID As Integer
    Dim sAETFMLS As String
    Dim rstAETCiteForm As ADODB.Recordset
    Dim sAuthorCiteForm As String
    Dim sEditorCiteForm As String
    Dim iEditorCount As Integer
    Dim report As procs
    Dim sJournalTitle As String
    Dim sJournalTitleShortForm As String
    Dim brepublished As Boolean
    'Dim sDay As String
    Set report = New procs
    
   
    'Dim iAuthorCounter As Integer
    
    'Dim iRecordID As Integer
    Dim dDate As Date
    
    sSourceType = Me.cmbSourceType.Text
    sDateAdded = Me.txtDateAdded.Text
    sDateUpdated = Me.txtDateUpdated.Text
    sInputInitials = Me.txtInputInitials.Text
    sTitle = Me.txtTitle.Text
    sYear = Me.txtYear.Text
    iLargerWorkID = Val(Me.txtLargerWorkID.Text)
    sArticleDesignation = Me.cmbArticleDesignation.Text
    iJournalID = Val(Me.txtJournalID.Text)
    sPublicationDay = Me.txtPublicationDay.Text
    sPageNumber = Me.txtPage.Text
    sVolume = Me.txtVolume.Text
    sCallNumber = Me.txtCallNumber.Text
    sPublicationMonth = Me.cmbPublicationMonthOrSeason.Text
    'sNotes = Me.txtNotes.Text
    sURL = Me.txtURL.Text
    sEditionAndPrinting = Me.txtEditionandPrinting.Text
    sPublisher = Me.txtPublisher.Text
    sOriginalPublicationDate = Me.txtOriginalPublicationDate.Text
    sSeriesVolume = Me.txtSeriesVolume.Text
    sTitleOfSeriesIfNotIssuedByAuthor = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
    bAllChaptersBySameAuthor = Me.chkAllChaptersBySameAuthor.Value
    sUnpublishedWorkType = Me.cmbUnpublishedType.Text
    sThesisDissertationType = Me.cmbThesisDissertationType.Text
    sLocation = Me.txtLocation.Text
    sWorkingPaperInfo = Me.txtWorkingPaper.Text
    sMiscellaneousType = Me.cmbMiscType.Text
    sLegislativeType = Me.cmbLegislativeType.Text
    sNameOfHouse = Me.txtLegislativeHouse
    sNumberOfCongress = Me.txtNumberOfCongress.Text
    sSessionOfCongress = Me.txtSessionOfCongress.Text
    sStateLegislativeSession = Me.txtStateLegislativeSession.Text
    sUSCCANCitation = Me.txtUSCCANCitation.Text
    sReportDocumentNumber = Me.txtReportOrDocumentNumber.Text
    sSuDocNumber = Me.txtSuDocNumber.Text
    sJournalTitle = Me.cmbJournalTitle.Text
    sJournalTitleShortForm = Me.txtJournaTitleShortForm.Text
    bLibraryCollection = Me.chkLibraryCollection
    brepublished = Me.chkRepublished
    
    If (Me.txtTitle.Text = "") Or (Me.cmbSourceType = "") Then
        MsgBox "Some required fields were left blank."
        GoTo CancelErr
        
    End If
    
    Select Case Me.cmbSourceType.Text
        Case "Journal Article"
            If (Me.cmbJournalTitle = "") Then
                MsgBox "Some required fields were left blank."
                GoTo CancelErr
                
            End If
            
        Case "Nonprint Material"
            If (Me.cmbMiscType = "") Then
            MsgBox "Some required fields were left blank."
                GoTo CancelErr
                
            End If
    End Select
    If Me.txtArticleID.Text <> "" Then iArticleID = Me.txtArticleID.Text
    If Me.txtLegislativeID.Text <> "" Then iLegislativeID = Me.txtLegislativeID.Text
    If Me.txtChapterID.Text <> "" Then iChapterID = Me.txtChapterID.Text
    If Me.txtTreatiseID.Text <> "" Then iTreatiseID = Me.txtTreatiseID.Text
    If Me.txtUnpublishedID.Text <> "" Then iUnpublishedID = Me.txtUnpublishedID.Text
    If Me.txtMiscID.Text <> "" Then iMiscID = Me.txtMiscID.Text
    If IsNumeric(Me.cmbRecordNumber.Text) Then iRecordID = Me.cmbRecordNumber.Text
    'sCheckTitle = Replace(sTitle, "'", "%")
    bDuplicate = False
    
    ' GoTo SaveErr
'Check to see if this is a duplicate entry
    If Me.tglNewRecords = True Then
        Set rstCheck = New ADODB.Recordset
        sCheckString = "SELECT * FROM tblrecords WHERE (PublicationYear='" & sYear & "')"
            
        If sPageNumber <> "" Then sCheckString = sCheckString & " AND (PageNumber = '" & sPageNumber & "')"
        '.CursorType = adOpenForwardOnly
        '.LockType = adLockReadOnly
        rstCheck.CursorLocation = adUseClient
        rstCheck.Open sCheckString, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
        Do While Not rstCheck.EOF
            If sTitle = rstCheck!Title Then bDuplicate = True
            rstCheck.MoveNext
        Loop
        If bDuplicate Then
            MsgBox "Duplicate Record Exists. Cannot Save.", vbOKOnly + vbCritical, "Saving Error"
            
            rstCheck.Close
            Set rstCheck = Nothing
            Exit Sub
        End If
        
        rstCheck.Close
        Set rstCheck = Nothing
    End If
    
    dDate = Now
    On Error GoTo CancelErr
    cnWriteDatabase.BeginTrans
        If Me.tglNewRecords.Value = True Then rstRecords.AddNew
            
            If sDateAdded <> "" Then rstRecords!DateRecordAdded = sDateAdded
            If (Me.tglUpdateRecords.Value = True) Or (Me.tglImportRecords.Value = True) Then rstRecords!dateRecordUpdated = dDate
            If sInputInitials <> "" Then rstRecords!InputInitials = sInputInitials
            If sSourceType <> "" Then rstRecords!DocumentType = sSourceType
            If sTitle <> "" Then rstRecords!Title = sTitle
            If sPageNumber <> "" Then rstRecords!PageNumber = sPageNumber
            If sYear <> "" Then rstRecords!PublicationYear = sYear
            'If sNotes <> "" Then rstRecords!Notes = sNotes
            If sURL <> "" Then
                rstRecords!ParallelURL = sURL
            Else
                rstRecords!ParallelURL = Null
            End If
            rstRecords!LibraryCOllection = Me.chkLibraryCollection
            rstRecords!Republished = Me.chkRepublished
            
            
        rstRecords.Update
    cnWriteDatabase.CommitTrans
    If Me.tglNewRecords = True Then iRecordID = rstRecords.Fields("RecordID")
    Select Case sSourceType
    
        Case "Chapter in Treatise"
            Set rstChapters = New ADODB.Recordset
            With rstChapters
                .ActiveConnection = cnWriteDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblChapters")
            End With
    
            cnWriteDatabase.BeginTrans
                If Me.tglNewRecords = True Then rstChapters.AddNew
    
                If Me.tglNewRecords.Value = False Then
                    rstChapters.MoveFirst
                    Do Until rstChapters!RecordID = iRecordID
                        rstChapters.MoveNext
                    Loop
                End If
                If Not rstChapters.EOF Then
                    'If Me.tglUpdateRecords = True Then rstChapters!chapterID = iChapterID
            
                    rstChapters!RecordID = iRecordID
                    rstChapters!LargerWorkID = iLargerWorkID
                    If sSeriesVolume <> "" Then
                        rstChapters!SeriesVolume = sSeriesVolume
                    Else
                        rstChapters!SeriesVolume = Null
                    End If
                    
                    If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then
                        rstChapters!TitleOfSeriesIfNotIssuedByAuthor = sTitleOfSeriesIfNotIssuedByAuthor
                    Else
                        rstChapters!TitleOfSeriesIfNotIssuedByAuthor = Null
                    End If
                    
                End If
                rstChapters.Update
            cnWriteDatabase.CommitTrans
            rstChapters.Close
            Set rstChapters = Nothing
        Case "Journal Article"
            cnWriteDatabase.BeginTrans
            
            Set rstArticles = New ADODB.Recordset
            With rstArticles
                .ActiveConnection = cnWriteDatabase
                .CursorType = adOpenKeyset
                .CursorLocation = adUseClient
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblArticles")
            End With

    
            dDate = Now
                If Me.tglNewRecords.Value = True Then rstArticles.AddNew
                If Me.tglNewRecords.Value = False Then
                    rstArticles.MoveFirst
                    Do Until rstArticles!RecordID = iRecordID
                        rstArticles.MoveNext
                    Loop
                End If
                If Not rstArticles.EOF Then
                    rstArticles!RecordID = iRecordID
                    'rstArticles!recordID = rstRecords!recordID
                    
                    'If Me.tglUpdateRecords = True Then rstArticles!articleID = iArticleID
            
                    If sVolume <> "" Then
                        rstArticles!Volume = sVolume
                    Else
                        rstArticles!Volume = Null
                    End If
                    If sPublicationMonth <> "" Then
                        rstArticles!PublicationMonthOrSeason = sPublicationMonth
                    Else
                        rstArticles!PublicationMonthOrSeason = Null
                    End If
                    If sPublicationDay <> "" Then
                        rstArticles!PublicationDay = sPublicationDay
                    Else
                        rstArticles!PublicationDay = Null
                    End If
                    If sArticleDesignation <> "" Then
                        rstArticles!ArticleDesignationForCitation = sArticleDesignation
                    Else
                        rstArticles!ArticleDesignationForCitation = Null
                    End If
                    
                    rstArticles!JournalID = iJournalID
                End If
                rstArticles.Update
            cnWriteDatabase.CommitTrans
            rstArticles.Close
        Case "Legislative Material"
            Set rstLegislativeMaterial = New ADODB.Recordset
    
            With rstLegislativeMaterial
                .ActiveConnection = cnWriteDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblLegislativeMaterial")
            End With
        
            dDate = Now
                cnWriteDatabase.BeginTrans
                If Me.tglNewRecords = True Then rstLegislativeMaterial.AddNew
                If Me.tglNewRecords.Value = False Then
                    rstLegislativeMaterial.MoveFirst
                    Do Until rstLegislativeMaterial!RecordID = iRecordID
                        rstLegislativeMaterial.MoveNext
                    Loop
                End If
                If Not rstLegislativeMaterial.EOF Then
                    rstLegislativeMaterial!RecordID = iRecordID
                    'If Me.tglUpdateRecords = True Then rstChapters!chapterID = iChapterID
            
                    rstLegislativeMaterial!materialtype = sLegislativeType
                    If sNameOfHouse <> "" Then rstLegislativeMaterial!NameOfHouse = sNameOfHouse
                    If sLegislativeType <> "" Then rstLegislativeMaterial!materialtype = sLegislativeType
                    If sNumberOfCongress <> "" Then rstLegislativeMaterial!NumberOfCongress = sNumberOfCongress
                    If sSessionOfCongress <> "" Then rstLegislativeMaterial!SessionOfCongress = sSessionOfCongress
                    If sStateLegislativeSession <> "" Then rstLegislativeMaterial!StateLegislativeSession = sStateLegislativeSession
                    If sUSCCANCitation <> "" Then rstLegislativeMaterial!USCCANCitation = sUSCCANCitation
                    If sReportDocumentNumber <> "" Then rstLegislativeMaterial!ReportOrDocumentNumber = sReportDocumentNumber
                    If sSuDocNumber <> "" Then rstLegislativeMaterial!SuDocNumber = sSuDocNumber
                End If
                rstLegislativeMaterial.Update
            cnWriteDatabase.CommitTrans
            rstLegislativeMaterial.Close
            Set rstLegislativeMaterial = Nothing
            
        Case "Treatise"
            Set rstTreatises = New ADODB.Recordset

            With rstTreatises
                .ActiveConnection = cnWriteDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblTreatises")
            End With
    
            cnWriteDatabase.BeginTrans
                If Me.tglNewRecords.Value = True Then rstTreatises.AddNew
                If Me.tglNewRecords.Value = False Then
                    rstTreatises.MoveFirst
                    Do Until rstTreatises!RecordID = iRecordID
                        rstTreatises.MoveNext
                    Loop
                End If
                If Not rstTreatises.EOF Then
                'If Me.tglNewRecords = True Then rstTreatises.AddNew
                    rstTreatises!RecordID = iRecordID
                    If sEditionAndPrinting <> "" Then
                        rstTreatises!EditionAndPrinting = sEditionAndPrinting
                    Else
                        rstTreatises!EditionAndPrinting = Null
                    End If
                                                            
                    If sPublisher <> "" Then
                        rstTreatises!Publisher = sPublisher
                    Else
                        rstTreatises!Publisher = Null
                    End If
                    
                    
                    If sOriginalPublicationDate <> "" Then
                        rstTreatises!OriginalPublicationDate = sOriginalPublicationDate
                    Else
                        rstTreatises!OriginalPublicationDate = Null
                    End If
                    
                    
                    If sSeriesVolume <> "" Then
                        rstTreatises!SeriesVolume = sSeriesVolume
                    Else
                        rstTreatises!SeriesVolume = Null
                    End If
                    
                    
                    If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then
                        rstTreatises!TitleOfSeriesIfNotIssuedByAuthor = sTitleOfSeriesIfNotIssuedByAuthor
                    Else
                        rstTreatises!TitleOfSeriesIfNotIssuedByAuthor = Null
                    End If
                    
                    
                    If sCallNumber <> "" Then
                        rstTreatises!CallNumber = sCallNumber
                    Else
                        rstTreatises!CallNumber = Null
                    End If
                    
                
                End If
                rstTreatises.Update
            cnWriteDatabase.CommitTrans
            rstTreatises.Close
            Set rstTreatises = Nothing
            
        Case "Unpublished Work"
            Set rstUnpublishedWork = New ADODB.Recordset
            With rstUnpublishedWork
                .ActiveConnection = cnWriteDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblUnpublishedWork")
            End With
            
            cnWriteDatabase.BeginTrans
                If Me.tglNewRecords = True Then rstUnpublishedWork.AddNew
                If Me.tglNewRecords.Value = False Then
                    rstUnpublishedWork.MoveFirst
                    Do Until rstUnpublishedWork!RecordID = iRecordID
                        rstUnpublishedWork.MoveNext
                    Loop
                End If
                If Not rstUnpublishedWork.EOF Then
                    rstUnpublishedWork!RecordID = rstRecords!RecordID
                    If sUnpublishedWorkType <> "" Then rstUnpublishedWork!Type = sUnpublishedWorkType
                    If sThesisDissertationType <> "" Then rstUnpublishedWork.Fields("Thesis/Dissertation Type") = sThesisDissertationType
                    If sPublicationMonth <> "" Then
                        rstUnpublishedWork!PublicationMonth = sPublicationMonth
                    Else
                        rstUnpublishedWork!PublicationMonth = Null
                    End If
                    
                    If sPublicationDay <> "" Then
                        rstUnpublishedWork!PublicationDay = sPublicationDay
                    Else
                        rstUnpublishedWork!PublicationMonth = Null
                    End If
                    
                    If sLocation <> "" Then
                        rstUnpublishedWork!Location = sLocation
                     Else
                        rstUnpublishedWork!Location = Null
                    End If
                End If
                rstUnpublishedWork.Update
            cnWriteDatabase.CommitTrans
            rstUnpublishedWork.Close
            Set rstUnpublishedWork = Nothing
            
        Case "Nonprint Material"
            Set rstMisc = New ADODB.Recordset
            With rstMisc
                .ActiveConnection = cnWriteDatabase
                .CursorType = adOpenKeyset
                .CursorLocation = adUseClient
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblMisc")
            End With
            cnWriteDatabase.BeginTrans
                
                If ((Me.tglNewRecords = True)) Then rstMisc.AddNew
                If Me.tglNewRecords.Value = False Then
                    rstMisc.MoveFirst
                    Do Until rstMisc!RecordID = iRecordID
                        rstMisc.MoveNext
                    Loop
                End If
                If Not rstMisc.EOF Then
                    rstMisc!RecordID = iRecordID
                    
                    
                    If sMiscellaneousType <> "" Then rstMisc!RecordType = sMiscellaneousType
                    If sLocation <> "" Then
                        rstMisc!Location = sLocation
                    Else
                        rstMisc!Location = Null
                    End If
                    
                    If sWorkingPaperInfo <> "" Then
                        rstMisc!WorkingPaper = sWorkingPaperInfo
                    Else
                        rstMisc!WorkingPaper = Null
                    End If
                    
                    If sPublicationMonth <> "" Then
                        rstMisc!Month = sPublicationMonth
                    Else
                        rstMisc!Month = Null
                    End If
                    
                    If sPublicationDay <> "" Then
                        rstMisc!Day = sPublicationDay
                    Else
                        rstMisc!Day = Null
                    End If
                        
                End If
                rstMisc.Update
            cnWriteDatabase.CommitTrans
            rstMisc.Close
            Set rstMisc = Nothing
    End Select
    
    If (Me.tglUpdateRecords.Value = True) Or (Me.tglImportRecords.Value = True) Then
        Set rstAETDelete = New Recordset
        Set rstKeywordDelete = New Recordset
        rstAETDelete.CursorLocation = adUseClient
        rstKeywordDelete.CursorLocation = adUseClient
        rstAETDelete.Open "Select * from tblRecordsAET WHERE RecordID=" & iRecordID, cnWriteDatabase, adOpenKeyset, adLockOptimistic
        rstKeywordDelete.Open "Select * from tblRecordsKeywords WHERE RecordID=" & iRecordID, cnWriteDatabase, adOpenKeyset, adLockOptimistic
        Do While Not rstAETDelete.EOF
            rstAETDelete.Delete
            rstAETDelete.Update
            rstAETDelete.MoveNext
        Loop
        
        Do While Not rstKeywordDelete.EOF
            rstKeywordDelete.Delete
            rstKeywordDelete.Update
            rstKeywordDelete.MoveNext
        Loop
        
        rstAETDelete.Close
        rstKeywordDelete.Close
        Set rstAETDelete = Nothing
        Set rstKeywordDelete = Nothing

    End If
    
    
    'If Me.tglNewRecords.Value = True Then
        Set rstRecordsAET = New ADODB.Recordset
        With rstRecordsAET
            .ActiveConnection = cnWriteDatabase
            .CursorType = adOpenKeyset
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsAET")
        End With
        rstRecordsAET.MoveLast
        'iRecordsAETID = rstRecordsAET!RecordsAETID
        
        For icounter = 1 To cAuthors.Count
        '    iRecordsAETID = iRecordsAETID + 1
            rstRecordsAET.AddNew
                rstRecordsAET!RecordID = iRecordID
                rstRecordsAET!AETID = cAuthors.Item(icounter)
        '        rstRecordsAET!RecordsAETID = iRecordsAETID
            rstRecordsAET.Update
        Next
        
        
        For icounter = 1 To cEditors.Count
        '    iRecordsAETID = iRecordsAETID + 1
            rstRecordsAET.AddNew
                rstRecordsAET!RecordID = iRecordID
                rstRecordsAET!AETID = cEditors.Item(icounter)
        '        rstRecordsAET!RecordsAETID = iRecordsAETID

            rstRecordsAET.Update
        Next
        
        For icounter = 1 To cTranslators.Count
        '    iRecordsAETID = iRecordsAETID + 1
            rstRecordsAET.AddNew
                rstRecordsAET!RecordID = iRecordID
                rstRecordsAET!AETID = cTranslators.Item(icounter)
        '        rstRecordsAET!RecordsAETID = iRecordsAETID
            rstRecordsAET.Update
        Next
               
        rstRecordsAET.Close
        Set rstRecordsAET = Nothing
        
        Set rstRecordsKeywords = New ADODB.Recordset
        With rstRecordsKeywords
            .ActiveConnection = cnWriteDatabase
            .CursorType = adOpenKeyset
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords")
        End With
            
        
        For icounter = 1 To cKeywords.Count
            rstRecordsKeywords.AddNew
                rstRecordsKeywords!RecordID = iRecordID
                rstRecordsKeywords!KeywordID = cKeywords.Item(icounter)
            rstRecordsKeywords.Update
        Next
        rstRecordsKeywords.Close
        Set rstRecordsKeywords = Nothing
        
'keyword subpart
    sSQLString = "select * from qryRecordsKeywordsThesaurus where RecordID=" & iRecordID
    Set rstRecordsKeywordsThesaurus = New Recordset
    Set rstAllKeyword = New Recordset
    Set cAllKeywords = New Collection
    rstRecordsKeywordsThesaurus.CursorLocation = adUseClient
    rstRecordsKeywordsThesaurus.Open sSQLString, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    sSQLString = "select * from tblRecordsAllKeywords where RecordID=" & iRecordID
    rstAllKeyword.CursorLocation = adUseClient
    rstAllKeyword.Open sSQLString, cnWriteDatabase, adOpenKeyset, adLockOptimistic
    On Error GoTo KeywordSaveErr
    Do While Not rstRecordsKeywordsThesaurus.EOF
        'iRecordID = rstRecordsKeywordsThesaurus!RecordID
        sAllKeywordString = ""

        Do While iRecordID = rstRecordsKeywordsThesaurus!RecordID
            bDuplicate = False
            If rstRecordsKeywordsThesaurus!KeywordOrCodeSection <> "" Then
                bDuplicate = False
                sCurrentKeyword = rstRecordsKeywordsThesaurus!KeywordOrCodeSection
                For i = 1 To cAllKeywords.Count
                    If cAllKeywords.Item(i) = sCurrentKeyword Then bDuplicate = True
                Next
                If Not bDuplicate Then cAllKeywords.Add sCurrentKeyword
            End If
    '
            bDuplicate = False
    '
            If rstRecordsKeywordsThesaurus!ThesaurusEquivalent <> "" Then
                bDuplicate = False
                sCurrentKeyword = rstRecordsKeywordsThesaurus!ThesaurusEquivalent
                For i = 1 To cAllKeywords.Count
                    If cAllKeywords.Item(i) = sCurrentKeyword Then bDuplicate = True
                Next
                If Not bDuplicate Then cAllKeywords.Add sCurrentKeyword
            End If
            rstRecordsKeywordsThesaurus.MoveNext
        Loop
    '
KeywordEOFErr:
        For i = 1 To cAllKeywords.Count
            If sAllKeywordString <> "" Then sAllKeywordString = sAllKeywordString & " "
            sAllKeywordString = sAllKeywordString & cAllKeywords.Item(i)
        Next
    '
        cnWriteDatabase.BeginTrans
        
        If rstAllKeyword.EOF Then rstAllKeyword.AddNew
            rstAllKeyword!RecordID = iRecordID
            rstAllKeyword!AllKeywords = sAllKeywordString
        rstAllKeyword.Update
        cnWriteDatabase.CommitTrans
        
        'Me.lblRecNum.Caption = "Record No. " & iRecordID & " processed."
        'Me.lblRecNum.Refresh
    Loop
    If (Not rstAllKeyword.EOF) And cAllKeywords.Count = 0 Then
        cnWriteDatabase.BeginTrans
            rstAllKeyword.Delete
            rstAllKeyword.Update
        cnWriteDatabase.CommitTrans
    End If
    Set cAllKeywords = Nothing
    '
    'cnDatabase.Close
    'Set cnDatabase = Nothing
    'Set rstRecordsAuthors = Nothing
    Set rstRecordsKeywordsThesaurus = Nothing
    'Set rstAllAuthor = Nothing
    Set rstAllKeyword = Nothing
'end keyword part
    
    
'author subpart
    sSQLString = "select * from qryAuthors where RecordID=" & iRecordID
    Set rstRecordsAuthors = New Recordset
    rstRecordsAuthors.CursorLocation = adUseClient
    rstRecordsAuthors.Open sSQLString, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    
    On Error GoTo AuthorSaveErr0
    
    sAuthorString = ""
    sAuthorLastString = ""
    If Not rstRecordsAuthors.EOF Then
        rstRecordsAuthors.MoveLast
        iAuthorCount = rstRecordsAuthors.RecordCount
        rstRecordsAuthors.MoveFirst
        sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
        Select Case iAuthorCount
            Case 1
                If rstRecordsAuthors!InstitutionalEntity <> "" Then
                    sAuthorString = rstRecordsAuthors!InstitutionalEntity
                    If sAETFMLS <> "" Then sAuthorString = sAuthorString & ", " & sAETFMLS
                Else
                    If sAETFMLS <> "" Then sAuthorString = sAETFMLS
                End If
                If rstRecordsAuthors!LastName <> "" Then sAuthorLastString = rstRecordsAuthors!LastName
            Case 2
                    If rstRecordsAuthors!InstitutionalEntity <> "" Then sAuthorString = rstRecordsAuthors!InstitutionalEntity & ","
                    If sAETFMLS <> "" Then sAuthorString = sAETFMLS
                    If rstRecordsAuthors!LastName <> "" Then sAuthorLastString = rstRecordsAuthors!LastName
                    rstRecordsAuthors.MoveNext
                    sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
                    sAuthorString = sAuthorString & " & "
                    sAuthorLastString = sAuthorLastString & " " & rstRecordsAuthors!LastName
                    
                    
                    sAuthorString = sAuthorString & sAETFMLS
                    
            Case Else
                    If sAETFMLS <> "" Then sAuthorString = sAETFMLS & " et al."
                    If rstRecordsAuthors!InstitutionalEntity <> "" Then sAuthorString = rstRecordsAuthors!InstitutionalEntity & ","
                    If rstRecordsAuthors!LastName <> "" Then sAuthorLastString = rstRecordsAuthors!LastName
                    rstRecordsAuthors.MoveNext
                    Do While Not rstRecordsAuthors.EOF
                        sAuthorLastString = sAuthorLastString & " " & rstRecordsAuthors!LastName
                        rstRecordsAuthors.MoveNext
                        
                    Loop
        End Select
     End If
     
     ''here we can perhaps check to see if published in a different source **republished
     
     ''check to see if republished *************ADD SECTION LATER***************
''look at qryCheckRepublished. Check to see if title is the same, and create a string variable of author
''last names and check against the field in the query. Check to see if "republished" is checked before checking.
''if there is a match in title and author from query, send user a message alert telling them to check the Record ID
''and see if there is a match (because it might be a book edition or something else)
''Even better: pop up a citation of the other record and ask if it is a republished source. If it is, mark both to republished

'causing error now because Title can have a quote character in it and it messes up the SQL query. Commenting out for now.
   'If Me.tglNewRecords = True Then
   '     Set rstCheck = New ADODB.Recordset
   '     sCheckString = "SELECT * FROM qryCheckRepublished WHERE (Title='" & sTitle & "')"
   '     sCheckString = sCheckString & " AND (AllAuthorLastNameOnly = '" & sAuthorLastString & "')"
   '     rstCheck.CursorType = adOpenForwardOnly
   '     rstCheck.LockType = adLockReadOnly
   '     rstCheck.CursorLocation = adUseClient
   '     rstCheck.Open sCheckString, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
   '     bDuplicate = False
   '     If Not rstCheck.EOF Then bDuplicate = True
   '     If bDuplicate Then MsgBox "This work might be republished. Please verify, go back and update records to check appropriate checkboxes if true.", vbOKOnly + vbInformation, "Alert"
   '
   '     rstCheck.Close
   '     Set rstCheck = Nothing
   'End If



     
AuthorEOFErr0:
    Call report.Get_AET_String(iRecordID, Me.cnReadDatabase, sAuthorCiteForm, sEditorCiteForm, cAuthors.Count, cEditors.Count)
    
    Set rstAETCiteForm = New ADODB.Recordset
    
    With rstAETCiteForm
        .ActiveConnection = cnWriteDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open "Select * from tblRecordsAETCiteForm WHERE RecordID=" & iRecordID
    End With
    cnWriteDatabase.BeginTrans

        If rstAETCiteForm.EOF Then rstAETCiteForm.AddNew
        rstAETCiteForm!RecordID = iRecordID
        rstAETCiteForm!authorciteform = sAuthorCiteForm
        rstAETCiteForm!Editorciteform = sEditorCiteForm
        
        rstAETCiteForm.Update
    
    cnWriteDatabase.CommitTrans

    
    rstAETCiteForm.Close
    Set rstAETCiteForm = Nothing

    sSQLString = "select * from tblRecordsAuthorCiteForm where RecordID=" & iRecordID
    Set rstAuthorCiteForm = New Recordset
    rstAuthorCiteForm.CursorLocation = adUseClient
    rstAuthorCiteForm.Open sSQLString, cnWriteDatabase, adOpenKeyset, adLockOptimistic
   
      Select Case sAuthorString
          Case ""
              If Not rstAuthorCiteForm.EOF Then
                cnWriteDatabase.BeginTrans
    
                rstAuthorCiteForm.Delete
                rstAuthorCiteForm.Update
                
                cnWriteDatabase.CommitTrans
              End If
          Case Else
              cnWriteDatabase.BeginTrans
    
              If rstAuthorCiteForm.EOF Then rstAuthorCiteForm.AddNew
              rstAuthorCiteForm!RecordID = iRecordID
              rstAuthorCiteForm!authorciteform = sAuthorString
              rstAuthorCiteForm.Update
              
              cnWriteDatabase.CommitTrans

      End Select
    rstAuthorCiteForm.Close
    Set rstAuthorCiteForm = Nothing
    On Error GoTo AuthorSaveErr1

authorEOFErr1:
    sSQLString = "select * from tblRecordsAllAuthorLastNameOnly where RecordID=" & iRecordID
    Set rstAuthorLast = New Recordset
    rstAuthorLast.CursorLocation = adUseClient
    rstAuthorLast.Open sSQLString, cnWriteDatabase, adOpenKeyset, adLockOptimistic
   
      Select Case sAuthorLastString
          Case ""
              If Not rstAuthorLast.EOF Then
                cnWriteDatabase.BeginTrans
    
                rstAuthorLast.Delete
                rstAuthorLast.Update
                
                cnWriteDatabase.CommitTrans

              End If
          
          Case Else
              cnWriteDatabase.BeginTrans
    
              If rstAuthorLast.EOF Then rstAuthorLast.AddNew
              rstAuthorLast!RecordID = iRecordID
              rstAuthorLast!AllAuthorLastNameOnly = sAuthorLastString
              rstAuthorLast.Update
              
              cnWriteDatabase.CommitTrans

        End Select
        'rstAuthorLast.Update
    'cnDatabase.CommitTrans
    rstAuthorLast.Close
    Set rstAuthorLast = Nothing
    'If Not rstRecordsAuthors.EOF Then rstRecordsAuthors.MoveFirst
    If rstRecordsAuthors.RecordCount > 0 Then rstRecordsAuthors.MoveFirst
    
    On Error GoTo AuthorSaveErr
    'Set rstRecordsAuthors = Nothing

    'iRecordID = rstRecordsAuthors!RecordID
    sFullAuthorString = ""
    Do While iRecordID = rstRecordsAuthors!RecordID
        If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
        If rstRecordsAuthors!InstitutionalEntity <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!InstitutionalEntity
        If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
        sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
        If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
        'If rstRecordsAuthors!FMLS <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!FMLS
        sAETFMLS = Full_AET(rstRecordsAuthors, "FL")
        If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
        If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
        'If rstRecordsAuthors!FL <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!FL
        sAETFMLS = Full_AET(rstRecordsAuthors, "LFM")
        If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
        If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
        'If rstRecordsAuthors!LFM <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!LFM
        rstRecordsAuthors.MoveNext
        
    Loop
AuthorEOFErr:
    sSQLString = "select * from tblRecordAllAuthor where RecordID=" & iRecordID
    Set rstAllAuthor = New ADODB.Recordset
    rstAllAuthor.CursorLocation = adUseClient
    rstAllAuthor.Open sSQLString, cnWriteDatabase, adOpenKeyset, adLockOptimistic
   
        Select Case sFullAuthorString
            Case ""
                If Not rstAllAuthor.EOF Then
                  cnWriteDatabase.BeginTrans
      
                    rstAllAuthor.Delete
                    rstAllAuthor.Update
                  cnWriteDatabase.CommitTrans
                End If
            Case Else
                cnWriteDatabase.BeginTrans
    
                
                If rstAllAuthor.EOF Then rstAllAuthor.AddNew
                rstAllAuthor!RecordID = iRecordID
                rstAllAuthor!AllAuthors = sFullAuthorString
                rstAllAuthor.Update
                cnWriteDatabase.CommitTrans
        End Select
    rstAllAuthor.Close
    Set rstAllAuthor = Nothing
NoAuthor:
'end author part
    Set rstBigRecordIndex = New ADODB.Recordset
    With rstBigRecordIndex
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = cnWriteDatabase
        .Open "Select * from tblBigTextIndex WHERE RecordID=" & iRecordID
    End With
    cnWriteDatabase.BeginTrans
        If rstBigRecordIndex.EOF Then rstBigRecordIndex.AddNew
        rstBigRecordIndex!RecordID = iRecordID
        rstBigRecordIndex!Title = sTitle
        rstBigRecordIndex!AllAuthors = sFullAuthorString
        rstBigRecordIndex!AllAuthorLastNameOnly = sAuthorLastString
        rstBigRecordIndex!AllKeywords = sAllKeywordString
        rstBigRecordIndex!JournalTitle = sJournalTitle
        rstBigRecordIndex!JournalTitleShortFOrm = sJournalTitleShortForm
        rstBigRecordIndex.Update
    cnWriteDatabase.CommitTrans
    
    rstBigRecordIndex.Close
    Set rstBigRecordIndex = Nothing
    If Me.tglNewRecords.Value = True Then
        Me.cmbRecordNumber.RemoveItem (Me.cmbRecordNumber.ListCount - 1)
        Me.cmbRecordNumber.AddItem iRecordID
        Me.cmbRecordNumber.AddItem "New Record"
        Call Set_Entry_Form
    End If
    Me.txtStatus.Text = "Saved"
    rstRecords.Requery
    If Me.tglUpdateRecords = True Then
       If iRecordID <> rstRecords.Fields("RecordID").Value Then
            rstRecords.MoveFirst
            Do Until rstRecords!RecordID = iRecordID
                rstRecords.MoveNext
            Loop
        End If
    End If
    'cnDatabase.Close
    'sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\database\ncpl.mdb"
    'cnDatabase.Open (sConnectionString)
    'rstJournals.Requery
    'rstAuthors.Requery
    'rstEditors.Requery
    'rstTranslators.Requery
    'rstArticles.Requery
    'rstChapters.Requery
    'rstMisc.Requery
    'rstLegislativeMaterial.Requery
    'rstRecords.Requery
    'rstRecordsAET.Requery
    'rstRecordsKeywords.Requery
    'rstTreatises.Requery
    'rstUnpublishedWork.Requery

KeywordSaveErr:
        Select Case Err
        Case 3021
            Resume KeywordEOFErr
        Case 0
            
        Case Else
            cnWriteDatabase.RollbackTrans
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Saving Error"
        End Select
AuthorSaveErr:
        Select Case Err
        Case 3021
            Resume AuthorEOFErr
        
        Case 0
            
        Case Else
            cnWriteDatabase.RollbackTrans
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Saving Error"
        End Select
AuthorSaveErr0:
        Select Case Err
        Case 3021
            Resume AuthorEOFErr0
        
        Case 0
            
        Case Else
            cnWriteDatabase.RollbackTrans
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Saving Error"
        End Select
AuthorSaveErr1:
        Select Case Err
        Case 3021
            Resume authorEOFErr1
        Case 0
        Case Else
            cnWriteDatabase.RollbackTrans
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Saving Error"
        End Select
       
CancelErr:
    Select Case Err
        Case 0
        Case Else
            cnWriteDatabase.RollbackTrans
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Saving Error"
    End Select
End Sub




Private Sub Form_Load()
    Dim sConnectionString As String
    Dim sRemoteConnectionString As String
    
    Dim dDate As Date
    'Set rstJournals = New ADODB.Recordset
    'Set rstAuthors = New ADODB.Recordset
    'Set rstEditors = New ADODB.Recordset
    'Set rstTranslators = New ADODB.Recordset
    'Set rstArticles = New ADODB.Recordset
    'Set rstChapters = New ADODB.Recordset
    'Set rstMisc = New ADODB.Recordset
    'Set rstLegislativeMaterial = New ADODB.Recordset
    'Set rstRecords = New ADODB.Recordset
    'Set rstRecordsAET = New ADODB.Recordset
    'Set rstRecordsKeywords = New ADODB.Recordset
    'Set rstTreatises = New ADODB.Recordset
    'Set rstUnpublishedWork = New ADODB.Recordset

    'Set rstLargerWorks = New ADODB.Recordset
    'Set rstKeywords = New ADODB.Recordset
    Set cnReadDatabase = New ADODB.Connection
    Set cnWriteDatabase = New ADODB.Connection
    Set cnRemoteReadDatabase = New ADODB.Connection
    Set cnRemoteWriteDatabase = New ADODB.Connection
    
    
    
    Set cAuthors = New Collection
    Set cEditors = New Collection
    Set cTranslators = New Collection
    Set cKeywords = New Collection
    
    sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLLive;Data Source=NCPL" 'for local access
    sRemoteConnectionString = "Provider=SQLOLEDB.1;Data Source=awssqldev.nyulaw.me;Initial Catalog=NCPLLive;User Id=barnesw;Password=philly"
    'sRemoteConnectionString = "Provider=MSDASQL; DRIVER=Sql Server;Server=awssqldev.nyulaw.me;Database=Ncpl;User Id=barnesw;Password=philly"
    'Provider=MSDASQL; DRIVER=Sql Server; SERVER=p42800; DATABASE=myDatabase; UID=MyUserID; PWD=MyPassword
    
    
    'Provider=sqloledb;Data Source=myServerAddress;Initial Catalog=myDataBase;
'User Id=myUsername;Password=myPassword;
    
    
    cnReadDatabase.Open (sConnectionString)
    cnWriteDatabase.Open (sConnectionString)
    cnRemoteReadDatabase.Open (sRemoteConnectionString)
    cnRemoteWriteDatabase.Open (sRemoteConnectionString)
    
    
    Me.cmbSourceType.CausesValidation = False
    Me.tglUpdateRecords = True
    Me.cmbAETChoice = "Author"
    'With rstJournals
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblJournals")
    'End With
    
    'With rstAuthors
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from qryAETRecords WHERE AETType='Author'")
    'End With
    
    'With rstEditors
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from qryAETRecords WHERE AETType='Editor'")
    'End With
    
    'With rstTranslators
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from qryAETRecords WHERE AETType='Translator'")
    'End With
    
    'With rstLargerWorks
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblLargerWorks")
    'End With
    
    'With rstKeywords
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblKeywords")
    'End With
   
    'With rstArticles
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblArticles")
    'End With

    'With rstChapters
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblChapters")
    'End With
    
    'With rstMisc
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblMisc")
    'End With
    
    'With rstLegislativeMaterial
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblLegislativeMaterial")
    'End With
    Set rstRecords = New ADODB.Recordset
    rstRecords.CursorLocation = adUseClient
    With rstRecords
        .ActiveConnection = cnWriteDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblRecords")
    End With
    'rstRecords.MoveFirst
    'With rstRecordsAET
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblRecordsAET")
    'End With
    
    'With rstRecordsKeywords
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblRecordsKeywords")
    'End With
    
    'With rstTreatises
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblTreatises")
    'End With
    
    'With rstUnpublishedWork
    '    .ActiveConnection = cnWriteDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from tblUnpublishedWork")
    'End With
    dDate = Now
    'Me.txtDateAdded = Left(dDate, 7)
    
    Call Populate_Comboboxes
    Call Erase_Form
    If tglUpdateRecords.Value = True Then
        Me.cmbRecordNumber.ListIndex = 0
        'rstRecords.MoveFirst
        'Call Fill_Form
    End If
End Sub

Private Sub Erase_Form()
    Dim sTempJournalTitle As String
    Dim iTempIndex As Integer
    
    Erase_Object lblLargerWorkID
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Chapter in Treatise")) Then Erase_Object txtLargerWorkID Else Erase_Object txtLargerWorkID, True
    
    
    Erase_Object lblArticleDesignation
    Erase_Object cmbArticleDesignation, True
    
    'Erase_Object lblJournalID
    Erase_Object lblJournalTitle
    Erase_Object lblPublicationDay
    Erase_Object lblVolume
    Erase_Object lblPublicationMonthOrSeason
    
    
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object txtJournalID Else Erase_Object txtJournalID, True
    
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object txtPublicationDay Else Erase_Object txtPublicationDay, True
    
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object txtVolume Else Erase_Object txtVolume, True
    
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object cmbPublicationMonthOrSeason Else Erase_Object cmbPublicationMonthOrSeason, True
    
    
    Erase_Object Me.lblCallNumber
    Erase_Object Me.txtCallNumber, True
    
    Erase_Object lblPage
    Erase_Object txtPage, True
    
    Erase_Object cmdEditJournal
    
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
        'sTempJournalTitle = Me.cmbJournalTitle.Text
        iTempIndex = Me.cmbJournalTitle.ListIndex
        Erase_Object cmbJournalTitle
        'Me.cmbJournalTitle.Text = sTempJournalTitle
        Me.cmbJournalTitle.ListIndex = 0
        Me.cmbJournalTitle.ListIndex = iTempIndex
        Erase_Object Me.txtJournaTitleShortForm
        
    Else
        Erase_Object cmbJournalTitle, True
        Erase_Object Me.txtJournaTitleShortForm, True
        
    End If

        
    Erase_Object Me.cmdNewLargerWork
    
    Erase_Object Me.chkKeepSelected
    Erase_Object Me.chkYear
    'Erase_Object lblJournalTitleShortForm
    Erase_Object Me.txtJournaTitleShortForm, True
    
    'Erase_Object lblOrganizationIssuingNewsletter
    'Erase_Object txtOrganizationIssuingNewsletter, True
    
    'Erase_Object lblCallNumber
    'Erase_Object txtCallNumber, True

    'Erase_Object Me.cmdEditLargerWork

    'Erase_Object lblPagination
    Erase_Object cmbPagination, True
    
    'Erase_Object lblNotes
    'Erase_Object txtNotes, True
    
    Erase_Object lblPage
    Erase_Object txtPage, True
    
    'Erase_Object lblPlaceOfPublication
    'Erase_Object txtPlaceOfPublication, True
    
    Erase_Object lblEditionandPrinting
    Erase_Object txtEditionandPrinting, True
    
    Erase_Object lblPublisher
    Erase_Object txtPublisher, True
    
    Erase_Object lblOriginalPublicationDate
    Erase_Object txtOriginalPublicationDate, True
    
    Erase_Object lblTitleOfSeriesIfNotIssuedByAuthor
    Erase_Object txtTitleOfSeriesIfNotIssuedByAuthor, True
    
    Erase_Object lblLocation
    Erase_Object txtLocation, True
    
    Erase_Object lblWorkingPaper
    Erase_Object txtWorkingPaper, True
    
    Erase_Object lblLegislativeHouse
    Erase_Object txtLegislativeHouse, True
    
    Erase_Object lblSeriesVolume
    Erase_Object txtSeriesVolume, True
    
    Erase_Object lblNumberOfCongress
    Erase_Object txtNumberOfCongress, True
    
    Erase_Object lblSessionOfCongress
    Erase_Object txtSessionOfCongress, True
    
    Erase_Object lblStateLegislativeSession
    Erase_Object txtStateLegislativeSession, True
    
    Erase_Object lblUSCCANCitation
    Erase_Object txtUSCCANCitation, True
    
    Erase_Object lblReportOrDocumentNumber
    Erase_Object txtReportOrDocumentNumber, True
    
    Erase_Object lblYear
    If Me.chkYear.Value = False Then Erase_Object txtYear, True Else Erase_Object txtYear
        
    Erase_Object lblSuDocNumber
    Erase_Object txtSuDocNumber, True
    
    Erase_Object lblUnpublishedType
    Erase_Object cmbUnpublishedType, True
    
    Erase_Object lblThesisDissertationType
    Erase_Object cmbThesisDissertationType, True
    
    Erase_Object lblMiscType
    Erase_Object cmbMiscType, True
    
    Erase_Object lblLegislativeType
    Erase_Object cmbLegislativeType, True
    
    Erase_Object lblLargerWorkTitle
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Chapter in Treatise")) Then Erase_Object cmbLargerWorkTitle Else Erase_Object cmbLargerWorkTitle, True
    
    Erase_Object chkAllChaptersBySameAuthor
    
    Erase_Object cmdNewJournal
    
    
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'rstJournals.Close
    'rstAuthors.Close
    'rstEditors.Close
    'rstTranslators.Close
    'rstLargerWorks.Close
    'rstKeywords.Close
    'rstArticles.Close
    'rstChapters.Close
    'rstMisc.Close
    'rstLegislativeMaterial.Close
    rstRecords.Close
    'rstRecordsAET.Close
    'rstRecordsKeywords.Close
    'rstTreatises.Close
    'rstUnpublishedWork.Close
    '
    cnReadDatabase.Close
    cnWriteDatabase.Close
    cnRemoteReadDatabase.Close
    cnRemoteWriteDatabase.Close
    
    
    Set rstJournals = Nothing
    Set rstAuthors = Nothing
    Set rstEditors = Nothing
    Set rstTranslators = Nothing
    Set rstLargerWorks = Nothing
    Set rstKeywords = Nothing
    
    Set cnWriteDatabase = Nothing
    Set cnReadDatabase = Nothing
    Set cnRemoteWriteDatabase = Nothing
    Set cnRemoteReadDatabase = Nothing
    
    
    
    Set rstArticles = Nothing
    Set rstChapters = Nothing
    Set rstMisc = Nothing
    Set rstLegislativeMaterial = Nothing
    Set rstRecords = Nothing
    Set rstRecordsAET = Nothing
    Set rstRecordsKeywords = Nothing
    Set rstTreatises = Nothing
    Set rstUnpublishedWork = Nothing
End Sub



Private Sub lblRecordNumber_Click()
    frmJump.Show
End Sub

Private Sub lstAuthors_DblClick()
    Call Manage_Lists(lstCurrentAuthors, lstAuthors, cAuthors)
    If cAuthors.Count > 0 Then
        If cAuthors.Count = 1 Then lblA.Text = "Author"
        If cAuthors.Count > 1 Then lblA.Text = "Authors"
    Else
        lblA.Text = "No Author"
    End If
        
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub lstCurrentAuthors_DblClick()
    Call Manage_Lists(lstAuthors, lstCurrentAuthors, cAuthors)
    Me.txtStatus.Text = "Not Saved"
    If cAuthors.Count > 0 Then
        If cAuthors.Count = 1 Then lblA.Text = "Author"
        If cAuthors.Count > 1 Then lblA.Text = "Authors"
    Else
        lblA.Text = "No Author"
    End If
        
End Sub
Private Sub lstEditors_DblClick()
    Call Manage_Lists(lstCurrentEditors, lstEditors, cEditors)
    If cEditors.Count > 0 Then
        If cEditors.Count = 1 Then lblE.Text = "Editor"
        If cEditors.Count > 1 Then lblE.Text = "Editors"
    Else
        lblE.Text = "No Editor"
    End If
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub lstCurrentEditors_DblClick()
    Call Manage_Lists(lstEditors, lstCurrentEditors, cEditors)
    If cEditors.Count > 0 Then
            If cEditors.Count = 1 Then lblE.Text = "Editor"
            If cEditors.Count > 1 Then lblE.Text = "Editors"
    Else
            lblE.Text = "No Editor"
    End If
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub lstNewKeywords_DblClick()
    Dim sSelText As String
    Dim iSelected As Integer
    Dim i As Integer
    
    sSelText = Me.lstNewKeywords.Text
    iSelected = Me.lstNewKeywords.ListIndex

    'For i = 0 To Me.lstKeywords.ListCount - 1
        
    'Next
    Me.lstKeywords.Text = sSelText
    Call Manage_Lists(lstCurrentKeywords, lstKeywords, cKeywords)
    Me.lstNewKeywords.RemoveItem (iSelected)
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub lstTranslators_DblClick()
    Call Manage_Lists(lstCurrentTranslators, lstTranslators, cTranslators)
        If cTranslators.Count > 0 Then
            If cTranslators.Count = 1 Then lblT.Text = "Translator"
            If cTranslators.Count > 1 Then lblT.Text = "Translators"
        Else
            lblT.Text = "No Translator"
        End If
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub lstCurrentTranslators_DblClick()
    Call Manage_Lists(lstTranslators, lstCurrentTranslators, cTranslators)
        If cTranslators.Count > 0 Then
            If cTranslators.Count = 1 Then lblT.Text = "Translator"
            If cTranslators.Count > 1 Then lblT.Text = "Translators"
        Else
            lblT.Text = "No Translator"
        End If
    Me.txtStatus.Text = "Not Saved"
End Sub
Private Sub lstKeywords_DblClick()
    Call Manage_Lists(lstCurrentKeywords, lstKeywords, cKeywords)
    Me.txtStatus.Text = "Not Saved"
End Sub
Private Sub lstCurrentKeywords_DblClick()
    Call Manage_Lists(lstKeywords, lstCurrentKeywords, cKeywords)
    Me.txtStatus.Text = "Not Saved"
End Sub

Public Sub Manage_Lists(oAdd As ListBox, oRemove As ListBox, cCollection As Collection, Optional iListindex As Long = 999999)
    Dim sItem As String
    Dim iID As Integer
    
    Dim iParenpos As Integer
    'sItem = oRemove.Text
    If iListindex = 999999 Then iListindex = oRemove.ListIndex
    sItem = oRemove.List(iListindex)
    oAdd.AddItem sItem
    oRemove.RemoveItem (iListindex)
    If Mid(oAdd.Name, 4, 7) = "Current" Then
        iParenpos = InStr(1, sItem, " (ID: ")
        iID = Val(Mid(sItem, iParenpos + 6, (Len(sItem) - (iParenpos + 6))))
    End If
    If Mid(oAdd.Name, 4, 7) = "Current" Then cCollection.Add iID Else _
        cCollection.Remove (iListindex + 1)
End Sub

Private Sub Position_Object(oObject As Object, LeftPos As Integer, TopPos As Integer)

    oObject.Left = LeftPos
    oObject.Top = TopPos
    oObject.Visible = True
    oObject.Enabled = True
End Sub

Private Sub Erase_Object(oObject As Object, Optional bErase As Boolean)
    oObject.Enabled = False
    oObject.Visible = False
    If bErase Then oObject.Text = ""
End Sub

Private Sub Fill_Form()
    Dim rstArticlesJournals As ADODB.Recordset
    Dim rstQryKeywords As ADODB.Recordset
    Dim rstAETLMFRecords As ADODB.Recordset
    Dim rstLargerWorksChapters As ADODB.Recordset
    Dim rstLegislative As ADODB.Recordset
    Dim rstTreatise As ADODB.Recordset
    Dim rstUnpublished As ADODB.Recordset
    Dim rstOther As ADODB.Recordset
    Dim iAETID As Integer
    Dim iKeywordID As Integer
    Dim sSourceType As String
    Dim icounter As Integer
    Dim iRecNum As Integer
    Dim sCurrentAuthor As String
    Dim sCurrentKeyword As String
    
    Dim iListCount As Integer
    Dim sAETType As String
    
    Set rstAETLMFRecords = New ADODB.Recordset
    Set rstQryKeywords = New ADODB.Recordset
    'If rstRecords.EOF Then rstRecords.MoveFirst
    If rstRecords.EOF Then rstRecords.MoveLast

    'Me.cmbRecordNumber.Text = rstRecords!recordid
    Me.txtTitle.Text = rstRecords!Title
    Me.cmbSourceType.Text = rstRecords!DocumentType
    If rstRecords!DateRecordAdded <> "" Then Me.txtDateAdded.Text = rstRecords!DateRecordAdded
    If rstRecords!dateRecordUpdated <> "" Then Me.txtDateUpdated.Text = rstRecords!dateRecordUpdated
    If rstRecords!InputInitials <> "" Then Me.txtInputInitials = rstRecords!InputInitials
    If rstRecords!PageNumber <> "" Then Me.txtPage = rstRecords!PageNumber
    If rstRecords!PublicationYear <> "" Then Me.txtYear = rstRecords!PublicationYear
    If rstRecords!ParallelURL <> "" Then Me.txtURL = rstRecords!ParallelURL
    
    'If rstRecords!Notes <> "" Then Me.txtNotes = rstRecords!Notes Else Me.txtNotes = ""
    If rstRecords!LibraryCOllection = True Then Me.chkLibraryCollection = 1 Else Me.chkLibraryCollection = 0
    If rstRecords!Republished = True Then Me.chkRepublished = 1 Else Me.chkRepublished = 0
    
    sSourceType = Me.cmbSourceType.Text
    'If Me.cmbRecordNumber.Text = "New Record" Then Me.cmbRecordNumber.Text = rtsrecords!recordid
    iRecNum = Me.cmbRecordNumber.Text
    'iRecNum = rstRecords!recordid
    Select Case sSourceType
        Case "Chapter in Treatise"
            Set rstLargerWorksChapters = New ADODB.Recordset
            rstLargerWorksChapters.CursorLocation = adUseClient
            rstLargerWorksChapters.Open "Select * FROM qryLargerworksChapters WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstLargerWorksChapters.EOF Then
                If rstLargerWorksChapters!LargerWorkTitle <> "" Then Me.cmbLargerWorkTitle = rstLargerWorksChapters!LargerWorkTitle
                Me.txtLargerWorkID.Text = rstLargerWorksChapters!LargerWorkID
                Me.txtChapterID.Text = rstLargerWorksChapters!chapterID
                If rstLargerWorksChapters!CallNumber <> "" Then Me.txtCallNumber = rstLargerWorksChapters!CallNumber
                If rstLargerWorksChapters!EditionAndPrinting <> "" Then Me.txtEditionandPrinting = rstLargerWorksChapters!EditionAndPrinting
                If rstLargerWorksChapters!Publisher <> "" Then Me.txtPublisher = rstLargerWorksChapters!Publisher
                If rstLargerWorksChapters!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate = rstLargerWorksChapters!OriginalPublicationDate
                If rstLargerWorksChapters!SeriesVolume <> "" Then Me.txtSeriesVolume = rstLargerWorksChapters!SeriesVolume
                If rstLargerWorksChapters!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor = rstLargerWorksChapters!TitleOfSeriesIfNotIssuedByAuthor
            End If
            rstLargerWorksChapters.Close
        Case "Journal Article"
            Set rstArticlesJournals = New ADODB.Recordset
            rstArticlesJournals.CursorLocation = adUseClient
            rstArticlesJournals.Open "Select * FROM qryarticlesjournals WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstArticlesJournals.EOF Then
                Me.cmbJournalTitle.Text = rstArticlesJournals!JournalTitle
                Me.txtArticleID.Text = rstArticlesJournals!articleID
                'frmNewJournal.txtJournalID = rstArticlesJournals!JournalID
                'frmNewJournal.txtNewJournal = rstArticlesJournals!JournalTitle
                'frmNewJournal.txtNewJournalShortForm = rstArticlesJournals!JournalTitleShortFOrm
                'frmNewJournal.cmbPagination.Text = rstArticlesJournals!Pagination
                'If rstArticlesJournals!CallNumber <> Null Then frmNewJournal.txtCallNumber = rstArticlesJournals!CallNumber
                'If rstArticlesJournals!PlaceOfPublication <> Null Then frmNewJournal.txtPlaceOfPublication = rstArticlesJournals!PlaceOfPublication
                
                If rstArticlesJournals!Volume <> "" Then Me.txtVolume.Text = rstArticlesJournals!Volume
                If rstArticlesJournals!PublicationMonthOrSeason <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstArticlesJournals!PublicationMonthOrSeason
                If rstArticlesJournals!PublicationDay <> "" Then Me.txtPublicationDay = rstArticlesJournals!PublicationDay
                If rstArticlesJournals!ArticleDesignationForCitation <> "" Then Me.cmbArticleDesignation = rstArticlesJournals!ArticleDesignationForCitation
                Me.txtJournalID = rstArticlesJournals!JournalID
                Me.txtJournaTitleShortForm = rstArticlesJournals!JournalTitleShortFOrm
                'If rstArticlesJournals!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstArticlesJournals!JournalTitleShortForm
                If rstArticlesJournals!Pagination <> "" Then Me.cmbPagination = rstArticlesJournals!Pagination
                If rstArticlesJournals!CallNumber <> "" Then Me.txtCallNumber = rstArticlesJournals!CallNumber
                'If rstArticlesJournals!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstArticlesJournals!PLaceOfPublication
                'Call Position_Article_Form
            End If
            rstArticlesJournals.Close
        Case "Legislative Material"
            Set rstLegislative = New ADODB.Recordset
            rstLegislative.CursorLocation = adUseClient
            rstLegislative.Open "Select * FROM tblLegislativeMaterial WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstLegislative.EOF Then
                Me.txtLegislativeID.Text = rstLegislative!LegislativeID
                If rstLegislative!materialtype <> "" Then Me.cmbLegislativeType.Text = rstLegislative!materialtype
                If rstLegislative!NameOfHouse <> "" Then Me.txtLegislativeHouse.Text = rstLegislative!NameOfHouse
                If rstLegislative!NumberOfCongress <> "" Then Me.txtNumberOfCongress.Text = rstLegislative!NumberOfCongress
                If rstLegislative!SessionOfCongress <> "" Then Me.txtSessionOfCongress.Text = rstLegislative!SessionOfCongress
                If rstLegislative!StateLegislativeSession <> "" Then Me.txtStateLegislativeSession.Text = rstLegislative!StateLegislativeSession
                If rstLegislative!USCCANCitation <> "" Then Me.txtUSCCANCitation.Text = rstLegislative!USCCANCitation
                If rstLegislative!ReportOrDocumentNumber <> "" Then Me.txtReportOrDocumentNumber.Text = rstLegislative!ReportOrDocumentNumber
                If rstLegislative!SuDocNumber <> "" Then Me.txtSuDocNumber.Text = rstLegislative!SuDocNumber
            End If
            rstLegislative.Close

        Case "Treatise"
            Set rstTreatise = New ADODB.Recordset
            rstTreatise.CursorLocation = adUseClient
            rstTreatise.Open "Select * FROM tblTreatises WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstTreatise.EOF Then
                Me.txtTreatiseID.Text = rstTreatise!TreatiseID
                If rstTreatise!EditionAndPrinting <> "" Then Me.txtEditionandPrinting.Text = rstTreatise!EditionAndPrinting
                If rstTreatise!Publisher <> "" Then Me.txtPublisher.Text = rstTreatise!Publisher
                If rstTreatise!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate.Text = rstTreatise!OriginalPublicationDate
                If rstTreatise!SeriesVolume <> "" Then Me.txtSeriesVolume.Text = rstTreatise!SeriesVolume
                If rstTreatise!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = rstTreatise!TitleOfSeriesIfNotIssuedByAuthor
                If rstTreatise!CallNumber <> "" Then Me.txtCallNumber.Text = rstTreatise!CallNumber
            End If
            rstTreatise.Close

        Case "Unpublished Work"
            Set rstUnpublished = New ADODB.Recordset
            rstUnpublished.CursorLocation = adUseClient
            rstUnpublished.Open "Select * FROM tblUnpublishedWork WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstUnpublished.EOF Then
                Me.txtUnpublishedID.Text = rstUnpublished!UnpublishedWorkID
                If rstUnpublished!Type <> "" Then Me.cmbUnpublishedType.Text = rstUnpublished!Type
                If rstUnpublished.Fields("Thesis/Dissertation Type") <> "" Then Me.cmbThesisDissertationType.Text = rstUnpublished.Fields("Thesis/Dissertation Type")
                If rstUnpublished!PublicationMonth <> "" Then Me.cmbPublicationMonthOrSeason = rstUnpublished!PublicationMonth
                If rstUnpublished!PublicationDay <> "" Then Me.txtPublicationDay.Text = rstUnpublished!PublicationDay
                If rstUnpublished!Location <> "" Then Me.txtLocation.Text = rstUnpublished!Location
                
            End If
            rstUnpublished.Close
            
        Case "Nonprint Material"
            Set rstOther = New ADODB.Recordset
            rstOther.CursorLocation = adUseClient
            rstOther.Open "Select * FROM tblMisc WHERE RecordID = " & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
            If Not rstOther.EOF Then
                Me.txtMiscID.Text = rstOther!MiscID
                If rstOther!RecordType <> "" Then Me.cmbMiscType.Text = rstOther!RecordType
                If rstOther!Location <> "" Then Me.txtLocation.Text = rstOther!Location
                If rstOther!WorkingPaper <> "" Then Me.txtWorkingPaper.Text = rstOther!WorkingPaper
                If rstOther!Month <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstOther!Month
                If rstOther!Day <> "" Then Me.txtPublicationDay.Text = rstOther!Day
                'If rstOther!Location <> "" Then Me.txtLocation.Text = rstOther!Location
            End If
            rstOther.Close
        End Select
    'rstRecordsAET.MoveFirst
    'rstAuthors.MoveFirst
    'rstAETLMFRecords.Open "SELECT * FROM qryAETLMFRecords WHERE RecordID=" & iRecNum, cnDatabase, adOpenStatic, adLockPessimistic
    rstAETLMFRecords.CursorLocation = adUseClient
    rstAETLMFRecords.Open "SELECT * FROM qryAETRecords WHERE RecordID=" & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    
    If rstAETLMFRecords.EOF Then
        lblA.Text = "No Author"
        lblE.Text = "No Editor"
        lblT.Text = "No Translator"
        
    End If
    
    Do While Not rstAETLMFRecords.EOF
        sCurrentAuthor = ""
        If rstAETLMFRecords!InstitutionalEntity <> "" Then sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords!InstitutionalEntity
        If rstAETLMFRecords!LastName <> "" Then
            If sCurrentAuthor <> "" Then sCurrentAuthor = sCurrentAuthor & ", "
            sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords!LastName
        End If
        If rstAETLMFRecords!FirstName <> "" Then
            If sCurrentAuthor <> "" Then sCurrentAuthor = sCurrentAuthor & ", "
            sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords!FirstName
        End If
        If rstAETLMFRecords!MiddleName <> "" Then sCurrentAuthor = sCurrentAuthor & " " & rstAETLMFRecords!MiddleName
        If rstAETLMFRecords!Suffix <> "" Then sCurrentAuthor = sCurrentAuthor & " " & rstAETLMFRecords!Suffix
        sCurrentAuthor = sCurrentAuthor & " (ID: " & rstAETLMFRecords!AETID & ")"
        
        'sCurrentAuthor = rstAETLMFRecords!FullName & " (ID: " & rstAETLMFRecords!AETID & ")"
        iAETID = rstAETLMFRecords!AETID
        sAETType = rstAETLMFRecords!AETType
        Select Case sAETType
            Case "Author"
                cAuthors.Add iAETID
                lstCurrentAuthors.AddItem sCurrentAuthor
        
                For iListCount = 0 To (lstAuthors.ListCount - 1)
                    If lstAuthors.List(iListCount) = sCurrentAuthor Then
                        lstAuthors.RemoveItem (iListCount)
                    End If
                Next
            Case "Editor"
                cEditors.Add iAETID
                lstCurrentEditors.AddItem sCurrentAuthor
        
                For iListCount = 0 To (lstEditors.ListCount - 1)
                    If lstEditors.List(iListCount) = sCurrentAuthor Then
                        lstEditors.RemoveItem (iListCount)
                    End If
                Next
            Case "Translator"
                cTranslators.Add iAETID
                lstCurrentTranslators.AddItem sCurrentAuthor
        
                For iListCount = 0 To (lstTranslators.ListCount - 1)
                    If lstTranslators.List(iListCount) = sCurrentAuthor Then
                        lstTranslators.RemoveItem (iListCount)
                    End If
                Next
        End Select
        If cAuthors.Count > 0 Then
            If cAuthors.Count = 1 Then lblA.Text = "Author"
            If cAuthors.Count > 1 Then lblA.Text = "Authors"
        Else
            lblA.Text = "No Author"
        End If
        If cEditors.Count > 0 Then
            If cEditors.Count = 1 Then lblE.Text = "Editor"
            If cEditors.Count > 1 Then lblE.Text = "Editors"
        Else
            lblE.Text = "No Editor"
        End If
        If cTranslators.Count > 0 Then
            If cTranslators.Count = 1 Then lblT.Text = "Translator"
            If cTranslators.Count > 1 Then lblT.Text = "Translators"
        Else
            lblT.Text = "No Translator"
        End If
        
        rstAETLMFRecords.MoveNext
    Loop
    rstQryKeywords.CursorLocation = adUseClient
    rstQryKeywords.Open "SELECT * FROM qryKeywords WHERE RecordID=" & iRecNum, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rstQryKeywords.EOF
        sCurrentKeyword = rstQryKeywords!KeywordOrCodeSection & " (ID: " & rstQryKeywords!KeywordID & ")"
        iKeywordID = rstQryKeywords!KeywordID
        cKeywords.Add iKeywordID
        lstCurrentKeywords.AddItem sCurrentKeyword

        For iListCount = 0 To (lstKeywords.ListCount - 1)
            If lstKeywords.List(iListCount) = sCurrentKeyword Then
                lstKeywords.RemoveItem (iListCount)
            End If
        Next
    

        rstQryKeywords.MoveNext
    Loop


    rstQryKeywords.Close
    
    rstAETLMFRecords.Close

    'Me.lstNewKeywords.Clear
    If Me.cmbRecordNumber.Text <> "" Then
        Call suggest_keywords
    End If

    Set rstQryKeywords = Nothing
    Set rstArticlesJournals = Nothing
    Set rstAETLMFRecords = Nothing
    Set rstLegislative = Nothing
    Set rstOther = Nothing
    Set rstTreatise = Nothing
    Set rstUnpublished = Nothing
    Me.txtStatus.Text = "Unchanged"
End Sub

Private Sub Change_Record_Lists()
    Dim icounter As Integer
    For icounter = 1 To cAuthors.Count
        'Manage_Lists lstAuthors, lstCurrentAuthors, cAuthors, (iCounter - 1)
    
        Manage_Lists lstAuthors, lstCurrentAuthors, cAuthors, 0
    Next
    For icounter = 1 To cEditors.Count
        Manage_Lists lstEditors, lstCurrentEditors, cEditors, 0
    Next
    For icounter = 1 To cTranslators.Count
        Manage_Lists lstTranslators, lstCurrentTranslators, cTranslators, 0
    Next
    For icounter = 1 To cKeywords.Count
        Manage_Lists lstKeywords, lstCurrentKeywords, cKeywords, 0
    Next

End Sub


Private Sub mneNewAuthor_Click(Index As Integer)
    frmNewAuthor.Show
End Sub

Private Sub mnuNewJournal_Click(Index As Integer)
    frmNewJournal.Show
End Sub

Private Sub tglNewRecords_Click()
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglNewRecords.Value = True
    If tglNewRecords.Value = True Then
        tglUpdateRecords.Value = False
        tglImportRecords.Value = False
        iSaveListIndex = Me.cmbRecordNumber.ListIndex
        Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
        Me.cmdSave.Caption = "Save"
        Me.cmbRecordNumber.Enabled = False
    End If
    Call Set_Entry_Form
    lblA.Text = "No Author"
    lblE.Text = "No Editor"
    lblT.Text = "No Translator"

End Sub



Private Sub tglUpdateRecords_Click()
    'If Me.tglImportRecords.Value = True Then Call Refresh_Record_List
    If (Me.tglImportRecords.Value = False) And (Me.tglNewRecords.Value = False) And _
        (Me.tglUpdateRecords.Value = True) Then GoTo Already_Update
    Me.cmbSourceType.CausesValidation = True
    
    If tglUpdateRecords.Value = True Then
        tglNewRecords.Value = False
        tglImportRecords.Value = False
        Me.txtStatus.Enabled = True
    End If
    If Not (rstRecords.State = 0) And (tglUpdateRecords.Value = True) Then
        'rstRecords.Requery
        Call Refresh_Record_List
        rstRecords.MoveFirst
        Me.cmbRecordNumber.Enabled = True
        Me.cmbRecordNumber = rstRecords!RecordID
    End If
    Call Change_Record_Lists
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglUpdateRecords.Value = True
    'If tglNewRecords.Value = false Then Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
    If tglUpdateRecords.Value = True Then
        'Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
        Me.cmdSave.Caption = "Update"
        Me.chkKeepSelected.Enabled = False
        Me.chkKeepSelected.Value = False
        Me.chkSource.Enabled = False
        Me.chkSource.Value = False
        Me.chkYear.Enabled = False
        Me.chkYear.Value = False
        Me.cmdDelete.Enabled = True
        Me.cmbRecordNumber.Enabled = True
        Me.cmdNextRecord.Enabled = True
        Me.cmdPreviousRecord.Enabled = True
        Me.lblStatus.Visible = True
        Me.txtStatus.Visible = True
        'If Me.cmbRecordNumber.ListCount > 0 Then Me.cmbRecordNumber.ListIndex = iSaveListIndex
    End If

Already_Update:
End Sub


Private Sub tglImportRecords_Click()
    Me.cmdSave.Caption = "Save"
    If tglImportRecords.Value = True Then
        tglUpdateRecords.Value = False
        tglNewRecords.Value = False
        'Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
        Me.cmbRecordNumber.Enabled = True
        Me.cmdSave.Caption = "Update"
        Me.chkKeepSelected.Enabled = False
        Me.chkKeepSelected.Value = False
        Me.chkSource.Enabled = False
        Me.chkSource.Value = False
        Me.chkYear.Enabled = False
        Me.chkYear.Value = False
        Me.cmdDelete.Enabled = True
        Me.cmbRecordNumber.Enabled = True
        Me.cmdNextRecord.Enabled = True
        Me.cmdPreviousRecord.Enabled = True
        Me.lblStatus.Visible = True
        Me.txtStatus.Visible = True
        'If Me.cmbRecordNumber.ListCount > 0 Then Me.cmbRecordNumber.ListIndex = iSaveListIndex
        frmFilter.Show
        frmFilter.txtQuery.SetFocus
    End If
    
    'Call Change_Record_Lists
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglImportRecords.Value = True
            
End Sub

Private Sub Clear_Form()
'    Me.cmbSourceType.Text = ""
    Me.txtInputInitials = ""
    Me.txtDateAdded = ""
    Me.txtDateUpdated = ""
    Me.txtInputInitials = ""
    If Me.chkYear.Value = True Then Me.txtYear = ""
    Me.txtTitle = ""
    Me.txtArticleID = ""
    Me.txtChapterID = ""
    Me.txtUnpublishedID = ""
    Me.txtLegislativeID = ""
    Me.txtTreatiseID = ""
    Me.txtMiscID = ""
    Me.txtPublicationDay = ""
    'Me.txtNotes = ""
    Me.txtURL = ""
    
    Me.lstNewKeywords.Clear
    Me.chkRepublished = False
    'Me.lblA.Visible = False
    'Me.lblE.Visible = False
    'Me.lblT.Visible = False
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Set_Entry_Form()
    Dim dDate As Date
    Dim iSaveCmbListIndex As Integer
    Dim iTempIndex As Integer
    
    iSaveListIndex = Me.cmbSourceType.ListIndex
    If Me.chkSource.Value = 0 Then iSaveListIndex = 0
        
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
        iTempIndex = Me.cmbJournalTitle.ListIndex
    End If
        
    Call Erase_Form
    Call Clear_Form
    Me.cmbSourceType.CausesValidation = True
    'Me.cmbSourceType.ListIndex = 0 'default to Journal Entry
    Me.cmbSourceType.SetFocus
    dDate = Now
    
    If tglNewRecords.Value = True Then
        Me.txtInputInitials = "WLB"
        Me.txtDateAdded = dDate
        Me.txtStatus.Visible = False
        Me.lblStatus.Visible = False
        Me.chkKeepSelected.Enabled = True
        Me.chkSource.Enabled = True
        Me.chkYear.Enabled = True
        Me.cmbSourceType.ListIndex = -1 'this gets the next statement to effect a click procedure when it detects a change in value
        Me.cmbSourceType.ListIndex = iSaveListIndex
        'Me.cmbRecordNumber.Enabled = False
        Me.cmdNextRecord.Enabled = False
        Me.cmdPreviousRecord.Enabled = False
        Me.cmdDelete.Enabled = False
    
    End If
    
    Call Change_Record_Lists
    If ((Me.chkKeepSelected.Value = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
        Me.cmbJournalTitle.ListIndex = -1
        Me.cmbJournalTitle.ListIndex = iTempIndex
    End If
End Sub

Private Sub txtCallNumber_Change()
Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtEditionAndPrinting_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub



Private Sub txtLegislativeHouse_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtLocation_Change()
Me.txtStatus.Text = "Not Saved"
End Sub

Public Function Bill_Replace(sString As String, sStringToReplace As String, sReplacementString _
    As String) As String
        Dim lReplacePos As Long
        Dim lReplaceLength As Long
        Dim sLeftString As String
        Dim sRightString As String
        lReplacePos = 0
        lReplaceLength = 1
        Do While InStr((lReplacePos + lReplaceLength), sString, sStringToReplace) <> 0
            lReplacePos = InStr((lReplacePos + lReplaceLength), sString, sStringToReplace)
            sLeftString = Left(sString, lReplacePos - 1)
            sRightString = Right(sString, (Len(sString) - (lReplacePos - 1 + Len(sStringToReplace))))
            sString = sLeftString & sReplacementString & sRightString
            lReplaceLength = Len(sReplacementString)
        Loop
        Bill_Replace = sString
End Function

Private Sub txtTitle_Keyup(KeyCode As Integer, Shift As Integer)
    Dim sTitle As String
    Dim sLastTwo As String
    Dim sLastFour As String
    
    sTitle = Me.txtTitle.Text
    sLastTwo = Right(sTitle, 2)
    sLastFour = Right(sTitle, 4)
    
    
    If sLastTwo = "--" Then
        Me.txtTitle.Text = Bill_Replace(Me.txtTitle.Text, "--", Chr(151))
        Me.txtTitle.SelStart = Len(Me.txtTitle.Text)
    End If
    
    'If sLastFour = "sec." Then
    '    Me.txtTitle.Text = Bill_Replace(Me.txtTitle.Text, "sec.", Chr(167))
    '    Me.txtTitle.SelStart = Len(Me.txtTitle.Text)
    'End If
    
    
End Sub

Private Sub txtWorkingPaper_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

'Private Sub txtNotes_Change()
'    Me.txtStatus.Text = "Not Saved"
'End Sub
Private Sub txtURL_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtNumberOfCongress_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtOriginalPublicationDate_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtPage_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtPublicationDay_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtPublisher_Change()
Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtReportOrDocumentNumber_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtSeriesVolume_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtSessionOfCongress_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtStateLegislativeSession_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtSuDocNumber_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtTitle_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtTitleOfSeriesIfNotIssuedByAuthor_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtUSCCANCitation_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtVolume_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub txtYear_Change()
    Me.txtStatus.Text = "Not Saved"
End Sub

Private Sub Refresh_Record_List()
    rstRecords.Close
    With rstRecords
        .ActiveConnection = cnWriteDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblRecords")
    End With
    Call frmMain.populate_RecordID_List
    frmMain.cmbRecordNumber.ListIndex = 0
    
End Sub
