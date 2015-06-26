VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   Caption         =   "Input"
   ClientHeight    =   9225
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   118
      Text            =   "Status:Not Saved"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstNewKeywords 
      Height          =   840
      Left            =   9720
      Sorted          =   -1  'True
      TabIndex        =   117
      Top             =   6600
      Width           =   4215
   End
   Begin VB.CommandButton cmdGetNewKeywords 
      Caption         =   "Suggest New Keywords"
      Height          =   495
      Left            =   9720
      TabIndex        =   116
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewKeyword 
      Caption         =   "New Keyword"
      Height          =   255
      Left            =   2880
      TabIndex        =   115
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewAuthor 
      Caption         =   "New Author"
      Height          =   255
      Left            =   3240
      TabIndex        =   114
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   113
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtMiscID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   103
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUnpublishedID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   102
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtLegislativeID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   101
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTreatiseID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   100
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtChapterID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   99
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArticleID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   98
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbRecordNumber 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1560
      List            =   "frmMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   97
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextRecord 
      Caption         =   "-->"
      Height          =   495
      Left            =   6000
      TabIndex        =   92
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreviousRecord 
      Caption         =   "<--"
      Height          =   495
      Left            =   2640
      TabIndex        =   91
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4320
      TabIndex        =   90
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ListBox lstCurrentKeywords 
      Height          =   840
      Left            =   5280
      TabIndex        =   88
      Top             =   6600
      Width           =   4215
   End
   Begin VB.ListBox lstKeywords 
      Height          =   840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   87
      Top             =   6600
      Width           =   4215
   End
   Begin VB.ComboBox cmbAETChoice 
      Height          =   315
      Left            =   1440
      TabIndex        =   82
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtSuDocNumber 
      Height          =   285
      Left            =   4920
      TabIndex        =   75
      Top             =   7890
      Width           =   1695
   End
   Begin VB.TextBox txtLargerWorkID 
      Height          =   285
      Left            =   7680
      TabIndex        =   72
      Top             =   3930
      Width           =   1455
   End
   Begin VB.ComboBox cmbLargerWorkTitle 
      Height          =   315
      Left            =   -480
      TabIndex        =   70
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox txtReportOrDocumentNumber 
      Height          =   285
      Left            =   2280
      TabIndex        =   69
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox txtUSCCANCitation 
      Height          =   285
      Left            =   120
      TabIndex        =   67
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox txtStateLegislativeSession 
      Height          =   285
      Left            =   4920
      TabIndex        =   65
      Top             =   7050
      Width           =   1455
   End
   Begin VB.TextBox txtSessionOfCongress 
      Height          =   285
      Left            =   2520
      TabIndex        =   63
      Top             =   7050
      Width           =   1695
   End
   Begin VB.TextBox txtNumberOfCongress 
      Height          =   285
      Left            =   120
      TabIndex        =   61
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtLegislativeHouse 
      Height          =   285
      Left            =   3000
      TabIndex        =   59
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ComboBox cmbLegislativeType 
      Height          =   315
      Left            =   120
      TabIndex        =   57
      Top             =   6360
      Width           =   2055
   End
   Begin VB.ComboBox cmbMiscType 
      Height          =   315
      Left            =   7680
      TabIndex        =   55
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   4200
      TabIndex        =   53
      Top             =   4680
      Width           =   2415
   End
   Begin VB.ComboBox cmbUnpublishedType 
      Height          =   315
      Left            =   120
      TabIndex        =   50
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cmbThesisDissertationType 
      Height          =   315
      Left            =   2160
      TabIndex        =   49
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CheckBox chkAllChaptersBySameAuthor 
      Caption         =   "All Chapters By Same Author?"
      Height          =   255
      Left            =   6720
      TabIndex        =   48
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   5640
      Width           =   4095
   End
   Begin VB.TextBox txtSeriesVolume 
      Height          =   285
      Left            =   6840
      TabIndex        =   44
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtOriginalPublicationDate 
      Height          =   285
      Left            =   7680
      TabIndex        =   42
      Top             =   1290
      Width           =   1215
   End
   Begin VB.TextBox txtPublisher 
      Height          =   285
      Left            =   1680
      TabIndex        =   40
      Top             =   5400
      Width           =   4815
   End
   Begin VB.TextBox txtEditionAndPrinting 
      Height          =   285
      Left            =   120
      TabIndex        =   38
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtOrganizationIssuingNewsletter 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   36
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtNotes 
      Height          =   495
      Left            =   7560
      TabIndex        =   34
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtCallNumber 
      Height          =   285
      Left            =   4560
      TabIndex        =   32
      Top             =   5640
      Width           =   1815
   End
   Begin VB.ComboBox cmbPagination 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   30
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtPlaceOfPublication 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   28
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   6000
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cmbPublicationMonthOrSeason 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtVolume 
      Height          =   285
      Left            =   8400
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtPublicationDay 
      Height          =   285
      Left            =   6840
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtJournalID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox cmbJournalTitle 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txtJournalTitleShortForm 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   14
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ComboBox cmbArticleDesignation 
      Height          =   315
      Left            =   3600
      TabIndex        =   13
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   8535
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   9120
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtInputInitials 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   4
      Top             =   720
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
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   720
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
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox cmbSourceType 
      Height          =   315
      ItemData        =   "frmMain.frx":0004
      Left            =   120
      List            =   "frmMain.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.ListBox lstEditors 
      Enabled         =   0   'False
      Height          =   840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   83
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentEditors 
      Enabled         =   0   'False
      Height          =   840
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   84
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentTranslators 
      Enabled         =   0   'False
      Height          =   840
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   86
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentAuthors 
      Height          =   840
      Left            =   5280
      TabIndex        =   78
      Top             =   5160
      Width           =   4215
   End
   Begin VB.ListBox lstTranslators 
      Enabled         =   0   'False
      Height          =   840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   85
      Top             =   5160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstAuthors 
      Height          =   840
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   77
      Top             =   5160
      Width           =   4215
   End
   Begin VB.Label Label6 
      Caption         =   "Misc ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   112
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Treatise ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   111
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Unpublished ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   110
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Chapter ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   109
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Article ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   108
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Legis ID"
      Height          =   255
      Left            =   7560
      TabIndex        =   107
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblT 
      Caption         =   "T"
      Height          =   255
      Left            =   9600
      TabIndex        =   106
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblE 
      Caption         =   "E"
      Height          =   255
      Left            =   9600
      TabIndex        =   105
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblA 
      Caption         =   "A"
      Height          =   255
      Left            =   9600
      TabIndex        =   104
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSForms.ToggleButton tglImportRecords 
      Height          =   375
      Left            =   6840
      TabIndex        =   96
      Top             =   0
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "Import Records"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglUpdateRecords 
      Height          =   375
      Left            =   4920
      TabIndex        =   95
      Top             =   0
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "Update Records"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglNewRecords 
      Height          =   375
      Left            =   3120
      TabIndex        =   94
      Top             =   0
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "New Entries"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblRecordNumber 
      Caption         =   "Record Number"
      Height          =   255
      Left            =   120
      TabIndex        =   93
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblKeywords 
      Caption         =   "Select Keywords "
      Height          =   375
      Left            =   120
      TabIndex        =   89
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblAETChoice 
      Caption         =   "Select"
      Height          =   255
      Left            =   240
      TabIndex        =   81
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblDoubleClickToAdd 
      Caption         =   "Double-Click to Add or Remove"
      Height          =   615
      Left            =   4320
      TabIndex        =   80
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblArrow 
      Caption         =   "<<<--------->>>"
      Height          =   255
      Left            =   4320
      TabIndex        =   79
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblSuDocNumber 
      Caption         =   "SuDoc Number"
      Height          =   255
      Left            =   4920
      TabIndex        =   76
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label lblReportOrDocumentNumber 
      Caption         =   "Report/Document Number"
      Height          =   255
      Left            =   2280
      TabIndex        =   74
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label lblLargerWorkID 
      Caption         =   "Larger Work ID"
      Height          =   255
      Left            =   7680
      TabIndex        =   73
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblLargerWorkTitle 
      Caption         =   "Larger Work Title"
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblUSCCANCitation 
      Caption         =   "USCCAN Citation"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label lblStateLegislativeSession 
      Caption         =   "State Legislative Session"
      Height          =   255
      Left            =   4920
      TabIndex        =   66
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label lblSessionOfCongress 
      Caption         =   "Session of Congress"
      Height          =   255
      Left            =   2520
      TabIndex        =   64
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblNumberOfCongress 
      Caption         =   "Number of Congress"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblLegislativeHouse 
      Caption         =   "Name of House"
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblLegislativeType 
      Caption         =   "Legislative Type"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblMiscType 
      Caption         =   "Miscellaneous Type"
      Height          =   255
      Left            =   7680
      TabIndex        =   56
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLocation 
      Caption         =   "Location"
      Height          =   255
      Left            =   4200
      TabIndex        =   54
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblThesisDissertationType 
      Caption         =   "Thesis/Dissertation Type"
      Height          =   255
      Left            =   2160
      TabIndex        =   52
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblUnpublishedType 
      Caption         =   "Unpublished Work Type"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblTitleOfSeriesIfNotIssuedByAuthor 
      Caption         =   "Title of Series (If Not Issued By Author)"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblSeriesVolume 
      Caption         =   "Series Volume"
      Height          =   255
      Left            =   6840
      TabIndex        =   45
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblOriginalPublicationDate 
      Caption         =   "Original Publication Date"
      Height          =   255
      Left            =   7680
      TabIndex        =   43
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblPublisher 
      Caption         =   "Publisher"
      Height          =   255
      Left            =   7320
      TabIndex        =   41
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblEditionAndPrinting 
      Caption         =   "Edition/Printing"
      Height          =   255
      Left            =   8520
      TabIndex        =   39
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblOrganizationIssuingNewsletter 
      Caption         =   "Organization Issuing Newsletter"
      Height          =   255
      Left            =   5640
      TabIndex        =   37
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes"
      Height          =   255
      Left            =   8640
      TabIndex        =   35
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblCallNumber 
      Caption         =   "Call Number"
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblPagination 
      Caption         =   "Pagination"
      Height          =   255
      Left            =   1560
      TabIndex        =   31
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblPlaceOfPublication 
      Caption         =   "Place of Publication"
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblPage 
      Caption         =   "Page Number"
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume"
      Height          =   255
      Left            =   8400
      TabIndex        =   25
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblPublicationMonthOrSeason 
      Caption         =   "Publication Month (Or Season)"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblPublicationDay 
      Caption         =   "Publication Day"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblJournalID 
      Caption         =   "Journal ID"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblJournalTitleShortForm 
      Caption         =   "Journal Short Form"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblJournalTitle 
      Caption         =   "Journal Title"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblArticleDesignation 
      Caption         =   "Article Designation"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblInputInitials 
      Caption         =   "Input Initials"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblDateUpdated 
      Caption         =   "Date Updated"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDateAdded 
      Caption         =   "Date Added"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblSourceType 
      Caption         =   "Source Type"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
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
Const cnLeftAlign = 120
Const cnLine1LabelTop = 240
Const cnLine1TextTop = 480
Const cnLine2LabelTop = 1080
Const cnLine2TextTop = 1290
Const cnLine3LabelTop = 1800
Const cnLine3TextTop = 2040
Const cnLine4LabelTop = 2520
Const cnLine4TextTop = 2760
Const cnLine5LabelTop = 3240
Const cnLine5TextTop = 3480
Const cnLine6LabelTop = 3960
Const cnLine6TextTop = 4200
Public rstJournals As ADODB.Recordset
Public rstAuthors As ADODB.Recordset
Public rstEditors As ADODB.Recordset
Public rstTranslators As ADODB.Recordset
Dim rstArticles As ADODB.Recordset
Dim rstChapters As ADODB.Recordset
Dim rstMisc As ADODB.Recordset
Dim rstLegislativeMaterial As ADODB.Recordset
Dim rstRecords As ADODB.Recordset
Dim rstRecordsAET As ADODB.Recordset
Dim rstRecordsKeywords As ADODB.Recordset
Dim rstTreatises As ADODB.Recordset
Dim rstUnpublishedWork As ADODB.Recordset

Dim rstLargerWorks As ADODB.Recordset
Dim rstKeywords As ADODB.Recordset
Public cnDatabase As ADODB.Connection
Dim iSaveListIndex As Integer
Public cAuthors As Collection
Dim cEditors As Collection
Dim cTranslators As Collection
Dim cKeywords As Collection


Private Sub Populate_Comboboxes()
    Dim iCounter As Integer
    
    Do While Not rstJournals.EOF
        If rstJournals.Fields("JournalTitle").Value <> "" Then
            cmbJournalTitle.AddItem rstJournals.Fields("JournalTitle").Value
        End If
        rstJournals.MoveNext
    Loop
    
    Do While Not rstLargerWorks.EOF
        If rstLargerWorks.Fields("LargerWorkTitle").Value <> "" Then
            cmbLargerWorkTitle.AddItem rstLargerWorks.Fields("LargerWorkTitle").Value
        End If
        rstLargerWorks.MoveNext
    Loop
    
    Do While Not rstAuthors.EOF
        If rstAuthors.Fields("FullName").Value <> "" Then
            lstAuthors.AddItem rstAuthors.Fields("FullName").Value & " (ID: " & rstAuthors!AETID & ")"
        End If
        rstAuthors.MoveNext
    Loop
    
    Do While Not rstEditors.EOF
        If rstEditors.Fields("FullName").Value <> "" Then
            lstEditors.AddItem rstEditors.Fields("FullName").Value & " (ID: " & rstEditors!AETID & ")"
        End If
        rstEditors.MoveNext
    Loop
    
    Do While Not rstTranslators.EOF
        If rstTranslators.Fields("FullName").Value <> "" Then
            lstTranslators.AddItem rstTranslators.Fields("FullName").Value & " (ID: " & rstTranslators!AETID & ")"
        End If
        rstTranslators.MoveNext
    Loop
    
    Do While Not rstKeywords.EOF
        If rstKeywords.Fields("KeywordOrCodeSection").Value <> "" Then
            lstKeywords.AddItem rstKeywords.Fields("KeywordOrCodeSection").Value & " (ID: " & rstKeywords!KeywordID & ")"
        End If
        rstKeywords.MoveNext
    Loop
    
    Do While Not rstRecords.EOF
        cmbRecordNumber.AddItem rstRecords!recordid
        rstRecords.MoveNext
    Loop
    cmbRecordNumber.AddItem ("New Record")
    
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
        Case "Editors"
            
            Erase_Object lstAuthors
            Erase_Object lstTranslators
            Erase_Object lstCurrentAuthors
            Erase_Object lstCurrentTranslators
            lstEditors.Visible = True
            lstEditors.Enabled = True
            lstCurrentEditors.Visible = True
            lstCurrentEditors.Enabled = True
            
        Case "Translators"
        
            Erase_Object lstEditors
            Erase_Object lstAuthors
            Erase_Object lstCurrentEditors
            Erase_Object lstCurrentAuthors
            lstCurrentTranslators.Visible = True
            lstCurrentTranslators.Enabled = True
            lstTranslators.Visible = True
            lstTranslators.Enabled = True
            
    End Select
End Sub



Private Sub cmbJournalTitle_Click()
    Dim sJournalTitle As String
    sJournalTitle = cmbJournalTitle.Text
    'sJournalTitle = Replace(sJournalTitle, "'", "*")
    rstJournals.MoveFirst
    Do Until (rstJournals!journaltitle = sJournalTitle) Or rstJournals.EOF
        rstJournals.MoveNext
    Loop
    
    'rstJournals.Find "JournalTitle LIKE '" & sJournalTitle & "'"
    If Not rstJournals.EOF Then
        Me.txtJournalID = rstJournals!JournalID
        If rstJournals!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstJournals!JournalTitleShortForm
        If rstJournals!Pagination <> "" Then Me.cmbPagination = rstJournals!Pagination
        If rstJournals!CallNumber <> "" Then Me.txtCallNumber = rstJournals!CallNumber
        If rstJournals!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstJournals!PLaceOfPublication
    End If
End Sub

Private Sub cmbLargerWorkTitle_Click()
    Dim sLargerWorkTitle As String
    sLargerWorkTitle = cmbLargerWorkTitle.Text
    'sJournalTitle = Replace(sJournalTitle, "'", "*")
    rstLargerWorks.MoveFirst
    Do Until (rstLargerWorks!LargerworkTitle = sLargerWorkTitle) Or rstLargerWorks.EOF
        rstLargerWorks.MoveNext
    Loop
    
    'rstJournals.Find "JournalTitle LIKE '" & sJournalTitle & "'"
    If Not rstLargerWorks.EOF Then
        Me.txtLargerWorkID = rstLargerWorks!LargerWorkID
        If rstLargerWorks!CallNumber <> "" Then Me.txtCallNumber = rstLargerWorks!CallNumber
        If rstLargerWorks!EditionAndPrinting <> "" Then Me.txtEditionAndPrinting = rstLargerWorks!EditionAndPrinting
        If rstLargerWorks!Publisher <> "" Then Me.txtPublisher = rstLargerWorks!Publisher
        If rstLargerWorks!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate = rstLargerWorks!OriginalPublicationDate
        If rstLargerWorks!SeriesVolume <> "" Then Me.txtSeriesVolume = rstLargerWorks!SeriesVolume
        If rstLargerWorks!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor = rstLargerWorks!TitleOfSeriesIfNotIssuedByAuthor
        
        'TitleOfSeriesIfNotIssuedByAuthor
        
        
    End If
End Sub



Private Sub cmbRecordNumber_Click()
    Dim iRecNum As Integer
    If Me.cmbRecordNumber.ListIndex <> cmbRecordNumber.ListCount - 1 Then
        Me.tglUpdateRecords.Value = True
        If IsNumeric(Me.cmbRecordNumber.Text) Then iRecNum = Me.cmbRecordNumber.Text
        rstRecords.MoveFirst
        rstRecords.Find "RecordID=" & iRecNum
        Call Erase_Form
        Call Clear_Form
        Call Change_Record_Lists
        Call Fill_Form
    End If
    If Me.cmbRecordNumber.Text = "New Record" Then
        Me.tglNewRecords.Value = True
        
    End If
End Sub

Private Sub cmbSourceType_Click()
    Dim sSourceType As String
    
    Call Erase_Form
    
    sSourceType = cmbSourceType.Text
    Select Case sSourceType
        Case "Journal Article"
            Call Article_Form
        Case "Treatise"
            Call Treatise_Form
        Case "Chapter in Treatise"
            Call Chapter_Form
        Case "Unpublished Work"
            Call Unpublished_Form
        Case "Legislative Material"
            Call Legislative_Form
        Case "Nonprint Material"
            Call Misc_Form
    End Select
End Sub

Private Sub cmbSourceType_Validate(Cancel As Boolean)
    
    If cmbSourceType.Text = "" Then
        MsgBox "Please Enter a Source Type."
        Cancel = True
    End If
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
    iCurrentRecnum = Me.cmbRecordNumber.Text
    Set rstOldKeywords = New Recordset
    With rstOldKeywords
        .ActiveConnection = cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT AllKeywords from tblRecordsAllKeywords WHERE RecordID=" & iCurrentRecnum)
    End With
    If Not rstOldKeywords.EOF Then sOldKeywordString = rstOldKeywords!AllKeywords
    
    Set rstOldKeywords = Nothing
    
    Set cSuggestedKeywords = New Collection
    sTitleText = Me.txtTitle.Text
    'next line taken out later
    sTitleText = sTitleText & " " & sOldKeywordString
    Me.lstNewKeywords.Clear
    Set rstKeywordCheck = New Recordset
    With rstKeywordCheck
        .ActiveConnection = cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from tblKeywords")
    End With
    
    Set rstThesaurusCheck = New Recordset
    
    Do While Not rstKeywordCheck.EOF
        sKeywordText = rstKeywordCheck!keywordorcodesection
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
    With rstExistingKeywordBigCat
        .ActiveConnection = cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryRecordsKeywordsThesaurus WHERE LargerCategory=1 AND RecordID=" & iCurrentRecnum)
    End With
    
    Do While Not rstExistingKeywordBigCat.EOF
        Set rstBigCategory = New Recordset
        sTempText = rstExistingKeywordBigCat!keywordorcodesection
        With rstBigCategory
                .ActiveConnection = cnDatabase
                .CursorType = adOpenForwardOnly
                .LockType = adLockReadOnly
                .Open ("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'")
        End With
        If Not rstBigCategory.EOF Then
                sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
                sKeywordText = rstThesaurusCheck!keywordorcodesection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
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
    sJournalLocation = Me.txtPlaceOfPublication
    If Not (sJournalLocation = "") Then
    
    
        Set rstJournalKeyword = New Recordset
            With rstJournalKeyword
                    .ActiveConnection = cnDatabase
                    .CursorType = adOpenForwardOnly
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
        .ActiveConnection = cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryKeywordThesaurus where Not(IsNull(ThesaurusEquivalent))")
    End With
    
    Do While Not rstThesaurusCheck.EOF
        If rstThesaurusCheck!largercategory = 1 Then
            Set rstBigCategory = New Recordset
            sTempText = rstThesaurusCheck!keywordorcodesection
            With rstBigCategory
                .ActiveConnection = cnDatabase
                .CursorType = adOpenForwardOnly
                .LockType = adLockReadOnly
                .Open ("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'")
            End With
            If Not rstBigCategory.EOF Then
                sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
                sKeywordText = rstThesaurusCheck!keywordorcodesection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
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
        If rstThesaurusCheck!largercategory = 0 Then
            bDuplicate = False
            sThesaurusText = rstThesaurusCheck!ThesaurusEquivalent
            sKeywordText = rstThesaurusCheck!keywordorcodesection & " (ID: " & rstThesaurusCheck!KeywordID & ")"
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

Private Sub cmdNewAuthor_Click()
    frmNewAuthor.Show
End Sub

Private Sub cmdNewJournal_Click()
    frmNewJournal.Show
End Sub

Private Sub cmdNewKeyword_Click()
    frmKeywordThesaurus.Show
End Sub

Private Sub cmdNextRecord_Click()
    Dim iCounter As Integer
    iCounter = (Me.cmbRecordNumber.ListIndex) + 1
    If iCounter < Me.cmbRecordNumber.ListCount Then
        Me.cmbRecordNumber.ListIndex = iCounter
    End If
    'rstRecords.MoveNext
    'If rstRecords.EOF Then rstRecords.MoveLast
    'Call Erase_Form
    'Call Change_Record_Lists
    'Call Fill_Form
End Sub

Private Sub cmdPreviousRecord_Click()
    Dim iCounter As Integer
    iCounter = (Me.cmbRecordNumber.ListIndex) - 1
    If iCounter > -1 Then
        Me.cmbRecordNumber.ListIndex = iCounter
    End If

End Sub

Private Sub cmdSave_Click()
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
    Dim sNotes As String
    Dim sEditionAndPrinting As String
    Dim sPublisher As String
    Dim sOriginalPublicationDate As String
    Dim sSeriesVolume As String
    Dim sTitleOfSeriesIfNotIssuedByAuthor As String
    Dim bAllChaptersBySameAuthor As String
    Dim sUnpublishedWorkType As String
    Dim sThesisDissertationType As String
    Dim sLocation As String
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
    Dim iRecordID As Integer
    Dim iCounter As Integer
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
    
    'Dim sDay As String
    
    
   
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
    sNotes = Me.txtNotes.Text
    sEditionAndPrinting = Me.txtEditionAndPrinting.Text
    sPublisher = Me.txtPublisher.Text
    sOriginalPublicationDate = Me.txtOriginalPublicationDate.Text
    sSeriesVolume = Me.txtSeriesVolume.Text
    sTitleOfSeriesIfNotIssuedByAuthor = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
    bAllChaptersBySameAuthor = Me.chkAllChaptersBySameAuthor.Value
    sUnpublishedWorkType = Me.cmbUnpublishedType.Text
    sThesisDissertationType = Me.cmbThesisDissertationType.Text
    sLocation = Me.txtLocation.Text
    sMiscellaneousType = Me.cmbMiscType.Text
    sLegislativeType = Me.cmbLegislativeType.Text
    sNameOfHouse = Me.txtLegislativeHouse
    sNumberOfCongress = Me.txtNumberOfCongress.Text
    sSessionOfCongress = Me.txtSessionOfCongress.Text
    sStateLegislativeSession = Me.txtStateLegislativeSession.Text
    sUSCCANCitation = Me.txtUSCCANCitation.Text
    sReportDocumentNumber = Me.txtReportOrDocumentNumber.Text
    sSuDocNumber = Me.txtSuDocNumber.Text
    
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
    If Me.tglNewRecords = True Then
        Set rstCheck = New ADODB.Recordset
        sCheckString = "SELECT * FROM tblrecords WHERE (PublicationYear='" & sYear & "')"
            
        If sPageNumber <> "" Then sCheckString = sCheckString & " AND (PageNumber = '" & sPageNumber & "')"
        rstCheck.Open sCheckString, cnDatabase, adOpenKeyset, adLockOptimistic
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
    
    If Me.tglNewRecords.Value = True Then rstRecords.AddNew
        If sDateAdded <> "" Then rstRecords!DateRecordAdded = sDateAdded
        If Me.tglUpdateRecords.Value = True Then rstRecords!dateRecordUpdated = dDate
        If sInputInitials <> "" Then rstRecords!InputInitials = sInputInitials
        If sSourceType <> "" Then rstRecords!DocumentType = sSourceType
        If sTitle <> "" Then rstRecords!Title = sTitle
        If sPageNumber <> "" Then rstRecords!PageNumber = sPageNumber
        If sYear <> "" Then rstRecords!PublicationYear = sYear
        If sNotes <> "" Then rstRecords!Notes = sNotes
        'If Me.tglUpdateRecords = True Then rstRecords!RecordID = iRecordID
        If Me.tglNewRecords = True Then iRecordID = rstRecords.Fields("RecordID")
    rstRecords.Update
    
    Select Case sSourceType
        Case "Chapter in Treatise"
            If Me.tglNewRecords = True Then rstChapters.AddNew

            If Me.tglNewRecords.Value = False Then
                rstChapters.MoveFirst
                Do Until rstChapters!recordid = iRecordID
                    rstChapters.MoveNext
                Loop
            End If
            If Not rstChapters.EOF Then
                'If Me.tglUpdateRecords = True Then rstChapters!chapterID = iChapterID
        
                rstChapters!recordid = iRecordID
                rstChapters!LargerWorkID = iLargerWorkID
                If sSeriesVolume <> "" Then rstChapters!SeriesVolume = sSeriesVolume
                If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstChapters!TitleOfSeriesIfNotIssuedByAuthor = sTitleOfSeriesIfNotIssuedByAuthor
                
            End If
            rstChapters.Update
        Case "Journal Article"
            If Me.tglNewRecords.Value = True Then rstArticles.AddNew
            If Me.tglNewRecords.Value = False Then
                rstArticles.MoveFirst
                Do Until rstArticles!recordid = iRecordID
                    rstArticles.MoveNext
                Loop
            End If
            If Not rstArticles.EOF Then
                rstArticles!recordid = iRecordID
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
                
        Case "Legislative Material"
            If Me.tglNewRecords = True Then rstLegislativeMaterial.AddNew
            If Me.tglNewRecords.Value = False Then
                rstLegislativeMaterial.MoveFirst
                Do Until rstLegislativeMaterial!recordid = iRecordID
                    rstLegislativeMaterial.MoveNext
                Loop
            End If
            If Not rstLegislativeMaterial.EOF Then
                rstLegislativeMaterial!recordid = iRecordID
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
        Case "Treatise"
            If Me.tglNewRecords.Value = True Then rstTreatises.AddNew
            If Me.tglNewRecords.Value = False Then
                rstTreatises.MoveFirst
                Do Until rstTreatises!recordid = iRecordID
                    rstTreatises.MoveNext
                Loop
            End If
            If Not rstTreatises.EOF Then
            'If Me.tglNewRecords = True Then rstTreatises.AddNew
                rstTreatises!recordid = iRecordID
                If sEditionAndPrinting <> "" Then rstTreatises!EditionAndPrinting = sEditionAndPrinting
                If sPublisher <> "" Then rstTreatises!Publisher = sPublisher
                If sOriginalPublicationDate <> "" Then rstTreatises!OriginalPublicationDate = sOriginalPublicationDate
                If sSeriesVolume <> "" Then rstTreatises!SeriesVolume = sSeriesVolume
                If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstTreatises!TitleOfSeriesIfNotIssuedByAuthor = sTitleOfSeriesIfNotIssuedByAuthor
                If sCallNumber <> "" Then rstTreatises!CallNumber = sCallNumber
            End If
            rstTreatises.Update

        Case "Unpublished Work"
            If Me.tglNewRecords = True Then rstUnpublishedWork.AddNew
            If Me.tglNewRecords.Value = False Then
                rstUnpublishedWork.MoveFirst
                Do Until rstUnpublishedWork!recordid = iRecordID
                    rstUnpublishedWork.MoveNext
                Loop
            End If
            If Not rstUnpublishedWork.EOF Then
                rstUnpublishedWork!recordid = rstRecords!recordid
                If sUnpublishedWorkType <> "" Then rstUnpublishedWork!Type = sUnpublishedWorkType
                If sThesisDissertationType <> "" Then rstUnpublishedWork.Fields("Thesis/Dissertation Type") = sThesisDissertationType
                If rstUnpublishedWork!PublicationMonth <> "" Then rstUnpublishedWork!PublicationMonth = sPublicationMonth
                If rstUnpublishedWork!PublicationDay <> "" Then rstUnpublishedWork!PublicationDay = sPublicationDay
                If rstUnpublishedWork!Location <> "" Then rstUnpublishedWork!Location = sLocation
            End If
            rstUnpublishedWork.Update
            
        Case "Nonprint Material"
            If Me.tglNewRecords = True Then rstMisc.AddNew
            If Me.tglNewRecords.Value = False Then
                rstMisc.MoveFirst
                Do Until rstMisc!recordid = iRecordID
                    rstMisc.MoveNext
                Loop
            End If
            If Not rstMisc.EOF Then
                rstMisc!recordid = iRecordID
                rstMisc!RecordType = sMiscellaneousType
                rstMisc!Location = sLocation
                rstMisc!Month = sPublicationMonth
                rstMisc!Day = sPublicationDay
            End If
            rstMisc.Update
    End Select
    
    If Me.tglUpdateRecords.Value = True Then
        Set rstAETDelete = New Recordset
        Set rstKeywordDelete = New Recordset
        rstAETDelete.Open "Select * from tblRecordsAET WHERE RecordID=" & iRecordID, cnDatabase, adOpenKeyset, adLockOptimistic
        rstKeywordDelete.Open "Select * from tblRecordsKeywords WHERE RecordID=" & iRecordID, cnDatabase, adOpenKeyset, adLockOptimistic
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
        For iCounter = 1 To cAuthors.Count
            rstRecordsAET.AddNew
                rstRecordsAET!recordid = iRecordID
                rstRecordsAET!AETID = cAuthors.Item(iCounter)
            rstRecordsAET.Update
        Next
        
        
        For iCounter = 1 To cEditors.Count
            rstRecordsAET.AddNew
                rstRecordsAET!recordid = iRecordID
                rstRecordsAET!AETID = cEditors.Item(iCounter)
            rstRecordsAET.Update
        Next
        
        For iCounter = 1 To cTranslators.Count
            rstRecordsAET.AddNew
                rstRecordsAET!recordid = iRecordID
                rstRecordsAET!AETID = cTranslators.Item(iCounter)
            rstRecordsAET.Update
        Next
        
        For iCounter = 1 To cKeywords.Count
            rstRecordsKeywords.AddNew
                rstRecordsKeywords!recordid = iRecordID
                rstRecordsKeywords!KeywordID = cKeywords.Item(iCounter)
            rstRecordsKeywords.Update
        Next
    'End If
    'If sRecordType = "Journal Article" Then
    '    If iJournalNumber = 0 Then
    '        rstJournals.AddNew
    '            rstJournals!JournalTitle = sJournalTitle
    '            rstJournals!JournalTitleShortForm = sShortForm
    '            rstJournals!Pagination = sJournalType
    '            iJournalNumber = rstJournals.Fields("JournalID")
    '        rstJournals.Update
    '    End If
    '    rstArticles.AddNew
    '        rstArticles!RecordID = iRecordID
    '        If sVolume <> "" Then rstArticles!Volume = sVolume
    '        If sMonthSeason <> "" Then rstArticles!PublicationMonthOrSeason = sMonthSeason
    '        If sDay <> "" Then rstArticles!PublicationDay = sDay
    '        rstArticles!JournalID = iJournalNumber
    '    rstArticles.Update
    'End If
    'rstSave.AddNew
    '  rstSave!DocumentType = sSourceTypeID
    '  rstSave!JournalTitle = sJournalTitle
    '  rstSave!JournalTitleShortForm = sShortForm
    '  If sArticleDesignation <> "" Then rstSave!ArticleDesignationForCitation = sArticleDesignation
    '  rstSave!PublicationMonth = sMonthSeason
    '  If sDay <> "" Then rstSave!PublicationDay = sDay
    '  rstSave!publicationyear = sYear
    '  rstSave!PageNumber = sPage
    '  rstSave!Volume = sVolume
    '  rstSave!ArticleTitle = sTitle
    '  rstSave!DateRecordAdded = dDate
    '  rstSave!InputInitials = "WLB"
    '  iRecordID = rstSave.Fields("tblRecordinfoBlank.RecordID")
    'rstSave.Update
    'Set rstAuthor = New ADODB.Recordset
            
    'If Not bNoAuthor Then
    '    For iAuthorCounter = 1 To cAuthors.Count
    '        rstAuthor.Open "select * from tblAuthorsEditorsTranslators " & _
    '        " WHERE ((FirstName='" & cAuthorFirst(iAuthorCounter) & _
    '        "') AND (MiddleName='" & cAuthorMiddle(iAuthorCounter) & _
    '        "') AND (LastName='" & cAuthorLast(iAuthorCounter) & _
    '        "') AND (Suffix='" & cAuthorSuffix(iAuthorCounter) & _
    '        "') AND (AETType='Author'))", cnRecordInfo, adOpenStatic
    '
    '        If Not rstAuthor.EOF Then
    '            lAuthorID = rstAuthor.Fields("AETID")
    '
    '        End If
    '        If rstAuthor.EOF Then
    '            rstAET.AddNew
    '                rstAET!FirstName = cAuthorFirst.Item(iAuthorCounter)
    '                If cAuthorMiddle.Item(iAuthorCounter) <> "" Then rstAET!MiddleName = cAuthorMiddle.Item(iAuthorCounter)
    '                rstAET!LastName = cAuthorLast.Item(iAuthorCounter)
    '                rstAET!Suffix = cAuthorSuffix.Item(iAuthorCounter)
    '                rstAET!AETType = "Author"
    '                lAuthorID = rstAET.Fields("AETID")
    '            rstAET.Update
    '        End If
    '        rstAuthor.Close
    '
    '        rstRecordsAET.AddNew
    '           rstRecordsAET!RecordID = iRecordID
    '           rstRecordsAET!AETID = lAuthorID
    '        rstRecordsAET.Update
    '
    ''        rstSave.AddNew
    '            If cAuthorFirst.Item(iAuthorCounter) <> "" Then rstSave!AuthorFirstName = cAuthorFirst.Item(iAuthorCounter)
    '            If cAuthorMiddle.Item(iAuthorCounter) <> "" Then rstSave!AuthorMiddleName = cAuthorMiddle.Item(iAuthorCounter)
    '            If cAuthorLast.Item(iAuthorCounter) <> "" Then rstSave!AuthorLastName = cAuthorLast.Item(iAuthorCounter)
    '            If cAuthorSuffix.Item(iAuthorCounter) <> "" Then rstSave!AuthorSuffix = cAuthorSuffix.Item(iAuthorCounter)
    '            rstSave.Fields("tblAuthorBlank.RecordID") = iRecordID
    '        rstSave.Update
    '    Next
    'End If
    '
    
    'sSQL = "INSERT INTO tblRecordInfoBlank (ArticleTitle, JournalTitle, JournalTitleShortForm, PageNumber, Volume, PublicationMonth, PublicationDay, PublicationYear, DocumentType, InputInitials) " & _
        "VALUES('" & sTitle & "', '" & sJournalTitle & "', '" & sShortForm & "', '" & sPage & "', '" & _
        sVolume & "', '" & sMonthSeason & "', '" & iDay & "', '" & sYear & "'," & sSourceTypeID & ", WLB)"
    'Call cnRecordInfo.Execute(sSQL)
    'Set rstAuthor = Nothing
'SaveErr:
'        Select Case Err
'        Case Else
'            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
'             vbOKOnly + vbCritical, "Saving Error"
'        End Select
    If Me.tglNewRecords.Value = True Then
        Me.cmbRecordNumber.RemoveItem (Me.cmbRecordNumber.ListCount - 1)
        Me.cmbRecordNumber.AddItem iRecordID
        Me.cmbRecordNumber.AddItem "New Record"
        Call Set_Entry_Form
    End If
CancelErr:
End Sub


Private Sub Form_Load()
    Dim sConnectionstring As String
    Dim dDate As Date
    Set rstJournals = New ADODB.Recordset
    Set rstAuthors = New ADODB.Recordset
    Set rstEditors = New ADODB.Recordset
    Set rstTranslators = New ADODB.Recordset
    Set rstArticles = New ADODB.Recordset
    Set rstChapters = New ADODB.Recordset
    Set rstMisc = New ADODB.Recordset
    Set rstLegislativeMaterial = New ADODB.Recordset
    Set rstRecords = New ADODB.Recordset
    Set rstRecordsAET = New ADODB.Recordset
    Set rstRecordsKeywords = New ADODB.Recordset
    Set rstTreatises = New ADODB.Recordset
    Set rstUnpublishedWork = New ADODB.Recordset

    Set rstLargerWorks = New ADODB.Recordset
    Set rstKeywords = New ADODB.Recordset
    Set cnDatabase = New ADODB.Connection
    
    Set cAuthors = New Collection
    Set cEditors = New Collection
    Set cTranslators = New Collection
    Set cKeywords = New Collection
    
    sConnectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\database\ncpl.mdb"
    cnDatabase.Open (sConnectionstring)
    
    Me.cmbSourceType.CausesValidation = False
    Me.tglUpdateRecords = True
    Me.cmbAETChoice = "Author"
    With rstJournals
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblJournals")
    End With
    
    With rstAuthors
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from qryAETLMF WHERE AETType='Author'")
    End With
    
    With rstEditors
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from qryAETLMF WHERE AETType='Editor'")
    End With
    
    With rstTranslators
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from qryAETLMF WHERE AETType='Translator'")
    End With
    
    With rstLargerWorks
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblLargerWorks")
    End With
    
    With rstKeywords
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywords")
    End With
   
    With rstArticles
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblArticles")
    End With

    With rstChapters
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblChapters")
    End With
    
    With rstMisc
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblMisc")
    End With
    
    With rstLegislativeMaterial
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblLegislativeMaterial")
    End With
    
    With rstRecords
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblRecords")
    End With
    rstRecords.MoveFirst
    With rstRecordsAET
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblRecordsAET")
    End With
    
    With rstRecordsKeywords
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblRecordsKeywords")
    End With
    
    With rstTreatises
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblTreatises")
    End With
    
    With rstUnpublishedWork
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblUnpublishedWork")
    End With
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
    'Me.lblDateAdded.Visible = False
    'Me.lblDateAdded.Enabled = False
    'Me.txtDateAdded.Visible = False
    'Me.txtDateAdded.Enabled = False
    
    'Me.lblDateUpdated.Visible = False
    'Me.lblDateUpdated.Enabled = False
    'Me.txtDateUpdated.Visible = False
    'Me.txtDateUpdated.Enabled = False
    
    'Me.lblInputInitials.Visible = False
    'Me.lblInputInitials.Enabled = False
    'Me.txtInputInitials.Visible = False
    'Me.txtInputInitials.Enabled = False
    
    'Me.lblTitle.Enabled = False
    'Me.lblTitle.Visible = False
    'Me.txtTitle.Enabled = False
    'Me.txtTitle.Visible = False
    
    'Me.lblYear.Enabled = False
    'Me.lblYear.Visible = False
    'Me.txtYear.Enabled = False
    'Me.txtYear.Visible = False
    
    Erase_Object lblLargerWorkID
    Erase_Object txtLargerWorkID, True
    
    Erase_Object lblArticleDesignation
    Erase_Object cmbArticleDesignation, True
    
    Erase_Object lblJournalID
    Erase_Object txtJournalID, True
    
    Erase_Object lblPublicationDay
    Erase_Object txtPublicationDay, True
    
    Erase_Object lblPage
    Erase_Object txtPage, True
    
    Erase_Object lblVolume
    Erase_Object txtVolume, True
    
    Erase_Object lblJournalTitle
    Erase_Object cmbJournalTitle, True
    
    Erase_Object lblPublicationMonthOrSeason
    Erase_Object cmbPublicationMonthOrSeason, True
    
    Erase_Object lblJournalTitleShortForm
    Erase_Object txtJournalTitleShortForm, True
    
    Erase_Object lblOrganizationIssuingNewsletter
    Erase_Object txtOrganizationIssuingNewsletter, True
    
    Erase_Object lblCallNumber
    Erase_Object txtCallNumber, True

    Erase_Object lblPagination
    Erase_Object cmbPagination, True
    
    Erase_Object lblNotes
    Erase_Object txtNotes, True
    
    Erase_Object lblPage
    Erase_Object txtPage, True
    
    Erase_Object lblPlaceOfPublication
    Erase_Object txtPlaceOfPublication, True
    
    Erase_Object lblEditionAndPrinting
    Erase_Object txtEditionAndPrinting, True
    
    Erase_Object lblPublisher
    Erase_Object txtPublisher, True
    
    Erase_Object lblOriginalPublicationDate
    Erase_Object txtOriginalPublicationDate, True
    
    Erase_Object lblTitleOfSeriesIfNotIssuedByAuthor
    Erase_Object txtTitleOfSeriesIfNotIssuedByAuthor, True
    
    Erase_Object lblLocation
    Erase_Object txtLocation, True
    
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
    Erase_Object cmbLargerWorkTitle, True
    
    Erase_Object chkAllChaptersBySameAuthor
    
    Erase_Object cmdNewJournal
    
End Sub

Private Sub Article_Form()
    
    Position_Object lblArticleDesignation, cnLeftAlign, cnLine6LabelTop
    Position_Object cmbArticleDesignation, cnLeftAlign, cnLine6TextTop
      
    Position_Object lblJournalID, cnLeftAlign, cnLine4LabelTop
    Position_Object txtJournalID, cnLeftAlign, cnLine4TextTop

    Position_Object lblPublicationDay, 6520, cnLine6LabelTop
    Position_Object txtPublicationDay, 6520, cnLine6TextTop

    Position_Object lblPage, 7920, cnLine6LabelTop
    Position_Object txtPage, 7920, cnLine6TextTop
    
    Position_Object lblVolume, 2640, cnLine6LabelTop
    Position_Object txtVolume, 2640, cnLine6TextTop
    
    Position_Object lblJournalTitle, cnLeftAlign, cnLine3LabelTop
    Position_Object cmbJournalTitle, cnLeftAlign, cnLine3TextTop
    
    Position_Object lblPublicationMonthOrSeason, 4080, cnLine6LabelTop
    Position_Object cmbPublicationMonthOrSeason, 4080, cnLine6TextTop
    
    Position_Object lblJournalTitleShortForm, 5760, cnLine3LabelTop
    Position_Object txtJournalTitleShortForm, 5760, cnLine3TextTop
        
    Position_Object lblOrganizationIssuingNewsletter, cnLeftAlign, cnLine5LabelTop
    Position_Object txtOrganizationIssuingNewsletter, cnLeftAlign, cnLine5TextTop
            
    Position_Object lblCallNumber, 4800, cnLine5LabelTop
    Position_Object txtCallNumber, 4800, cnLine5TextTop
    txtCallNumber.BackColor = "-2147483629"
    txtCallNumber.Enabled = False
            
    Position_Object lblPagination, 1560, cnLine4LabelTop
    Position_Object cmbPagination, 1560, cnLine4TextTop
    
    Position_Object lblPlaceOfPublication, 4920, cnLine4LabelTop
    Position_Object txtPlaceOfPublication, 4920, cnLine4TextTop
    
    Me.cmdNewJournal.Visible = True
    Me.cmdNewJournal.Enabled = True

End Sub

Private Sub Treatise_Form()
          
    Position_Object lblCallNumber, 4560, cnLine4LabelTop
    Position_Object txtCallNumber, 4560, cnLine4TextTop
    txtCallNumber.BackColor = "-2147483643"
    txtCallNumber.Enabled = True
    
    Position_Object lblSeriesVolume, 6840, cnLine3LabelTop
    Position_Object txtSeriesVolume, 6840, cnLine3TextTop
    
    Position_Object lblEditionAndPrinting, cnLeftAlign, cnLine3LabelTop
    Position_Object txtEditionAndPrinting, cnLeftAlign, cnLine3TextTop
    
    Position_Object lblPublisher, 1680, cnLine3LabelTop
    Position_Object txtPublisher, 1680, cnLine3TextTop

    Position_Object lblOriginalPublicationDate, 10680, cnLine2LabelTop
    Position_Object txtOriginalPublicationDate, 10680, cnLine2TextTop
    
    Position_Object lblTitleOfSeriesIfNotIssuedByAuthor, cnLeftAlign, cnLine4LabelTop
    Position_Object txtTitleOfSeriesIfNotIssuedByAuthor, cnLeftAlign, cnLine4TextTop
    
End Sub

Private Sub Chapter_Form()
    
    Position_Object lblLargerWorkTitle, cnLeftAlign, cnLine3LabelTop
    Position_Object cmbLargerWorkTitle, cnLeftAlign, cnLine3TextTop
    
    Position_Object chkAllChaptersBySameAuthor, 6720, cnLine5TextTop
    
    Position_Object lblLargerWorkID, 5880, cnLine3LabelTop
    Position_Object txtLargerWorkID, 5880, cnLine3TextTop
    
    Position_Object lblPage, 7680, cnLine3LabelTop
    Position_Object txtPage, 7680, cnLine3TextTop
    
    Position_Object lblCallNumber, 4560, cnLine5LabelTop
    Position_Object txtCallNumber, 4560, cnLine5TextTop
    txtCallNumber.BackColor = "-2147483643"
    txtCallNumber.Enabled = True
    
    Position_Object lblSeriesVolume, 6840, cnLine4LabelTop
    Position_Object txtSeriesVolume, 6840, cnLine4TextTop
    
    Position_Object lblEditionAndPrinting, cnLeftAlign, cnLine4LabelTop
    Position_Object txtEditionAndPrinting, cnLeftAlign, cnLine4TextTop
    
    Position_Object lblPublisher, 1680, cnLine4LabelTop
    Position_Object txtPublisher, 1680, cnLine4TextTop

    Position_Object lblOriginalPublicationDate, 10680, cnLine2LabelTop
    Position_Object txtOriginalPublicationDate, 10680, cnLine2TextTop

    Position_Object lblTitleOfSeriesIfNotIssuedByAuthor, cnLeftAlign, cnLine5LabelTop
    Position_Object txtTitleOfSeriesIfNotIssuedByAuthor, cnLeftAlign, cnLine5TextTop
    
End Sub

Private Sub Misc_Form()

    Position_Object lblPublicationDay, 2640, cnLine4LabelTop
    Position_Object txtPublicationDay, 2640, cnLine4TextTop
    
    Position_Object lblPublicationMonthOrSeason, cnLeftAlign, cnLine4LabelTop
    Position_Object cmbPublicationMonthOrSeason, cnLeftAlign, cnLine4TextTop
    
    Position_Object lblLocation, 4200, cnLine4LabelTop
    Position_Object txtLocation, 4200, cnLine4TextTop
    
    Position_Object lblMiscType, cnLeftAlign, cnLine3LabelTop
    Position_Object cmbMiscType, cnLeftAlign, cnLine3TextTop
    
End Sub

Private Sub Legislative_Form()
  
    Position_Object lblLegislativeHouse, 3000, cnLine3LabelTop
    Position_Object txtLegislativeHouse, 3000, cnLine3TextTop
    
    Position_Object lblNumberOfCongress, cnLeftAlign, cnLine4LabelTop
    Position_Object txtNumberOfCongress, cnLeftAlign, cnLine4TextTop
    
    Position_Object lblSessionOfCongress, 2520, cnLine4LabelTop
    Position_Object txtSessionOfCongress, 2520, cnLine4TextTop
    
    Position_Object lblStateLegislativeSession, 4920, cnLine4LabelTop
    Position_Object txtStateLegislativeSession, 4920, cnLine4TextTop
    
    Position_Object lblUSCCANCitation, cnLeftAlign, cnLine5LabelTop
    Position_Object txtUSCCANCitation, cnLeftAlign, cnLine5TextTop
    
    Position_Object lblReportOrDocumentNumber, 2280, cnLine5LabelTop
    Position_Object txtReportOrDocumentNumber, 2280, cnLine5TextTop
    
    Position_Object lblSuDocNumber, 4920, cnLine5LabelTop
    Position_Object txtSuDocNumber, 4920, cnLine5TextTop
    
    Position_Object lblLegislativeType, cnLeftAlign, cnLine3LabelTop
    Position_Object cmbLegislativeType, cnLeftAlign, cnLine3TextTop
    
End Sub

Private Sub Unpublished_Form()
    Position_Object lblPublicationDay, 2640, cnLine4LabelTop
    Position_Object txtPublicationDay, 2640, cnLine4TextTop
    
    Position_Object lblPublicationMonthOrSeason, cnLeftAlign, cnLine4LabelTop
    Position_Object cmbPublicationMonthOrSeason, cnLeftAlign, cnLine4TextTop
    
    Position_Object lblLocation, 4200, cnLine4LabelTop
    Position_Object txtLocation, 4200, cnLine4TextTop
    
    Position_Object lblUnpublishedType, cnLeftAlign, cnLine3LabelTop
    Position_Object cmbUnpublishedType, cnLeftAlign, cnLine3TextTop
    
    Position_Object lblThesisDissertationType, 2160, cnLine3LabelTop
    Position_Object cmbThesisDissertationType, 2160, cnLine3TextTop

End Sub

Private Sub Form_Unload(Cancel As Integer)
    rstJournals.Close
    rstAuthors.Close
    rstEditors.Close
    rstTranslators.Clone
    rstLargerWorks.Close
    rstKeywords.Close
    rstArticles.Close
    rstChapters.Close
    rstMisc.Close
    rstLegislativeMaterial.Close
    rstRecords.Close
    rstRecordsAET.Close
    rstRecordsKeywords.Close
    rstTreatises.Close
    rstUnpublishedWork.Close
    
    cnDatabase.Close
    Set rstJournals = Nothing
    Set rstAuthors = Nothing
    Set rstEditors = Nothing
    Set rstTranslators = Nothing
    Set rstLargerWorks = Nothing
    Set rstKeywords = Nothing
    Set cnDatabase = Nothing
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



Private Sub lstAuthors_DblClick()
    Call Manage_Lists(lstCurrentAuthors, lstAuthors, cAuthors)
End Sub

Private Sub lstCurrentAuthors_DblClick()
    Call Manage_Lists(lstAuthors, lstCurrentAuthors, cAuthors)
End Sub
Private Sub lstEditors_DblClick()
    Call Manage_Lists(lstCurrentEditors, lstEditors, cEditors)
End Sub

Private Sub lstCurrentEditors_DblClick()
    Call Manage_Lists(lstEditors, lstCurrentEditors, cEditors)
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
End Sub

Private Sub lstTranslators_DblClick()
    Call Manage_Lists(lstCurrentTranslators, lstTranslators, cTranslators)
End Sub

Private Sub lstCurrentTranslators_DblClick()
    Call Manage_Lists(lstTranslators, lstCurrentTranslators, cTranslators)
End Sub
Private Sub lstKeywords_DblClick()
    Call Manage_Lists(lstCurrentKeywords, lstKeywords, cKeywords)
End Sub
Private Sub lstCurrentKeywords_DblClick()
    Call Manage_Lists(lstKeywords, lstCurrentKeywords, cKeywords)
End Sub

Public Sub Manage_Lists(oAdd As ListBox, oRemove As ListBox, cCollection As Collection, Optional iListIndex As Long = 999999)
    Dim sItem As String
    Dim iID As Integer
    
    Dim iParenpos As Integer
    'sItem = oRemove.Text
    If iListIndex = 999999 Then iListIndex = oRemove.ListIndex
    sItem = oRemove.List(iListIndex)
    oAdd.AddItem sItem
    oRemove.RemoveItem (iListIndex)
    If Mid(oAdd.Name, 4, 7) = "Current" Then
        iParenpos = InStr(1, sItem, " (ID: ")
        iID = Val(Mid(sItem, iParenpos + 6, (Len(sItem) - (iParenpos + 6))))
    End If
    If Mid(oAdd.Name, 4, 7) = "Current" Then cCollection.Add iID Else _
        cCollection.Remove (iListIndex + 1)
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
    Dim iCounter As Integer
    Dim iRecNum As Integer
    Dim sCurrentAuthor As String
    Dim sCurrentKeyword As String
    
    Dim iListCount As Integer
    Dim sAETType As String
    
    Set rstAETLMFRecords = New ADODB.Recordset
    Set rstQryKeywords = New ADODB.Recordset
    If rstRecords.EOF Then rstRecords.MoveFirst
    'Me.cmbRecordNumber.Text = rstRecords!recordid
    Me.txtTitle.Text = rstRecords!Title
    Me.cmbSourceType.Text = rstRecords!DocumentType
    If rstRecords!DateRecordAdded <> "" Then Me.txtDateAdded.Text = rstRecords!DateRecordAdded
    If rstRecords!dateRecordUpdated <> "" Then Me.txtDateUpdated.Text = rstRecords!dateRecordUpdated
    If rstRecords!InputInitials <> "" Then Me.txtInputInitials = rstRecords!InputInitials
    If rstRecords!PageNumber <> "" Then Me.txtPage = rstRecords!PageNumber
    If rstRecords!PublicationYear <> "" Then Me.txtYear = rstRecords!PublicationYear
    If rstRecords!Notes <> "" Then Me.txtNotes = rstRecords!Notes
    sSourceType = Me.cmbSourceType.Text
    iRecNum = Me.cmbRecordNumber.Text
    Select Case sSourceType
        Case "Chapter in Treatise"
            Set rstLargerWorksChapters = New ADODB.Recordset
            rstLargerWorksChapters.Open "Select * FROM qryLargerworksChapters WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
            If Not rstLargerWorksChapters.EOF Then
                If rstLargerWorksChapters!LargerworkTitle <> "" Then Me.cmbLargerWorkTitle = rstLargerWorksChapters!LargerworkTitle
                Me.txtLargerWorkID.Text = rstLargerWorksChapters!LargerWorkID
                Me.txtChapterID.Text = rstLargerWorksChapters!chapterID
                If rstLargerWorksChapters!CallNumber <> "" Then Me.txtCallNumber = rstLargerWorksChapters!CallNumber
                If rstLargerWorksChapters!EditionAndPrinting <> "" Then Me.txtEditionAndPrinting = rstLargerWorksChapters!EditionAndPrinting
                If rstLargerWorksChapters!Publisher <> "" Then Me.txtPublisher = rstLargerWorksChapters!Publisher
                If rstLargerWorksChapters!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate = rstLargerWorksChapters!OriginalPublicationDate
                If rstLargerWorksChapters!SeriesVolume <> "" Then Me.txtSeriesVolume = rstLargerWorksChapters!SeriesVolume
                If rstLargerWorksChapters!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor = rstLargerWorksChapters!TitleOfSeriesIfNotIssuedByAuthor
            End If
            rstLargerWorksChapters.Close
        Case "Journal Article"
            Set rstArticlesJournals = New ADODB.Recordset
            rstArticlesJournals.Open "Select * FROM qryarticlesjournals WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
            If Not rstArticlesJournals.EOF Then
                Me.cmbJournalTitle.Text = rstArticlesJournals!journaltitle
                Me.txtArticleID.Text = rstArticlesJournals!articleID
                If rstArticlesJournals!Volume <> "" Then Me.txtVolume.Text = rstArticlesJournals!Volume
                If rstArticlesJournals!PublicationMonthOrSeason <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstArticlesJournals!PublicationMonthOrSeason
                If rstArticlesJournals!PublicationDay <> "" Then Me.txtPublicationDay = rstArticlesJournals!PublicationDay
                If rstArticlesJournals!ArticleDesignationForCitation <> "" Then Me.cmbArticleDesignation = rstArticlesJournals!ArticleDesignationForCitation
                Me.txtJournalID = rstArticlesJournals!JournalID
                If rstArticlesJournals!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstArticlesJournals!JournalTitleShortForm
                If rstArticlesJournals!Pagination <> "" Then Me.cmbPagination = rstArticlesJournals!Pagination
                If rstArticlesJournals!CallNumber <> "" Then Me.txtCallNumber = rstArticlesJournals!CallNumber
                If rstArticlesJournals!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstArticlesJournals!PLaceOfPublication
            End If
            rstArticlesJournals.Close
        Case "Legislative Material"
            Set rstLegislative = New ADODB.Recordset
            rstLegislative.Open "Select * FROM tblLegislativeMaterial WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
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
            rstTreatise.Open "Select * FROM tblTreatises WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
            If Not rstTreatise.EOF Then
                Me.txtTreatiseID.Text = rstTreatise!TreatiseID
                If rstTreatise!EditionAndPrinting <> "" Then Me.txtEditionAndPrinting.Text = rstTreatise!EditionAndPrinting
                If rstTreatise!Publisher <> "" Then Me.txtPublisher.Text = rstTreatise!Publisher
                If rstTreatise!OriginalPublicationDate <> "" Then Me.txtOriginalPublicationDate.Text = rstTreatise!OriginalPublicationDate
                If rstTreatise!SeriesVolume <> "" Then Me.txtSeriesVolume.Text = rstTreatise!SeriesVolume
                If rstTreatise!TitleOfSeriesIfNotIssuedByAuthor <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = rstTreatise!TitleOfSeriesIfNotIssuedByAuthor
                If rstTreatise!CallNumber <> "" Then Me.txtCallNumber.Text = rstTreatise!CallNumber
            End If
            rstTreatise.Close

        Case "Unpublished Work"
            Set rstUnpublished = New ADODB.Recordset
            rstUnpublished.Open "Select * FROM tblUnpublishedWork WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
            If Not rstUnpublished.EOF Then
                Me.txtUnpublishedID.Text = rstUnpublished!UnpublishedWorkID
                If rstUnpublished!Type <> "" Then Me.cmbUnpublishedType.Text = rstUnpublished!Type
                If rstUnpublished.Fields("Thesis/Dissertation Type") <> "" Then Me.cmbThesisDissertationType.Text = rstUnpublished.Fields("Thesis/Dissertation Type")
                If rstUnpublished!PublicationMonth <> "" Then Me.cmbPublicationMonthOrSeason = rstUnpublished!PublicationMonth
                If rstUnpublished!PublicationDay <> "" Then Me.txtPublicationDay.Text = rstUnpublished!PublicationDay
                If rstUnpublished!Location <> "" Then Me.txtLocation = rstUnpublished!Location
            End If
            rstUnpublished.Close
            
        Case "Nonprint Material"
            Set rstOther = New ADODB.Recordset
            rstOther.Open "Select * FROM tblUnpublishedWork WHERE RecordID = " & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
            If Not rstOther.EOF Then
                Me.txtMiscID.Text = rstOther!MiscID
                If rstOther!RecordType <> "" Then Me.cmbMiscType.Text = rstOther!RecordType
                If rstOther!Location <> "" Then Me.txtLocation.Text = rstOther!Location
                If rstOther!Month <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstOther!Month
                If rstOther!Day <> "" Then Me.txtPublicationDay.Text = rstOther!Day
            End If
            rstOther.Close
        End Select
    'rstRecordsAET.MoveFirst
    'rstAuthors.MoveFirst
    rstAETLMFRecords.Open "SELECT * FROM qryAETLMFRecords WHERE RecordID=" & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
    Do While Not rstAETLMFRecords.EOF
        sCurrentAuthor = rstAETLMFRecords!FullName & " (ID: " & rstAETLMFRecords!AETID & ")"
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
        If cAuthors.Count > 0 Then lblA.Visible = True
        If cEditors.Count > 0 Then lblE.Visible = True
        If cTranslators.Count > 0 Then lblT.Visible = True
        
        
        rstAETLMFRecords.MoveNext
    Loop
    rstQryKeywords.Open "SELECT * FROM qryKeywords WHERE RecordID=" & iRecNum, cnDatabase, adOpenKeyset, adLockOptimistic
    
    Do While Not rstQryKeywords.EOF
        sCurrentKeyword = rstQryKeywords!keywordorcodesection & " (ID: " & rstQryKeywords!KeywordID & ")"
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

    'Do While Not rstRecordsAET.EOF
    '    If rstRecordsAET!RecordID = iRecNum Then
    '        rstAuthors.MoveFirst
    '        sCurrentAuthor = ""
    '        Do While Not rstAuthors.EOF
    '            If rstAuthors!AETID = rstRecordsAET!AETID Then
    '                sCurrentAuthor = rstAuthors!FullName & " (ID: " & rstAuthors!AETID & ")"
    '                lstCurrentAuthors.AddItem sCurrentAuthor
    '                lstAuthors.RemoveItem (sCurrentAuthor)
    '            End If
    '            rstAuthors.MoveNext
    '        Loop
    '    End If
    '    rstRecordsAET.MoveNext
    'Loop
    
    'For iCounter = 1 To cAuthors.Count
    '    rstRecordsAET.AddNew
    '        rstRecordsAET!RecordID = iRecordID
    '        rstRecordsAET!aetid = cAuthors.Item(iCounter)
    '    rstRecordsAET.Update
    'Next
    
    
    'For iCounter = 1 To cEditors.Count
    '    rstRecordsAET.AddNew
    '        rstRecordsAET!RecordID = iRecordID
    '        rstRecordsAET!aetid = cEditors.Item(iCounter)
    '    rstRecordsAET.Update
    'Next
    
    'For iCounter = 1 To cTranslators.Count
    '    rstRecordsAET.AddNew
    '        rstRecordsAET!RecordID = iRecordID
    '        rstRecordsAET!aetid = cTranslators.Item(iCounter)
    '    rstRecordsAET.Update
    'Next
    
    'For iCounter = 1 To cKeywords.Count
    '    rstRecordsKeywords.AddNew
    '        rstRecordsKeywords!RecordID = iRecordID
    '        rstRecordsKeywords!KeywordID = cKeywords.Item(iCounter)
    '    rstRecordsKeywords.Update
    'Next

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
    
End Sub

Private Sub Change_Record_Lists()
    Dim iCounter As Integer
    For iCounter = 1 To cAuthors.Count
        'Manage_Lists lstAuthors, lstCurrentAuthors, cAuthors, (iCounter - 1)
    
        Manage_Lists lstAuthors, lstCurrentAuthors, cAuthors, 0
    Next
    For iCounter = 1 To cEditors.Count
        Manage_Lists lstEditors, lstCurrentEditors, cEditors, 0
    Next
    For iCounter = 1 To cTranslators.Count
        Manage_Lists lstTranslators, lstCurrentTranslators, cTranslators, 0
    Next
    For iCounter = 1 To cKeywords.Count
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
    Call Set_Entry_Form
    
End Sub



Private Sub tglUpdateRecords_Click()
        
    Me.cmbSourceType.CausesValidation = True
    
    If tglUpdateRecords.Value = True Then
        tglNewRecords.Value = False
        tglImportRecords.Value = False
        Me.txtStatus.Enabled = True
    End If
    If Not (rstRecords.State = 0) Then rstRecords.MoveFirst
    Call Change_Record_Lists
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglUpdateRecords.Value = True
    'If tglNewRecords.Value = false Then Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
    If tglUpdateRecords.Value = True Then
        'Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
        Me.cmdSave.Caption = "Update"
        If Me.cmbRecordNumber.ListCount > 0 Then Me.cmbRecordNumber.ListIndex = iSaveListIndex
    End If

End Sub


Private Sub tglImportRecords_Click()
    Me.cmdSave.Caption = "Save"
    If tglImportRecords.Value = True Then
        tglUpdateRecords.Value = False
        tglNewRecords.Value = False
    End If
    Call Change_Record_Lists
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglImportRecords.Value = True
End Sub

Private Sub Clear_Form()
'    Me.cmbSourceType.Text = ""
    Me.txtInputInitials = ""
    Me.txtDateAdded = ""
    Me.txtDateUpdated = ""
    Me.txtInputInitials = ""
    Me.txtYear = ""
    Me.txtTitle = ""
    Me.txtArticleID = ""
    Me.txtChapterID = ""
    Me.txtUnpublishedID = ""
    Me.txtLegislativeID = ""
    Me.txtTreatiseID = ""
    Me.txtMiscID = ""
    Me.txtPublicationDay = ""
    Me.lblA.Visible = False
    Me.lblE.Visible = False
    Me.lblT.Visible = False
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Set_Entry_Form()
    Dim dDate As Date
    Call Erase_Form
    Call Clear_Form
    Me.cmbSourceType.CausesValidation = True
    Me.cmbSourceType.ListIndex = 0 'default to Journal Entry
    Me.cmbSourceType.SetFocus
    dDate = Now
    If tglNewRecords.Value = True Then
        tglUpdateRecords.Value = False
        tglImportRecords.Value = False
        Me.txtInputInitials = "WLB"
        Me.txtDateAdded = dDate
        Me.txtStatus.Visible = False
    End If
    
    Call Change_Record_Lists
    If (tglNewRecords.Value = False) And (tglUpdateRecords.Value = False) And _
        (tglImportRecords.Value = False) Then tglNewRecords.Value = True
    If tglNewRecords.Value = True Then
        iSaveListIndex = Me.cmbRecordNumber.ListIndex
        Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
        Me.cmdSave.Caption = "Save"
        
    End If
End Sub
