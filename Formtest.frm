VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      TabIndex        =   77
      Text            =   "<<<--------->>>"
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   675
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   76
      Text            =   "Form1.frx":0000
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      TabIndex        =   75
      Text            =   "<<<--------->>>"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   74
      Text            =   "Select Keywords"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   675
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   73
      Text            =   "Form1.frx":001E
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   72
      Text            =   "Select"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   66
      Text            =   "Volume"
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   65
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   2040
      TabIndex        =   64
      Text            =   "(click to keep journal selected for multiple entries)"
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   63
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
      TabIndex        =   62
      Text            =   "Article Designation"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   2280
      TabIndex        =   60
      Text            =   "Source Type"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   59
      Text            =   "Record Number"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Left            =   11280
      TabIndex        =   58
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   11280
      TabIndex        =   57
      Text            =   "Input Initials"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   8160
      TabIndex        =   56
      Text            =   "Date Updated"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   9720
      TabIndex        =   55
      Text            =   "Date Added"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox lstAuthors 
      Height          =   840
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   5400
      Width           =   4215
   End
   Begin VB.ListBox lstTranslators 
      Enabled         =   0   'False
      Height          =   840
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentAuthors 
      Height          =   840
      Left            =   5640
      TabIndex        =   36
      Top             =   5400
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentTranslators 
      Enabled         =   0   'False
      Height          =   840
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   35
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentEditors 
      Enabled         =   0   'False
      Height          =   840
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox lstEditors 
      Enabled         =   0   'False
      Height          =   840
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ComboBox cmbSourceType 
      Height          =   315
      ItemData        =   "Form1.frx":003C
      Left            =   2280
      List            =   "Form1.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   1200
      Width           =   3135
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
      Left            =   9720
      TabIndex        =   31
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
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   8160
      TabIndex        =   30
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   11520
      TabIndex        =   29
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   28
      Top             =   3120
      Width           =   12135
   End
   Begin VB.ComboBox cmbArticleDesignation 
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   3840
      Width           =   2175
   End
   Begin VB.ComboBox cmbJournalTitle 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   2400
      Width           =   6135
   End
   Begin VB.TextBox txtPublicationDay 
      Height          =   285
      Left            =   9720
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtVolume 
      Height          =   285
      Left            =   7800
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   5640
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Height          =   615
      Left            =   480
      TabIndex        =   22
      Top             =   9000
      Width           =   12135
   End
   Begin VB.TextBox txtEditionAndPrinting 
      Height          =   285
      Left            =   480
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtSeriesVolume 
      Height          =   285
      Left            =   7200
      TabIndex        =   20
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      Height          =   285
      Left            =   480
      TabIndex        =   19
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CheckBox chkAllChaptersBySameAuthor 
      Caption         =   "All Chapters By Same Author?"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   5880
      Width           =   2655
   End
   Begin VB.ComboBox cmbAETChoice 
      Height          =   315
      Left            =   1320
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ListBox lstKeywords 
      Height          =   840
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   7320
      Width           =   4215
   End
   Begin VB.ListBox lstCurrentKeywords 
      Height          =   840
      Left            =   5640
      TabIndex        =   15
      Top             =   7320
      Width           =   4215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   9960
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreviousRecord 
      Caption         =   "<--"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   9960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextRecord 
      Caption         =   "-->"
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   9960
      Width           =   1215
   End
   Begin VB.ComboBox cmbRecordNumber 
      Height          =   315
      ItemData        =   "Form1.frx":0040
      Left            =   480
      List            =   "Form1.frx":0042
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtSuDocNumber 
      Height          =   285
      Left            =   4920
      TabIndex        =   75
      Top             =   7890
      Width           =   1695
   End
   Begin VB.TextBox txtArticleID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   12960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtChapterID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Top             =   12720
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
   Begin VB.TextBox txtLegislativeID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9600
      TabIndex        =   8
      Top             =   12840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtUnpublishedID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   7
      Top             =   12960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtMiscID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   6
      Top             =   12120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdNewJournal 
      Caption         =   "New Journal"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewAuthor 
      Caption         =   "New Author"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewKeyword 
      Caption         =   "New Keyword"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetNewKeywords 
      Caption         =   "Suggest New Keywords"
      Height          =   255
      Left            =   10080
      TabIndex        =   2
      Top             =   7080
      Width           =   2535
   End
   Begin VB.ListBox lstNewKeywords 
      Height          =   840
      Left            =   10080
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Text            =   "Status:Not Saved"
      Top             =   11040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entry Information"
      Height          =   975
      Left            =   7920
      TabIndex        =   54
      Top             =   720
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record Information"
      Height          =   975
      Left            =   240
      TabIndex        =   61
      Top             =   720
      Width           =   6615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   71
      Top             =   1800
      Width           =   12735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Author Information"
      Height          =   1935
      Left            =   240
      TabIndex        =   78
      Top             =   4680
      Width           =   12735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Keyword Information"
      Height          =   1815
      Left            =   240
      TabIndex        =   80
      Top             =   6720
      Width           =   12735
   End
   Begin VB.Frame Frame6 
      Caption         =   "Notes"
      Height          =   1095
      Left            =   240
      TabIndex        =   81
      Top             =   8640
      Width           =   12735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   5175
      Left            =   -3600
      TabIndex        =   79
      Top             =   9840
      Width           =   30000
   End
   Begin VB.Label lblSeparate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   19995
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes"
      Height          =   255
      Left            =   9000
      TabIndex        =   52
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblSeriesVolume 
      Caption         =   "Series Volume"
      Height          =   255
      Left            =   9720
      TabIndex        =   51
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblTitleOfSeriesIfNotIssuedByAuthor 
      Caption         =   "Title of Series (If Not Issued By Author)"
      Height          =   255
      Left            =   480
      TabIndex        =   50
      Top             =   5640
      Width           =   2775
   End
   Begin MSForms.ToggleButton tglNewRecords 
      Height          =   375
      Left            =   2040
      TabIndex        =   49
      Top             =   120
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
   Begin MSForms.ToggleButton tglUpdateRecords 
      Height          =   375
      Left            =   6240
      TabIndex        =   48
      Top             =   120
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
   Begin MSForms.ToggleButton tglImportRecords 
      Height          =   375
      Left            =   10320
      TabIndex        =   47
      Top             =   120
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
   Begin VB.Label lblA 
      Caption         =   "A"
      Height          =   255
      Left            =   9960
      TabIndex        =   46
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblE 
      Caption         =   "E"
      Height          =   255
      Left            =   9960
      TabIndex        =   45
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblT 
      Caption         =   "T"
      Height          =   255
      Left            =   9960
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Legis ID"
      Height          =   255
      Left            =   8760
      TabIndex        =   43
      Top             =   12840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Article ID"
      Height          =   255
      Left            =   6600
      TabIndex        =   42
      Top             =   12960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Chapter ID"
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   12720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Unpublished ID"
      Height          =   255
      Left            =   4800
      TabIndex        =   40
      Top             =   12960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Misc ID"
      Height          =   255
      Left            =   6600
      TabIndex        =   39
      Top             =   12120
      Visible         =   0   'False
      Width           =   855
   End

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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblEditionAndPrinting_Click()

End Sub

Private Sub tglJournalTitle_Click()

End Sub

Private Sub Label5_Click()

End Sub
