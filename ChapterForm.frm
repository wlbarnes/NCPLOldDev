VERSION 5.00
Begin VB.Form ChapterForm 
   Caption         =   "Form2"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   LinkTopic       =   "Form2"
   ScaleHeight     =   10335
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkYear 
      Caption         =   "Keep year"
      Height          =   255
      Left            =   10920
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Text            =   "Page Number"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox lblPage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Text            =   "Page Number"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   9000
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   3240
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   10920
      TabIndex        =   13
      Text            =   "Publication Year"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   10920
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox cmbLargerWorkTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   9615
   End
   Begin VB.CommandButton cmdEditLargerWork 
      Caption         =   "Edit This Larger Work"
      Height          =   315
      Left            =   10320
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CheckBox chkKeepSelected 
      Caption         =   "Check to keep same Larger Work selected for multiple entries"
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   4815
   End
   Begin VB.CommandButton cmdNewLargerWork 
      Caption         =   "New Larger Work"
      Height          =   315
      Left            =   10320
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox lblLargerWorkTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Text            =   "Larger Work Title"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lblTitleOfSeriesIfNotIssuedByAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Text            =   "Title Of Series If Not Issued By Author"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox lblSeriesVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Text            =   "Series Volume"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Text            =   "Chapter Title"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   11775
   End
   Begin VB.TextBox txtSeriesVolume 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   12735
   End
End
Attribute VB_Name = "ChapterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEditLargerWork_Click()

End Sub

Private Sub lblTitleOfSeriesIfNotIssuedByAuthor_Change()

End Sub

Private Sub txtYear_Change()

End Sub
