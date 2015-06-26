VERSION 5.00
Begin VB.Form TreatiseForm 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   14955
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblEditionAndPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   17
      Text            =   "Edition And Printing"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox lblCallNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   10440
      TabIndex        =   16
      Text            =   "Call Number"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox lblTitleOfSeriesIfNotIssuedByAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   1800
      TabIndex        =   15
      Text            =   "Title Of Series If Not Issued By Author"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox lblPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   7440
      TabIndex        =   14
      Text            =   "Publisher"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lblSeriesVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Text            =   "Series Volume"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox lblOriginalPublicationDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   5160
      TabIndex        =   12
      Text            =   "Original Publication Date"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox lblYear 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Text            =   "Publication Year"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Text            =   "Title"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   11775
   End
   Begin VB.TextBox txtCallNumber 
      Height          =   285
      Left            =   10440
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtEditionAndPrinting 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtPublisher 
      Height          =   285
      Left            =   7440
      TabIndex        =   4
      Top             =   3120
      Width           =   4815
   End
   Begin VB.TextBox txtOriginalPublicationDate 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtSeriesVolume 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   8415
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "Keep year"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Frame frmCitationInfo 
      Caption         =   "Citation Information"
      Height          =   2775
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   12735
   End
End
Attribute VB_Name = "TreatiseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblOriginalPublicationDate_Change()

End Sub
