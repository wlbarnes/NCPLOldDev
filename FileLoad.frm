VERSION 5.00
Begin VB.Form FrmFileShow 
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   7335
   End
   Begin VB.FileListBox lstFile 
      Height          =   675
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   7335
   End
   Begin VB.DirListBox lstDir 
      Height          =   1440
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmFileShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    'MsgBox "You selected " & lstFile.List(lstFile.ListIndex)
    'Dim wApp As Word.Application
    Dim sFileName As String
    'Dim outList As ms
    'Dim wDoc As Word.Document
    
    
    'Set wApp = New Word.Application
    'Set Form1.Document1 = New Word.Document
    
    'Form1.Document1.Visible = True
    sFileName = Me.lstFile.Path
    If Right((lstDir.Path), 1) <> "\" Then sFileName = sFileName & "\"
    sFileName = sFileName & Me.lstFile.List(Me.lstFile.ListIndex)
    
    'frmWord.oleWord.Action = 0
    'frmWord.oleWord.CreateEmbed ""
    frmWord.oleWord.InsertObjDlg
    frmWord.oleWord.Class = "Word.Document.8"
    frmWord.oleWord.CreateEmbed sFileName, "Word.Document.8"
    
    frmWord.oleWord.SourceDoc = sFileName

    frmWord.oleWord.Action = 7
    
    'Set frmword.objWD = oleWord.object.Application
    Me.Hide
    'Form1.rtbFile.LoadFile sFileName
    'Set wDoc = wApp.Documents.Open(sFileName)
    'Form1.OLE1.CreateEmbed sFileName
    'Set wDoc = Form1.OLE1.SourceDoc
    
    'Form1.OLE1.Visible = True
    'Form1.OLE1.Enabled = True
End Sub

Private Sub lstDir_Change()
    lstFile.Path = lstDir.Path
    
    
End Sub

Private Sub lstDrive_Change()
    lstDir.Path = lstDrive.Drive
    lstFile.Path = lstDir.Path
    
    
End Sub


