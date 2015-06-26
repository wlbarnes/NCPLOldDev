VERSION 5.00
Begin VB.Form frmWord 
   Caption         =   "Form1"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OLE oleWord 
      Class           =   "Word.Document.8"
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14175
   End
End
Attribute VB_Name = "frmWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objWD As Word.Application

Private Sub cmdLoad_Click()
    FrmFileShow.Show
End Sub

Private Sub Command1_Click()
    MsgBox objWD.Selection.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' MsgBox "keyascii"
End Sub

Private Sub Form_Load()
    frmWord.oleWord.InsertObjDlg
    'Me.oleWord.AutoActivate = True
    oleWord.DoVerb (vbOLEInPlaceActivate)
    'oleWord.Action = 7
    
    'Set objWD = oleWord.object.Application
    'oleWord.object.Application.Dialogs.Item(wdDialogFileOpen).Show
    'objWD.Selection.TypeText "Hello"
    'objWD.KeyString
    'objWD.Dialogs.Item(wdDialogFileOpen).Show
End Sub

Private Sub oleWord_KeyDown(KeyCode As Integer, Shift As Integer)
 'MsgBox (KeyCode)
End Sub

Private Sub oleWord_KeyPress(KeyAscii As Integer)
 'MsgBox (KeyAscii)
End Sub

Private Sub oleWord_KeyUp(KeyCode As Integer, Shift As Integer)
'MsgBox (KeyCode)
End Sub

Private Sub oleWord_Updated(Code As Integer)
'MsgBox ("updated")
End Sub
