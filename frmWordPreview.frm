VERSION 5.00
Begin VB.Form frmWordPreview 
   Caption         =   "Preview Citation Form"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
   End
   Begin VB.OLE OLEWord 
      Class           =   "Word.Document.8"
      Height          =   2775
      Left            =   210
      OleObjectBlob   =   "frmWordPreview.frx":0000
      TabIndex        =   1
      Top             =   210
      Width           =   13215
   End
End
Attribute VB_Name = "frmWordPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Bopen As Boolean
    Dim iCount As Integer
    Dim i As Integer
    
    
    'iCount = frmWordPreview.OLEWord.object.Application.Documents.Count
    'If iCount > 2 Then Bopen = True
    
    'For i = 1 To iCount
    '    If frmWordPreview.OLEWord.object.Application.Documents(iCount) = "Document in Unnamed" Then frmWordPreview.OLEWord.object.Application.Documents(iCount).Close
    'Next
    
    frmWordPreview.OLEWord.object.Application.Documents(1).Close
    'If Not Bopen Then
    frmWordPreview.OLEWord.object.Application.Application.Quit
    
End Sub
