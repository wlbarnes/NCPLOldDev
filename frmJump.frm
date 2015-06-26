VERSION 5.00
Begin VB.Form frmJump 
   Caption         =   "Form1"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdJump 
      Caption         =   "Jump"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtRecNum 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblJump 
      Caption         =   "Jump to Record Number:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmJump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdJump_Click()
    On Error GoTo Err:
    frmMain.cmbRecordNumber.Text = Me.txtRecNum.Text
    Unload Me
Err:
    Select Case Err
        Case 0
        Case Else
            MsgBox "Record does not exist"
            Unload Me
    End Select
    Exit Sub
End Sub
