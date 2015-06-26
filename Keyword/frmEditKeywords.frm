VERSION 5.00
Begin VB.Form frmEditKeywords 
   Caption         =   "Edit Keywords"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKeyword 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditKeywords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    If Me.cmdAdd.Caption = "Edit" Then
        frmKeywordChange.lstKeywords.List(frmKeywordChange.lstKeywords.ListIndex) = Me.txtKeyword.Text
        frmKeywordChange.rstKeywords.MoveFirst
        Do Until frmKeywordChange.rstKeywords!KeywordID = frmKeywordChange.txtKeywordID
            frmKeywordChange.rstKeywords.MoveNext
        Loop
        
        frmKeywordChange.cnDatabase.BeginTrans
        On Error GoTo data_Error
        frmKeywordChange.rstKeywords!KeywordOrCodeSection = Me.txtKeyword.Text
        frmKeywordChange.rstKeywords.Update
        frmKeywordChange.cnDatabase.CommitTrans
        Call frmKeywordChange.Fill_Keyword_List
        Me.Hide
    End If
    If Me.cmdAdd.Caption = "Add" Then
        frmKeywordChange.cnDatabase.BeginTrans
        On Error GoTo data_Error
        frmKeywordChange.rstKeywords.AddNew
            frmKeywordChange.rstKeywords!KeywordOrCodeSection = Me.txtKeyword.Text
        frmKeywordChange.rstKeywords.Update
        frmKeywordChange.cnDatabase.CommitTrans
        Call frmKeywordChange.Fill_Keyword_List
        frmKeywordChange.lstKeywords.Text = Me.txtKeyword.Text
        Me.Hide
    End If
    
data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Editing Keyword"
             frmKeywordChange.rstKeywords.CancelUpdate
             
             frmKeywordChange.cnDatabase.RollbackTrans
             Me.Hide
             Exit Sub
             
             'Resume Next
        End Select
    End If

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub
