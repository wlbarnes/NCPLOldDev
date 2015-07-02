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
Private Sub Save_Keyword(rstKeyword As Recordset, cnConnection As Connection)
    Dim OldKeywordName As String
    Dim bigIndex As Recordset
    Dim oldKeywordString As String
    Dim newKeywordString As String
    
    OldKeywordName = frmKeywordChange.lstKeywords.List(frmKeywordChange.lstKeywords.ListIndex)
        
    Set bigIndex = New ADODB.Recordset
    With bigIndex
        .CursorLocation = adUseClient
        .ActiveConnection = cnConnection
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblBigTextIndex WHERE CONTAINS(AllKeywords, '" & OldKeywordName & "')")
    End With
    'bigIndex.Find ("CONTAINS(AllKeywords, '" & OldKeywordName & "')")
    
    Do While Not bigIndex.EOF
        oldKeywordString = bigIndex!AllKeywords
        newKeywordString = Replace(oldKeywordString, OldKeywordName, Me.txtKeyword.Text)
         cnConnection.BeginTrans
            On Error GoTo data_Error
            bigIndex!AllKeywords = newKeywordString
            bigIndex.Update
        cnConnection.CommitTrans
        bigIndex.MoveNext
    Loop
    
    
    If Me.cmdAdd.Caption = "Edit" Then
        frmKeywordChange.lstKeywords.List(frmKeywordChange.lstKeywords.ListIndex) = Me.txtKeyword.Text
        rstKeyword.MoveFirst
        Do Until rstKeyword!KeywordID = frmKeywordChange.txtKeywordID
            rstKeyword.MoveNext
        Loop
        
        cnConnection.BeginTrans
            On Error GoTo data_Error
            rstKeyword!KeywordOrCodeSection = Me.txtKeyword.Text
            rstKeyword.Update
        cnConnection.CommitTrans
        
    End If
    If Me.cmdAdd.Caption = "Add" Then
        cnConnection.BeginTrans
            On Error GoTo data_Error
            rstKeyword.AddNew
            rstKeyword!KeywordOrCodeSection = Me.txtKeyword.Text
            rstKeyword.Update
        cnConnection.CommitTrans
    End If
    Set bigIndex = Nothing
data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Editing Keyword"
             rstKeyword.CancelUpdate
             
             cnConnection.RollbackTrans
             Me.Hide
             Exit Sub
             
             'Resume Next
        End Select
    End If
End Sub

Private Sub cmdAdd_Click()
    Call Save_Keyword(frmKeywordChange.rstRemoteKeywords, frmKeywordChange.cnRemoteDatabase)
    Call Save_Keyword(frmKeywordChange.rstKeywords, frmKeywordChange.cnDatabase)
    
    Call frmKeywordChange.Fill_Keyword_List
    If Me.cmdAdd.Caption = "Add" Then frmKeywordChange.lstKeywords.Text = Me.txtKeyword.Text
    Me.Hide

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

