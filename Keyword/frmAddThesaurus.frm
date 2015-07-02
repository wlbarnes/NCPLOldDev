VERSION 5.00
Begin VB.Form frmAddThesaurus 
   Caption         =   "Add Thesaurus Equivalent"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRecNum 
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtThesaurus 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmAddThesaurus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim sTempstr As String

On Error GoTo data_Error

If Me.Caption = "Add Thesaurus Equivalent" Then
    sTempstr = Me.txtThesaurus.Text
    
    frmKeywordChange.cnRemoteDatabase.BeginTrans
        frmThesaurusEntry.rstRemoteThesaurusFill.AddNew
        frmThesaurusEntry.rstRemoteThesaurusFill!ThesaurusEquivalent = sTempstr
        frmThesaurusEntry.rstRemoteThesaurusFill.Update
    frmKeywordChange.cnRemoteDatabase.CommitTrans
    
    
    frmKeywordChange.cnDatabase.BeginTrans
        frmThesaurusEntry.rstThesaurusFill.AddNew
        frmThesaurusEntry.rstThesaurusFill!ThesaurusEquivalent = sTempstr
        frmThesaurusEntry.rstThesaurusFill.Update
    frmKeywordChange.cnDatabase.CommitTrans
    
    
    Me.txtThesaurus.Text = ""
    Call frmThesaurusEntry.fill_list
    frmThesaurusEntry.lstThesaurus.Text = sTempstr
    Me.txtThesaurus.SetFocus
End If

If Me.Caption = "Edit Thesaurus Equivalent" Then
    sTempstr = Me.txtThesaurus.Text
    
    frmThesaurusEntry.rstThesaurusFill.MoveFirst
    Do While frmThesaurusEntry.rstThesaurusFill!ThesaurusEquivalentID <> Me.txtRecNum.Text
        frmThesaurusEntry.rstThesaurusFill.MoveNext
    Loop
    
    frmThesaurusEntry.rstRemoteThesaurusFill.MoveFirst
    Do While frmThesaurusEntry.rstRemoteThesaurusFill!ThesaurusEquivalentID <> Me.txtRecNum.Text
        frmThesaurusEntry.rstRemoteThesaurusFill.MoveNext
    Loop
        
    frmKeywordChange.cnRemoteDatabase.BeginTrans
        frmThesaurusEntry.rstRemoteThesaurusFill!ThesaurusEquivalent = sTempstr
        frmThesaurusEntry.rstRemoteThesaurusFill.Update
    frmKeywordChange.cnRemoteDatabase.CommitTrans
    
    frmKeywordChange.cnDatabase.BeginTrans
        frmThesaurusEntry.rstThesaurusFill!ThesaurusEquivalent = sTempstr
        frmThesaurusEntry.rstThesaurusFill.Update
    frmKeywordChange.cnDatabase.CommitTrans
    
    
    Me.txtThesaurus.Text = ""
    Call frmThesaurusEntry.fill_list
    Me.Hide
End If

'Next two Ifs do not do anything??

If Me.Caption = "Add Thesaurus Stack" Then
    sTempstr = Me.txtThesaurus.Text
    frmKeywordChange.rstStacks.Requery
    frmKeywordChange.cnDatabase.BeginTrans
        frmKeywordChange.rstStacks.AddNew
                frmKeywordChange.rstStacks!StackName = sTempstr
        frmKeywordChange.rstStacks.Update
    frmKeywordChange.cnDatabase.CommitTrans
    Me.txtThesaurus.Text = "**"
    Call frmStackEntry.Fill_Stack_List
    Call frmStackEntry.Fill_ST_List
    frmStackEntry.lstStacks.Text = sTempstr
    Me.txtThesaurus.SetFocus
End If

If Me.Caption = "Edit Thesaurus Stack" Then
    sTempstr = Me.txtThesaurus.Text
    frmKeywordChange.rstStacks.Requery
    frmKeywordChange.rstStacks.MoveFirst
    Do While frmKeywordChange.rstStacks!StackID <> Me.txtRecNum.Text
        frmKeywordChange.rstStacks.MoveNext
    Loop
    frmKeywordChange.cnDatabase.BeginTrans
        frmKeywordChange.rstStacks!StackName = sTempstr
        frmKeywordChange.rstStacks.Update
    frmKeywordChange.cnDatabase.CommitTrans
    Me.txtThesaurus.Text = ""
    Call frmStackEntry.Fill_Stack_List
    Call frmStackEntry.Fill_ST_List
    Me.Hide
End If


data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Editing Keyword"
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

Private Sub txtThesaurus_Validate(Cancel As Boolean)
    Dim sText As String
    If (Me.Caption = "Add Thesaurus Stack") Or (Me.Caption = "Edit Thesaurus Stack") Then
        sText = Me.txtThesaurus.Text
        If Left(sText, 2) <> "**" Then
            MsgBox "Stacks must begin with two asterisks (**)", vbCritical, "Wrong Format"
            Cancel = True
        End If
    End If
End Sub
