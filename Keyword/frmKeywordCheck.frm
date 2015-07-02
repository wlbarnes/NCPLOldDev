VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmKeywordThesaurusChange 
   Caption         =   "Modify Keywords and Thesaurus Equivalents"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAddAsCategory 
      Caption         =   "Add Selection as Large Category?"
      Height          =   495
      Left            =   10440
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdGenThesaurus 
      Caption         =   "Thesaurus Table Entry"
      Height          =   735
      Left            =   7380
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdKeywordFolding 
      Caption         =   "Keyword Folding"
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtKeywordThesaurusID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox lstAllThesaurus 
      Height          =   2595
      Left            =   9720
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtAllThesaurusID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9840
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox lstThesaurusThesaurus 
      Height          =   2595
      Left            =   4320
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.ListBox lstThesaurusKeywords 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
   End
   Begin MSForms.CheckBox chkLargerCategory 
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   4680
      Width           =   3255
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5741;873"
      Value           =   "0"
      Caption         =   "Selected Thesaurus Entry to Act as Large Category?"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ToggleButton tglThesaurusEntry 
      Height          =   735
      Left            =   5040
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      VariousPropertyBits=   746588185
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2990;1296"
      Value           =   "1"
      Caption         =   "Keyword/Thesaurus Entry"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      Caption         =   "Thesaurus"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Thesaurus ID"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblDblClick 
      Caption         =   "<---------------------> Double-Click to Add/Remove Thesaurus Equivalents for Selected Keyword         <-------------------->"
      Height          =   1695
      Left            =   8520
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblThesaurusID 
      Caption         =   "Thesaurus ID"
      Height          =   255
      Left            =   11040
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblThesaurus 
      Caption         =   "Thesaurus"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblKeywordID 
      Caption         =   "Keyword ID"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblKeywords 
      Caption         =   "Keywords"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmKeywordThesaurusChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
    If Me.lstKeywords.SelCount = 0 Then
        MsgBox "Please select a Keyword first."
    Else
        frmEditThesaurus.Show
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim rstQueryThesaurus As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim sItem As String
        
    If Me.lstThesaurus.SelCount = 0 Then
        MsgBox "You have not selected a Thesaurus Equivalent to delete."
    Else
        'On Error GoTo data_Error
        Set rstQueryThesaurus = New ADODB.Recordset
        iThesaurusID = Me.txtThesaurusID.Text
        iKeywordID = Me.txtKeywordID.Text
        
        With rstQueryThesaurus
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblKeywordThesaurus WHERE KeywordID=" & iKeywordID & " AND ThesaurusID=" & iThesaurusID)
            '.Open ("SELECT * from qryThesaurus WHERE KeywordID=" & iKeywordID)
            
        End With
        'Do While rstQueryThesaurus!ThesaurusID <> iThesaurusID
        '    rstQueryThesaurus.MoveNext
        'Loop
        '& " AND ThesaurusID=" & iThesaurusID)
        
        'iItemnumber = lstThesaurus.ListIndex
        'iKeywordID = cThesaurusID.Item(iItemnumber + 1)
        'sItem = lstThesaurus.List(iItemnumber)
        
        'rstQueryThesaurus.MoveFirst
        'Do Until rstQueryThesaurus!ThesaurusID = iThesaurusID
            
        '    rstThesaurus.MoveNext
        'Loop
        frmKeywordChange.cnDatabase.BeginTrans
            
            rstQueryThesaurus.Delete
            rstQueryThesaurus.Update
            
        frmKeywordChange.cnDatabase.CommitTrans
        
        Set rstQueryThesaurus = Nothing
        Call Fill_Thesaurus_List
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

Private Sub chkLargerCategory_Click()
    
    Dim rstCategory As ADODB.Recordset
    
    If Me.lstThesaurusKeywords.SelCount = 0 Then
        MsgBox "You must select a thesaurus equivalent to act as larger category first."
    Else
        Set rstCategory = New ADODB.Recordset
        With rstCategory
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblKeywordThesaurus where KeywordID=" & Me.txtKeywordID & " AND ThesaurusID=" & Me.txtKeywordThesaurusID)
        End With
        frmKeywordChange.cnDatabase.BeginTrans
            On Error GoTo data_Error
            rstCategory!LargerCategory = Me.chkLargerCategory.Value
            rstCategory.Update
        frmKeywordChange.cnDatabase.CommitTrans
        
        'If rstCategory!LargerCategory = False Then Me.chkLargerCategory.Value = False
        'If rstCategory!LargerCategory = True Then Me.chkLargerCategory.Value = True
        
        Set rstCategory = Nothing
    End If
data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "Editing Keyword"
             rstCategory.CancelUpdate
             frmKeywordChange.cnDatabase.RollbackTrans
             
             Me.Hide
             Exit Sub
             
             'Resume Next
        End Select
    End If


End Sub

Private Sub cmdCombine_Click()
    Me.Hide
    frmCombine.Show
End Sub

Private Sub cmdGenThesaurus_Click()
    Unload Me
    Call frmThesaurusEntry.Show
    Call frmThesaurusEntry.fill_list
End Sub

Private Sub cmdKeywordFolding_Click()
    Me.Hide
    frmKeywordChange.Show
    Call frmKeywordChange.Fill_Keyword_List
End Sub

Private Sub cmdThesaurus_Click()
    frmThesaurusEntry.Show
End Sub

Private Sub cmdStack_Click()
    Me.Hide
    frmStackEntry.Show
    Call frmStackEntry.Fill_Stack_List
    Call frmStackEntry.Fill_ST_List
End Sub

Private Sub Form_Load()
    
    Call Fill_KT_List
    Call Fill_AT_List
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set rstKeywords = Nothing
    'Set rstThesaurus = Nothing
    'Set cKeywordID = Nothing
End Sub



Private Sub lstThesaurus_Click()
    Dim iItemnumber As Integer
    Dim iThesaurusID As Integer
    Dim sItem As String
    
    iItemnumber = lstThesaurus.ListIndex
    iThesaurusID = cThesaurusID.Item(iItemnumber + 1)
    sItem = lstThesaurus.List(iItemnumber)
    Me.txtThesaurusID.Text = iThesaurusID
End Sub

Public Sub Fill_TT_List()
    Dim rstQueryThesaurus As ADODB.Recordset
    Dim sCurrentItem As String
    Dim i As Integer
    Dim iThesaurusID As Integer
    Dim iKeywordID As Integer
    Set rstQueryThesaurus = New ADODB.Recordset
    iKeywordID = Me.txtKeywordID.Text
    Call Me.Fill_AT_List
    'rstquerythesaurus.Open ,,,
    With rstQueryThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryThesaurus WHERE KEYWORDID=" & iKeywordID)
    End With
    
    lstThesaurusThesaurus.Clear
    If Not rstQueryThesaurus.EOF Then
        rstQueryThesaurus.MoveFirst
        'Set cThesaurusID = New Collection
        Do While Not rstQueryThesaurus.EOF
            sCurrentItem = rstQueryThesaurus!ThesaurusEquivalent
            lstThesaurusThesaurus.AddItem sCurrentItem
            For i = 0 To (Me.lstAllThesaurus.ListCount - 1)
                If Me.lstAllThesaurus.List(i) = sCurrentItem Then
                    Me.lstAllThesaurus.RemoveItem (i)
                End If
            Next
            iThesaurusID = rstQueryThesaurus!ThesaurusID
            'cThesaurusID.Add iThesaurusID
            rstQueryThesaurus.MoveNext
        Loop
    End If
    Set rstQueryThesaurus = Nothing

End Sub
    
Public Sub Fill_KT_List()
    
    frmKeywordChange.rstKeywords.Requery
    Me.lstThesaurusKeywords.Clear
    If Not frmKeywordChange.rstKeywords.EOF Then
        frmKeywordChange.rstKeywords.MoveFirst
        Do While Not frmKeywordChange.rstKeywords.EOF
            Me.lstThesaurusKeywords.AddItem frmKeywordChange.rstKeywords!KeywordOrCodeSection
            'iIndex = rstKeywords!KeywordID
            'cKeywordID.Add iIndex ', iIndex
            frmKeywordChange.rstKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If
End Sub

Private Sub lstAllThesaurus_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    Dim i As Integer
    Dim sGetKeyText As String
    
    If Me.lstAllThesaurus.SelCount = 0 Then
        Me.txtAllThesaurusID.Text = ""
    Else
        sKeytext = Me.lstAllThesaurus.Text
        sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
        
        If Left(sKeytext, 2) = "**" Then
            Me.lblThesaurusID.Caption = "Stack ID"
            sGetKeyText = "SELECT StackID from tblThesaurusStack where StackName='" & sKeytext & "'"
        Else
            Me.lblThesaurusID.Caption = "Thesaurus ID"
            sGetKeyText = "SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sKeytext & "'"
        End If
        
        Set rstGetKeyNum = New ADODB.Recordset
        With rstGetKeyNum
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            '.Open ("SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sKeytext & "'")
            .Open (sGetKeyText)
        
        End With
                
        If Left(sKeytext, 2) = "**" Then
            iKeywordID = rstGetKeyNum!StackID
        Else
            iKeywordID = rstGetKeyNum!ThesaurusEquivalentID
        End If
        Set rstGetKeyNum = Nothing
    
        
        Me.txtAllThesaurusID = iKeywordID
        
        For i = 0 To (Me.lstThesaurusThesaurus.ListCount - 1)
            Me.lstThesaurusThesaurus.Selected(i) = False
        Next
    End If
End Sub

Private Sub lstAllThesaurus_DblClick()
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim iIndex As Integer
    Dim i As Integer
    Dim iStackID As Integer
    Dim bDuplicate As Boolean
    Dim rstStackThesaurus As ADODB.Recordset
    Dim rstCheckKeywordThesaurus As ADODB.Recordset
    Dim cThesaurusID As Collection
    Dim sThesaurusText As String
    iIndex = Me.lstAllThesaurus.ListIndex
    iKeywordID = Me.txtKeywordID
    
    If Me.lblThesaurusID.Caption = "Thesaurus ID" Then
        iThesaurusID = Me.txtAllThesaurusID
        frmKeywordChange.cnDatabase.BeginTrans
        On Error GoTo data_Error
        
            frmKeywordChange.rstKeywordsThesaurus.AddNew
                frmKeywordChange.rstKeywordsThesaurus!KeywordID = iKeywordID
                frmKeywordChange.rstKeywordsThesaurus!ThesaurusID = iThesaurusID
                If Me.chkAddAsCategory.Value = 1 Then
                    frmKeywordChange.rstKeywordsThesaurus!LargerCategory = True
                Else
                    frmKeywordChange.rstKeywordsThesaurus!LargerCategory = False

                End If
            frmKeywordChange.rstKeywordsThesaurus.Update
        frmKeywordChange.cnDatabase.CommitTrans
'save to remote database
        frmKeywordChange.cnRemoteDatabase.BeginTrans
        On Error GoTo data_Error
        
            frmKeywordChange.rstRemoteKeywordsThesaurus.AddNew
                frmKeywordChange.rstRemoteKeywordsThesaurus!KeywordID = iKeywordID
                frmKeywordChange.rstRemoteKeywordsThesaurus!ThesaurusID = iThesaurusID
                If Me.chkAddAsCategory.Value = 1 Then
                    frmKeywordChange.rstRemoteKeywordsThesaurus!LargerCategory = True
                Else
                    frmKeywordChange.rstRemoteKeywordsThesaurus!LargerCategory = False

                End If
            frmKeywordChange.rstRemoteKeywordsThesaurus.Update
        frmKeywordChange.cnRemoteDatabase.CommitTrans

        
        
    End If
        
    If Me.lblThesaurusID.Caption = "Stack ID" Then
        iStackID = Me.txtAllThesaurusID.Text
        
'next get thesaurus ids for keyword
        Set rstCheckKeywordThesaurus = New Recordset
        
        
        With rstCheckKeywordThesaurus
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open ("SELECT ThesaurusID from tblKeywordThesaurus where KeywordID=" & iKeywordID)
        End With
        
'compare thesaurus IDs and remove duplicates from collection
        If Not rstCheckKeywordThesaurus.EOF Then
            Do While Not rstCheckKeywordThesaurus.EOF
                iThesaurusID = rstCheckKeywordThesaurus!ThesaurusID
                For i = 1 To cThesaurusID.Count
                    If Not (i > cThesaurusID.Count) Then
                        If iThesaurusID = cThesaurusID.Item(i) Then
                            cThesaurusID.Remove (i)
                        End If
                    End If
                Next
                rstCheckKeywordThesaurus.MoveNext
            Loop
        End If
        
        Set rstCheckKeywordThesaurus = Nothing
        
'finally, add collection to database
        
        frmKeywordChange.rstKeywordsThesaurus.Requery
        frmKeywordChange.cnDatabase.BeginTrans
        On Error GoTo data_Error
        For i = 1 To cThesaurusID.Count
        
            frmKeywordChange.rstKeywordsThesaurus.AddNew
                frmKeywordChange.rstKeywordsThesaurus!KeywordID = iKeywordID
                frmKeywordChange.rstKeywordsThesaurus!ThesaurusID = cThesaurusID.Item(i)
                If Me.chkAddAsCategory.Value = 1 Then
                    frmKeywordChange.rstKeywordsThesaurus!LargerCategory = True
                End If
            frmKeywordChange.rstKeywordsThesaurus.Update
        
        Next
        frmKeywordChange.cnDatabase.CommitTrans
        
        
        frmKeywordChange.rstRemoteKeywordsThesaurus.Requery
        frmKeywordChange.cnRemoteDatabase.BeginTrans
        On Error GoTo data_Error
        For i = 1 To cThesaurusID.Count
            frmKeywordChange.rstRemoteKeywordsThesaurus.AddNew
                frmKeywordChange.rstRemoteKeywordsThesaurus!KeywordID = iKeywordID
                frmKeywordChange.rstRemoteKeywordsThesaurus!ThesaurusID = cThesaurusID.Item(i)
                If Me.chkAddAsCategory.Value = 1 Then
                    frmKeywordChange.rstRemoteKeywordsThesaurus!LargerCategory = True
                End If
            frmKeywordChange.rstRemoteKeywordsThesaurus.Update
        Next
        frmKeywordChange.cnRemoteDatabase.CommitTrans
        
        
        
        Set cThesaurusID = Nothing
        
    End If
                    
    Call Me.Fill_TT_List
    If iIndex >= Me.lstAllThesaurus.ListCount Then iIndex = Me.lstAllThesaurus.ListCount - 1
    Me.lstAllThesaurus.Selected(iIndex) = True
    Me.chkAddAsCategory.Value = 0
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

Private Sub lstThesaurusKeywords_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    
    
    
    sKeytext = Me.lstThesaurusKeywords.Text
    sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
    Set rstGetKeyNum = New ADODB.Recordset
    With rstGetKeyNum
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT KeywordID from tblKeywords where KeywordOrCodeSection ='" & sKeytext & "'")
    End With
    
    iKeywordID = rstGetKeyNum!KeywordID
    Set rstGetKeyNum = Nothing
    
    
    Me.txtKeywordID = iKeywordID
    
    Call Me.Fill_TT_List
    If Me.lstThesaurusThesaurus.ListCount > 0 Then Me.lstThesaurusThesaurus.Selected(0) = True

End Sub


Public Sub Fill_AT_List()

    Me.lstAllThesaurus.Clear
    
    frmKeywordChange.rstThesaurus.Requery
    If Not frmKeywordChange.rstThesaurus.EOF Then
        frmKeywordChange.rstThesaurus.MoveFirst
        Do While Not frmKeywordChange.rstThesaurus.EOF
            Me.lstAllThesaurus.AddItem frmKeywordChange.rstThesaurus!ThesaurusEquivalent
            frmKeywordChange.rstThesaurus.MoveNext
        Loop
    End If
End Sub

Private Sub lstThesaurusThesaurus_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim rstCategory As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    Dim i As Integer
    
    If Me.lstThesaurusThesaurus.SelCount = 0 Then
        Me.txtKeywordThesaurusID.Text = ""
    Else
        sKeytext = Me.lstThesaurusThesaurus.Text
        sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
        Set rstGetKeyNum = New ADODB.Recordset
        With rstGetKeyNum
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open ("SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sKeytext & "'")
        End With
        
        iKeywordID = rstGetKeyNum!ThesaurusEquivalentID
        Set rstGetKeyNum = Nothing
        
        
        Me.txtKeywordThesaurusID = iKeywordID
        
        Set rstCategory = New ADODB.Recordset
        With rstCategory
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open ("SELECT LargerCategory from tblKeywordThesaurus where KeywordID=" & Me.txtKeywordID & " AND ThesaurusID=" & Me.txtKeywordThesaurusID)
        End With
        If rstCategory!LargerCategory = False Then Me.chkLargerCategory.Value = False
        If rstCategory!LargerCategory = True Then Me.chkLargerCategory.Value = True
        
        Set rstCategory = Nothing
        For i = 0 To (Me.lstAllThesaurus.ListCount - 1)
            Me.lstAllThesaurus.Selected(i) = False
        Next
    End If
End Sub

Private Sub lstThesaurusThesaurus_DblClick()
    Dim rstQueryThesaurus As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim sItem As String
   
    iThesaurusID = Me.txtKeywordThesaurusID.Text
    iKeywordID = Me.txtKeywordID.Text
    
    On Error GoTo data_Error
    
    Set rstQueryThesaurus = New ADODB.Recordset
    With rstQueryThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywordThesaurus WHERE KeywordID=" & iKeywordID & " AND ThesaurusID=" & iThesaurusID)
        '.Open ("SELECT * from qryThesaurus WHERE KeywordID=" & iKeywordID)
    End With
    
    frmKeywordChange.cnDatabase.BeginTrans
        rstQueryThesaurus.Delete
        rstQueryThesaurus.Update
    frmKeywordChange.cnDatabase.CommitTrans
    
    Set rstQueryThesaurus = Nothing
    
    Set rstQueryThesaurus = New ADODB.Recordset
    With rstQueryThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnRemoteDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywordThesaurus WHERE KeywordID=" & iKeywordID & " AND ThesaurusID=" & iThesaurusID)
        '.Open ("SELECT * from qryThesaurus WHERE KeywordID=" & iKeywordID)
    End With
    
    frmKeywordChange.cnRemoteDatabase.BeginTrans
        rstQueryThesaurus.Delete
        rstQueryThesaurus.Update
    frmKeywordChange.cnRemoteDatabase.CommitTrans
    
    
    Set rstQueryThesaurus = Nothing
    
    
    frmKeywordChange.rstKeywordsThesaurus.Requery
    Call Me.Fill_TT_List
    
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
