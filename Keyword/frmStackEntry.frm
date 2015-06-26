VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmStackEntry 
   Caption         =   "Thesaurus Stack Entry"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdKeywordFolding 
      Caption         =   "Keyword Folding"
      Height          =   735
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenThesaurus 
      Caption         =   "Thesaurus Table Entry"
      Height          =   735
      Left            =   7380
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdKT 
      Caption         =   "Keyword/Thesaurus Entry"
      Height          =   735
      Left            =   5040
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteStack 
      Caption         =   "Delete Stack"
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditStack 
      Caption         =   "Edit Stack"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddStack 
      Caption         =   "Add Stack"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox lstStacks 
      Height          =   2595
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.ListBox lstStackThesaurus 
      Height          =   2595
      Left            =   4560
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtStackID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtAllThesaurusID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox lstAllThesaurus 
      Height          =   2595
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtKeywordThesaurusID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin MSForms.ToggleButton tglThesaurusStack 
      Height          =   735
      Left            =   9720
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      VariousPropertyBits=   746588185
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2990;1296"
      Value           =   "1"
      Caption         =   "Thesaurus Stack Entry"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblStacks 
      Caption         =   "Stacks"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblKeywordID 
      Caption         =   "Stack ID"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblThesaurus 
      Caption         =   "Thesaurus"
      Height          =   255
      Left            =   9960
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblThesaurusID 
      Caption         =   "Thesaurus ID"
      Height          =   255
      Left            =   11160
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblDblClick 
      Caption         =   "<---------------------> Double-Click to Add/Remove Thesaurus Equivalents for Selected Keyword         <-------------------->"
      Height          =   1695
      Left            =   8640
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Thesaurus ID"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Thesaurus"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmStackEntry"
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
        iKeywordID = Me.txtStackID.Text
        
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

Private Sub cmdGenThesaurus_Click()
    Me.Hide
    frmThesaurusEntry.Show
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

Private Sub cmdKT_Click()
    Me.Hide
    frmKeywordThesaurusChange.Show
    Call frmKeywordThesaurusChange.Fill_KT_List
    
    'Call frmKeywordThesaurusChange.Fill_TT_List
    
    
End Sub

Private Sub Form_Load()
    Call Fill_Stack_List
    Call Fill_AT_List
    If Me.lstStacks.ListCount > 0 Then Me.lstStacks.Selected(0) = True
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

Public Sub Fill_ST_List()
    Dim rstQueryThesaurus As ADODB.Recordset
    Dim sCurrentItem As String
    Dim i As Integer
    Dim iThesaurusID As Integer
    Dim iStackID As Integer
    
    Set rstQueryThesaurus = New ADODB.Recordset
    If Me.txtStackID.Text <> "" Then iStackID = Me.txtStackID.Text
    
    'rstquerythesaurus.Open ,,,
    With rstQueryThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT * from qryThesaurusStack WHERE StackID=" & iStackID)
    End With
    
    Call Me.Fill_AT_List
    Me.lstStackThesaurus.Clear
    If Not rstQueryThesaurus.EOF Then
        rstQueryThesaurus.MoveFirst
        'Set cThesaurusID = New Collection
        Do While Not rstQueryThesaurus.EOF
            sCurrentItem = rstQueryThesaurus!ThesaurusEquivalent
            Me.lstStackThesaurus.AddItem sCurrentItem
            For i = 0 To (Me.lstAllThesaurus.ListCount - 1)
                If Me.lstAllThesaurus.List(i) = sCurrentItem Then
                    Me.lstAllThesaurus.RemoveItem (i)
                End If
            Next
            iThesaurusID = rstQueryThesaurus!ThesaurusEquivalentID
            'cThesaurusID.Add iThesaurusID
            rstQueryThesaurus.MoveNext
        Loop
    End If
    Set rstQueryThesaurus = Nothing

End Sub
    
Public Sub Fill_Stack_List()
    
    frmKeywordChange.rstStacks.Requery
    Me.lstStacks.Clear
    If Not frmKeywordChange.rstStacks.EOF Then
        frmKeywordChange.rstStacks.MoveFirst
        Do While Not frmKeywordChange.rstStacks.EOF
            Me.lstStacks.AddItem frmKeywordChange.rstStacks!StackName
            'iIndex = rstKeywords!KeywordID
            'cKeywordID.Add iIndex ', iIndex
            frmKeywordChange.rstStacks.MoveNext
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
    
    If Me.lstAllThesaurus.SelCount = 0 Then
        Me.txtAllThesaurusID.Text = ""
    Else
        sKeytext = Me.lstAllThesaurus.Text
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
        
        
        Me.txtAllThesaurusID = iKeywordID
        
        For i = 0 To (Me.lstStackThesaurus.ListCount - 1)
            Me.lstStackThesaurus.Selected(i) = False
        Next
    End If
End Sub

Private Sub lstAllThesaurus_DblClick()
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim iIndex As Integer
    iIndex = Me.lstAllThesaurus.ListIndex
    iKeywordID = Me.txtStackID
    iThesaurusID = Me.txtAllThesaurusID
    frmKeywordChange.rstStackJunction.Requery
    frmKeywordChange.cnDatabase.BeginTrans
    On Error GoTo data_Error
    
        frmKeywordChange.rstStackJunction.AddNew
            frmKeywordChange.rstStackJunction!ThesaurusStackID = iKeywordID
            frmKeywordChange.rstStackJunction!ThesaurusID = iThesaurusID
        frmKeywordChange.rstStackJunction.Update
    frmKeywordChange.cnDatabase.CommitTrans
    
    Call Me.Fill_ST_List
    If iIndex >= Me.lstAllThesaurus.ListCount Then iIndex = Me.lstAllThesaurus.ListCount - 1
    Me.lstAllThesaurus.Selected(iIndex) = True
    
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

Private Sub lstStacks_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    
    
    
    sKeytext = Me.lstStacks.Text
    sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
    Set rstGetKeyNum = New ADODB.Recordset
    With rstGetKeyNum
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT StackID from tblThesaurusStack where StackName='" & sKeytext & "'")
    End With
    
    iKeywordID = rstGetKeyNum!StackID
    Set rstGetKeyNum = Nothing
    
    
    Me.txtStackID = iKeywordID
    
    Call Me.Fill_ST_List

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

Private Sub lstStackThesaurus_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    Dim i As Integer
    
    If Me.lstStackThesaurus.SelCount = 0 Then
        Me.txtKeywordThesaurusID.Text = ""
    Else
        sKeytext = Me.lstStackThesaurus.Text
        sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
        Set rstGetKeyNum = New ADODB.Recordset
        With rstGetKeyNum
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open ("SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sKeytext & "'")
        End With
        
        iKeywordID = rstGetKeyNum!ThesaurusEquivalentID
        Set rstGetKeyNum = Nothing
        
        
        Me.txtKeywordThesaurusID = iKeywordID
        
        
        For i = 0 To (Me.lstAllThesaurus.ListCount - 1)
            Me.lstAllThesaurus.Selected(i) = False
        Next
    End If
End Sub

Private Sub lstStackThesaurus_DblClick()
    Dim rstQueryThesaurus As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim sItem As String
   
    Set rstQueryThesaurus = New ADODB.Recordset
    iThesaurusID = Me.txtKeywordThesaurusID.Text
    iKeywordID = Me.txtStackID.Text
    
    With rstQueryThesaurus
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusStackJunction WHERE ThesaurusStackID=" & iKeywordID & " AND ThesaurusID=" & iThesaurusID)
        '.Open ("SELECT * from qryThesaurus WHERE KeywordID=" & iKeywordID)
        
    End With
    
    frmKeywordChange.cnDatabase.BeginTrans
        On Error GoTo data_Error
        
        rstQueryThesaurus.Delete
        rstQueryThesaurus.Update
        
    frmKeywordChange.cnDatabase.CommitTrans
    
    Set rstQueryThesaurus = Nothing
    frmKeywordChange.rstStackJunction.Requery
    Call Me.Fill_ST_List
    
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

Private Sub cmdAddStack_Click()
    frmAddThesaurus.txtThesaurus.Text = "**"
    frmAddThesaurus.Caption = "Add Thesaurus Stack"
    frmAddThesaurus.Show
End Sub

Private Sub cmdDeleteStack_Click()
    Dim iStackID As Integer
    Dim iIndex As Integer
    iIndex = Me.lstStacks.ListIndex
    If Me.lstStacks.SelCount = 0 Then
        MsgBox "You must select a word to delete."
    Else
        On Error GoTo data_Error
        'Call Get_Current_RecNum
        iStackID = Me.txtStackID.Text
        frmKeywordChange.rstStacks.Requery
        
        frmKeywordChange.rstStacks.MoveFirst
        Do While frmKeywordChange.rstStacks!StackID <> iStackID
            frmKeywordChange.rstStacks.MoveNext
        Loop
        frmKeywordChange.cnDatabase.BeginTrans
            frmKeywordChange.rstStacks.Delete
            frmKeywordChange.rstStacks.Update
        frmKeywordChange.cnDatabase.CommitTrans
        'Call Me.fill_list
        'Call frmKeywordThesaurusChange.Fill_TT_List
        Call Me.Fill_Stack_List
        
        If iIndex >= Me.lstStacks.ListCount Then iIndex = Me.lstStacks.ListCount - 1
        Me.lstStacks.Selected(iIndex) = True
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

Private Sub cmdEditStack_Click()
    If Me.lstStacks.SelCount = 0 Then
        MsgBox "You must select a word to edit."
    Else
        frmAddThesaurus.Caption = "Edit Thesaurus Stack"
        frmAddThesaurus.txtThesaurus.Text = Me.lstStacks.Text
        'Call Get_Current_RecNum
        frmAddThesaurus.txtRecNum.Text = Me.txtStackID.Text
        frmAddThesaurus.Show
    End If
'
End Sub

Private Sub Get_Current_RecNum()
    Dim sTE As String
    
    sTE = Me.lstThesaurus.Text
    Set rstThesaurusLookup = New ADODB.Recordset
    With rstThesaurusLookup
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sTE & "'")
    End With
    iCurrentRecord = rstThesaurusLookup!ThesaurusEquivalentID
    Set rstThesaurusLookup = Nothing
End Sub

