VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmThesaurusEntry 
   Caption         =   "Keyword Folding"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   ScaleHeight     =   9480
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdKeywordFolding 
      Caption         =   "Keyword Folding"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Keyword/Thesaurus Entry"
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteThesaurus 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditThesaurus 
      Caption         =   "Edit"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddThesaurus 
      Caption         =   "Add"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ListBox lstThesaurus 
      DataField       =   "ThesaurusEquivalent"
      Height          =   4155
      ItemData        =   "frmThesaurusEntry.frx":0000
      Left            =   2880
      List            =   "frmThesaurusEntry.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
   End
   Begin MSForms.ToggleButton tglTTable 
      Height          =   735
      Left            =   7380
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      VariousPropertyBits=   746588185
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2990;1296"
      Value           =   "1"
      Caption         =   "Thesaurus Table Entry"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmThesaurusEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstThesaurusFill As ADODB.Recordset
Dim rstThesaurusLookup As ADODB.Recordset
Dim iCurrentRecord As String


Private Sub cmdAddThesaurus_Click()
    frmAddThesaurus.txtThesaurus.Text = ""
    frmAddThesaurus.Caption = "Add Thesaurus Equivalent"
    frmAddThesaurus.Show
End Sub

Private Sub cmdCheck_Click()
    Me.Hide
    frmKeywordThesaurusChange.Show
    Call frmKeywordThesaurusChange.Fill_KT_List
End Sub

Private Sub cmdDeleteThesaurus_Click()
    If Me.lstThesaurus.SelCount = 0 Then
        MsgBox "You must select a word to delete."
    Else
        On Error GoTo data_Error
        Call Get_Current_RecNum
        rstThesaurusFill.MoveFirst
        Do While rstThesaurusFill!ThesaurusEquivalentID <> iCurrentRecord
            rstThesaurusFill.MoveNext
        Loop
        frmKeywordChange.cnDatabase.BeginTrans
            rstThesaurusFill.Delete
            rstThesaurusFill.Update
        frmKeywordChange.cnDatabase.CommitTrans
        Call Me.fill_list
        Call frmKeywordThesaurusChange.Fill_TT_List
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

Private Sub cmdEditThesaurus_Click()
    If Me.lstThesaurus.SelCount = 0 Then
        MsgBox "You must select a word to edit."
    Else
        frmAddThesaurus.Caption = "Edit Thesaurus Equivalent"
        frmAddThesaurus.txtThesaurus.Text = Me.lstThesaurus.Text
        Call Get_Current_RecNum
        frmAddThesaurus.txtRecNum.Text = iCurrentRecord
        frmAddThesaurus.Show
    End If
'
End Sub


Private Sub cmdKeywordFolding_Click()
    frmKeywordChange.Show
End Sub

Private Sub cmdStack_Click()
    Me.Hide
    frmStackEntry.Show
    Call frmStackEntry.Fill_ST_List
End Sub

Private Sub Form_Load()

    Set rstThesaurusFill = New ADODB.Recordset
'    Set cKeywordID = New Collection
    With rstThesaurusFill
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusEquivalent")
    End With
    Call fill_list
End Sub

Public Sub fill_list()

'    Dim iIndex As Integer
'    iIndex = 0
    lstThesaurus.Clear
    rstThesaurusFill.Requery
    rstThesaurusFill.MoveFirst
    If Not rstThesaurusFill.EOF Then
        rstThesaurusFill.MoveFirst
        Do While Not rstThesaurusFill.EOF
            lstThesaurus.AddItem rstThesaurusFill!ThesaurusEquivalent
            'iIndex = rstKeywords!KeywordID
'            cKeywordID.Add iIndex ', iIndex
            rstThesaurusFill.MoveNext
'            'iIndex = iIndex + 1
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstThesaurusFill = Nothing
End Sub

Private Sub Get_Current_RecNum()
    Dim sTE As String
    
    sTE = Me.lstThesaurus.Text
    Set rstThesaurusLookup = New ADODB.Recordset
    With rstThesaurusLookup
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT ThesaurusEquivalentID from tblThesaurusEquivalent where ThesaurusEquivalent='" & sTE & "'")
    End With
    iCurrentRecord = rstThesaurusLookup!ThesaurusEquivalentID
    Set rstThesaurusLookup = Nothing
End Sub
