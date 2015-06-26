VERSION 5.00
Begin VB.Form frmEditThesaurus 
   Caption         =   "New Equivalent for Keyword"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStack 
      Caption         =   "Thesaurus Stack Entry ->"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewEquiv 
      Caption         =   "New Thesaurus Equivalent ->"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cmbThesaurusEquiv 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   8655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditThesaurus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstThesaurusEq As ADODB.Recordset

Private Sub cmdAdd_Click()
    Dim i As Integer
    Dim iKeywordID As Integer
    Dim iThesaurusID As Integer
    Dim rstTestEOF As Recordset
    Dim rstQryThesaurus As Recordset
    Dim sTempString As String
    sTempString = ""
    If Me.cmbThesaurusEquiv.Text = "" Then
        MsgBox "You need to type something"
        Me.cmbThesaurusEquiv.SetFocus
    Else
        On Error GoTo data_Error
        sTempString = Me.cmbThesaurusEquiv.Text
        Set rstTestEOF = New ADODB.Recordset
        With rstTestEOF
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblThesaurusEquivalent WHERE ThesaurusEquivalent='" & sTempString & "'")
        End With
        'rstThesaurusEq.MoveFirst
        'Do Until ((rstThesaurusEq!ThesaurusEquivalent = sTempString) Or (rstThesaurusEq.EOF))
        '    rstThesaurusEq.MoveNext
        'Loop
        frmKeywordChange.cnDatabase.BeginTrans
        If rstTestEOF.EOF Then
            rstThesaurusEq.AddNew
            rstThesaurusEq!ThesaurusEquivalent = sTempString
            
            rstThesaurusEq.Update
            
            iThesaurusID = rstThesaurusEq!ThesaurusEquivalentID
            
        Else
            iThesaurusID = rstTestEOF!ThesaurusEquivalentID
        End If
        'iThesaurusID = rstThesaurusEq!thesaurusID
        iKeywordID = frmKeywordThesaurusChange.txtKeywordID.Text
        Set rstQryThesaurus = New Recordset
        With rstQryThesaurus
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from qryThesaurus WHERE KeywordID=" & iKeywordID & " AND ThesaurusID=" & iThesaurusID)
        End With
        If Not rstQryThesaurus.EOF Then
            MsgBox "This thesaurus term already exists for this keyword."
            Me.Hide
            Set rstQryThesaurus = Nothing
            'Set rstThesaurusEq = Nothing
        Else
            rstQryThesaurus.AddNew
                rstQryThesaurus!KeywordID = iKeywordID
                rstQryThesaurus!ThesaurusID = iThesaurusID
            rstQryThesaurus.Update
        End If
        frmKeywordChange.cnDatabase.CommitTrans
        
        Set rstQryThesaurus = Nothing
        Set rstTestEOF = Nothing
        Call frmKeywordThesaurusChange.Fill_Thesaurus_List
        'Me.Hide
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

Private Sub cmdNewEquiv_Click()
    frmAddThesaurus.Show
End Sub

Private Sub Form_Load()
    
    Set rstThesaurusEq = New ADODB.Recordset
    With rstThesaurusEq
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusEquivalent")
    End With
    
    Call Fill_Thesaurus_Combo

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstThesaurusEq = Nothing
End Sub

Public Sub Fill_Thesaurus_Combo()

Dim sTempstr As String

rstThesaurusEq.MoveFirst
Me.cmbThesaurusEquiv.Clear
Do While Not rstThesaurusEq.EOF
        sTempstr = rstThesaurusEq!ThesaurusEquivalent
        Me.cmbThesaurusEquiv.AddItem sTempstr
        rstThesaurusEq.MoveNext
Loop
End Sub
