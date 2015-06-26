VERSION 5.00
Begin VB.Form frmKeywordThesaurus 
   Caption         =   "Modify Keywords and Thesaurus Equivalents"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdModifyThesaurus 
      Caption         =   "Modify Thesaurus Equivalent for Selected Keyword"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdModifyKeyword 
      Caption         =   "Modify Keyword"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteThesaurus 
      Caption         =   "Delete Thesaurus Equivalent for Selected Keyword"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeleteKeyword 
      Caption         =   "Delete Keyword"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddThesaurus 
      Caption         =   "Add Thesaurus Equivalent for Selected Keyword"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddKeyword 
      Caption         =   "Add Keyword"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lstThesaurus 
      Height          =   2790
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.ListBox lstKeywords 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblKeywords 
      Caption         =   "Keywords"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmKeywordThesaurus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstKeywords As ADODB.Recordset
Dim rstThesaurus As ADODB.Recordset
Dim cKeywordID As Collection
    

Private Sub Form_Load()
    Dim iIndex As Integer
    Set rstKeywords = New ADODB.Recordset
    'Set rstThesaurus = New ADODB.Recordset
    Set cKeywordID = New Collection
    With rstKeywords
        .ActiveConnection = frmMain.cnWriteDatabase
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywords")
    End With
    iIndex = 0
    If Not rstKeywords.EOF Then
        rstKeywords.MoveFirst
        Do While Not rstKeywords.EOF
            lstKeywords.AddItem rstKeywords!keywordorcodesection
            iIndex = rstKeywords!KeywordID
            cKeywordID.Add iIndex ', iIndex
            rstKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstKeywords = Nothing
    Set rstThesaurus = Nothing
    Set cKeywordID = Nothing
End Sub

Private Sub lstKeywords_Click()
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    iItemnumber = lstKeywords.ListIndex
    iKeywordID = cKeywordID.Item(iItemnumber + 1)
    sItem = lstKeywords.List(iItemnumber)
    Me.txtKeywordID = iKeywordID
    Set rstThesaurus = Nothing
    Set rstThesaurus = New ADODB.Recordset
    
    With rstThesaurus
        .ActiveConnection = frmMain.cnWriteDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from qryThesaurus WHERE KEYWORDID=" & iKeywordID)
    End With
    'rstKeywords.MoveFirst
    'rstKeywords.Find ("rstKeywords!keywordorcodesection = sItem")
    'sItem = rstKeywords!KeywordID
    'MsgBox rstkeywords!KeywordID lstKeywords.List(itemnumber)
    lstThesaurus.Clear
    If Not rstThesaurus.EOF Then
        rstThesaurus.MoveFirst
        Do While Not rstThesaurus.EOF
            lstThesaurus.AddItem rstThesaurus!ThesaurusEquivalent
            rstThesaurus.MoveNext
        Loop
    End If
End Sub
