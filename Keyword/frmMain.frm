VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab MainTabs 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "Keyword Manipulation"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOldKeywordID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblKeywordID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblKeywords"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lboResults"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAddNewKeyword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditKeyword"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optDelete"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optConvert"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtOldKeywordID"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdStart"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtKeywordID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstThesaurus"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lstKeywords"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdDeleteKeyword"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Old/New Keyword Comparison"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Edit Keyword/Thesaurus Equivalents"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblThesaurusID"
      Tab(2).Control(1)=   "lblThesaurus"
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(4)=   "cmdThesaurus"
      Tab(2).Control(5)=   "txtThesaurusID"
      Tab(2).Control(6)=   "cmdDelete"
      Tab(2).Control(7)=   "cmdAdd"
      Tab(2).Control(8)=   "Text1"
      Tab(2).Control(9)=   "List1"
      Tab(2).Control(10)=   "List2"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Edit Thesaurus Table"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstThesaurusEquivalents"
      Tab(3).Control(1)=   "cmdAddThesaurus"
      Tab(3).Control(2)=   "cmdEditThesaurus"
      Tab(3).Control(3)=   "cmdDeleteThesaurus"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Thesaurus Stacks"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.CommandButton cmdDeleteKeyword 
         Caption         =   "Delete Keyword Completely"
         Height          =   615
         Left            =   4800
         TabIndex        =   34
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ListBox lstThesaurusEquivalents 
         DataField       =   "ThesaurusEquivalent"
         Height          =   3375
         ItemData        =   "frmMain.frx":008C
         Left            =   -72240
         List            =   "frmMain.frx":0093
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   1800
         Width           =   5895
      End
      Begin VB.CommandButton cmdAddThesaurus 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71880
         TabIndex        =   32
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditThesaurus 
         Caption         =   "Edit"
         Height          =   495
         Left            =   -69840
         TabIndex        =   31
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteThesaurus 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -67920
         TabIndex        =   30
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   8280
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Height          =   2790
         Left            =   -74400
         TabIndex        =   20
         Top             =   2760
         Width           =   4215
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   -69480
         TabIndex        =   19
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73200
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Thesaurus Equivalent"
         Height          =   495
         Left            =   -72960
         TabIndex        =   17
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Thesaurus Equivalent"
         Height          =   495
         Left            =   -68040
         TabIndex        =   16
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtThesaurusID 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68280
         TabIndex        =   15
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdThesaurus 
         Caption         =   "Thesaurus Manipulation"
         Height          =   495
         Left            =   -70080
         TabIndex        =   14
         Top             =   7440
         Width           =   1215
      End
      Begin VB.ListBox lstKeywords 
         Height          =   2790
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ListBox lstThesaurus 
         Height          =   2790
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtKeywordID 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Conversion"
         Height          =   735
         Left            =   7080
         TabIndex        =   7
         Top             =   5160
         Width           =   3255
      End
      Begin VB.TextBox txtOldKeywordID 
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optConvert 
         Caption         =   "Convert Deleted Keyword into Thesaurus Equivalent"
         Height          =   495
         Left            =   7080
         TabIndex        =   5
         Top             =   4200
         Width           =   3015
      End
      Begin VB.OptionButton optDelete 
         Caption         =   "Simply delete keyword; no conversion"
         Height          =   495
         Left            =   7080
         TabIndex        =   4
         Top             =   4680
         Width           =   3135
      End
      Begin VB.CommandButton cmdEditKeyword 
         Caption         =   "Edit Selected Keyword"
         Height          =   615
         Left            =   4800
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNewKeyword 
         Caption         =   "Add New Keyword"
         Height          =   615
         Left            =   4800
         TabIndex        =   2
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox lboResults 
         Height          =   1620
         Left            =   480
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   6480
         Width           =   11055
      End
      Begin VB.Label Label2 
         Caption         =   "Keywords"
         Height          =   255
         Left            =   -74400
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Keyword ID"
         Height          =   255
         Left            =   -71880
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblThesaurus 
         Caption         =   "Thesaurus"
         Height          =   255
         Left            =   -69480
         TabIndex        =   22
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblThesaurusID 
         Caption         =   "Thesaurus ID"
         Height          =   255
         Left            =   -66960
         TabIndex        =   21
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblKeywords 
         Caption         =   "Double Click to remove from list and add to Thesaurus. Then select which keyword should remain and be master keyword"
         Height          =   855
         Left            =   840
         TabIndex        =   13
         Top             =   4080
         Width           =   3495
      End
      Begin VB.Label lblKeywordID 
         Caption         =   "Keyword ID"
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblOldKeywordID 
         Caption         =   "Old Keyword ID"
         Height          =   495
         Left            =   8520
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Old Keyword ID"
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   5400
      TabIndex        =   28
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Records"
      Height          =   495
      Left            =   5400
      TabIndex        =   27
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5400
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstKeywords As ADODB.Recordset
Dim rstThesaurus As ADODB.Recordset
Dim rstRecordsKeywords As ADODB.Recordset
Dim rstKeywordsThesaurus As ADODB.Recordset
Dim cKeywordID As Collection
Public cnDatabase As ADODB.Connection
Dim iSelectedKeywordID As Integer
Dim iToRemoveKeywordID As Integer
Dim sSelectedKeyword As String
Dim sToRemoveKeyword As String
Dim iRemovedItemNumber As Integer
Public avardata As Variant
Public rstThesaurusFill As ADODB.Recordset
Dim rstThesaurusLookup As ADODB.Recordset
Dim iCurrentRecord As String

Private Sub cmdAddNewKeyword_Click()
    frmEditKeywords.cmdAdd.Caption = "Add"
    frmEditKeywords.Caption = "Add Keywords"
    'frmEditKeywords.txtKeyword = Me.lstKeywords.List(lstKeywords.ListIndex)
    frmEditKeywords.Show
    frmEditKeywords.txtKeyword.SetFocus
End Sub

Private Sub cmdCheck_Click()
    Dim iSelIndex As Integer
    frmKeywordThesaurusChange.Show
    Call frmKeywordThesaurusChange.Fill_KT_List
    iSelIndex = Me.lstKeywords.ListIndex
    frmKeywordThesaurusChange.lstKeywords.ListIndex = iSelIndex
End Sub

Private Sub cmdEditKeyword_Click()
    If (Me.lstKeywords.SelCount = 0) Then
        MsgBox "You need to select a keyword first"
    Else
        frmEditKeywords.cmdAdd.Caption = "Edit"
        frmEditKeywords.Caption = "Edit Keywords"
        frmEditKeywords.txtKeyword = Me.lstKeywords.List(lstKeywords.ListIndex)
        frmEditKeywords.Show
        frmEditKeywords.txtKeyword.SetFocus
    End If
End Sub

Private Sub cmdStart_Click()
    Dim iThesaurusID As Integer
    Dim bNoDuplicate As Boolean
    Dim rstRKTest As Recordset
    Dim cRecNums As Collection
    Dim cThesaurusEquivs As Collection
    Dim rstOldKeywordEquivs As Recordset
    Dim rstKeyWordThesaurusOld As Recordset
    Dim bDuplicateThesaurus As Boolean
    Dim i As Integer
    Dim iTempID As Integer
    Dim itmpInt As Integer
    
'check to see if all sections made
    If (Me.lstKeywords.SelCount = 0) Or (Me.lstThesaurus.ListCount = 0) Or _
        ((Me.optConvert = False) And (Me.optDelete = False)) Then
        MsgBox "You did not make all necessary selections", vbCritical, "Error"
        'Cancel = True
    Else
    
'put thesaurus equivalents of keyword to be deleted in a collection
        Set rstOldKeywordEquivs = New Recordset
        With rstOldKeywordEquivs
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT DISTINCT ThesaurusID from tblKeywordThesaurus where KeywordID=" & iToRemoveKeywordID)
        End With
        Set cThesaurusEquivs = New Collection
            
        If Not rstOldKeywordEquivs.EOF Then
            For i = 1 To rstOldKeywordEquivs.RecordCount
                iTempID = rstOldKeywordEquivs!ThesaurusID
                cThesaurusEquivs.Add iTempID
                rstOldKeywordEquivs.MoveNext
            Next
        End If
        Set rstOldKeywordEquivs = Nothing
        
'begin the transaction
        cnDatabase.BeginTrans
        On Error GoTo data_Error
        
'open a rstRecordsKeywords recordset of to get record numbers that will be affected by deletion
        Set rstRecordsKeywords = New Recordset
        With rstRecordsKeywords
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords where KeywordID=" & iToRemoveKeywordID)
        End With
        
'set up a test to see if any of the converted thesaurus equivalents would duplicate a current thesaurus quivalent of the selected Keyword
        Set rstRKTest = New Recordset
        Set cRecNums = New Collection
        With rstRKTest
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords where KeywordID=" & iSelectedKeywordID)
        End With
        
        Do While Not rstRKTest.EOF
            'If rstRecordsKeywords!KeywordID = iSelectedKeywordID Then bNoDuplicate = False
            itmpInt = rstRKTest!recordID
            cRecNums.Add itmpInt
            rstRKTest.MoveNext
        Loop
        
        'Set rstRKTest = Nothing
        bNoDuplicate = True
        'If Not rstRecordsKeywords.EOF Then
            Do While Not rstRecordsKeywords.EOF
                For i = 1 To cRecNums.Count
                    If rstRecordsKeywords!recordID = cRecNums.Item(i) Then bNoDuplicate = False
                Next
                If bNoDuplicate Then
                    rstRecordsKeywords!KeywordID = iSelectedKeywordID
                    rstRecordsKeywords.Update
                End If
                bNoDuplicate = True
                rstRecordsKeywords.MoveNext
            Loop
            rstRecordsKeywords.MoveFirst
        'End If
        
        'Set cRecNums = Nothing
'change keyword ID of keyword to be removed to selected keyword; this removes the keyword that was chosen to be removed
        'Do While Not rstRecordsKeywords.EOF
            'If bNoDuplicate Then
            '    rstRecordsKeywords!KeywordID = iSelectedKeywordID
            '    rstRecordsKeywords.Update
            'End If
        '    rstRecordsKeywords.MoveNext
        'Loop
        
        
        'Do While Not rstRecordsKeywords.EOF
        '    If bNoDuplicate Then
        '        rstRecordsKeywords!KeywordID = iSelectedKeywordID
        '        rstRecordsKeywords.Update
        '    End If
        '    rstRecordsKeywords.MoveNext
        'Loop
        
        Set rstRecordsKeywords = Nothing

'delete the keyword itself from the keyword table
        rstKeywords.MoveFirst
        Do Until rstKeywords!KeywordID = iToRemoveKeywordID
            rstKeywords.MoveNext
        Loop
        rstKeywords.Delete
        rstKeywords.Update
        
'put the thesaurus equivalents of the deleted keyword as thesaurus equivalents of selected keyword
        Set rstKeyWordThesaurusOld = New Recordset
        With rstKeyWordThesaurusOld
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblKeywordThesaurus where KeywordID=" & iSelectedKeywordID)
        End With
        If rstKeyWordThesaurusOld.EOF Then
            For i = 1 To cThesaurusEquivs.Count
                rstKeyWordThesaurusOld.AddNew
                    rstKeyWordThesaurusOld!KeywordID = iSelectedKeywordID
                    rstKeyWordThesaurusOld!ThesaurusID = cThesaurusEquivs.Item(i)
                rstKeyWordThesaurusOld.Update
            Next
        Else
            For i = 1 To cThesaurusEquivs.Count
                rstKeyWordThesaurusOld.MoveFirst
                bDuplicateThesaurus = False
                Do While (Not rstKeyWordThesaurusOld.EOF)
                    If rstKeyWordThesaurusOld!ThesaurusID = cThesaurusEquivs.Item(i) Then bDuplicateThesaurus = True
                    rstKeyWordThesaurusOld.MoveNext
                Loop
                If Not bDuplicateThesaurus Then
                    rstKeyWordThesaurusOld.AddNew
                        rstKeyWordThesaurusOld!KeywordID = iSelectedKeywordID
                        rstKeyWordThesaurusOld!ThesaurusID = cThesaurusEquivs.Item(i)
                    rstKeyWordThesaurusOld.Update
                End If
            Next
        End If
        
'this sees if option to convert deleted keyword to a thesaurus equivalent is true; if so, makes the conversion
        If Me.optConvert = True Then
            rstThesaurus.MoveFirst
            Do While (Not rstThesaurus.EOF)
                If (Not rstThesaurus!ThesaurusEquivalent = sToRemoveKeyword) Then rstThesaurus.MoveNext Else GoTo exit_loop1
            Loop
            
exit_loop1:
            If rstThesaurus.EOF Then
                rstThesaurus.AddNew
                rstThesaurus!ThesaurusEquivalent = sToRemoveKeyword
                rstThesaurus.Update
                
            End If
            iThesaurusID = rstThesaurus!ThesaurusEquivalentID
            Set rstThesaurus = Nothing
            
            Set rstKeywordsThesaurus = New Recordset
            With rstKeywordsThesaurus
                .ActiveConnection = cnDatabase
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblKeywordThesaurus")
            End With
            Do While (Not rstKeywordsThesaurus.EOF)
                If (Not rstKeywordsThesaurus!ThesaurusID = iThesaurusID) Then rstKeywordsThesaurus.MoveNext Else GoTo exit_loop2
            Loop
exit_loop2:
            If rstKeywordsThesaurus.EOF Then
                rstKeywordsThesaurus.AddNew
                    rstKeywordsThesaurus!KeywordID = iSelectedKeywordID
                    rstKeywordsThesaurus!ThesaurusID = iThesaurusID
                rstKeywordsThesaurus.Update
            End If
            Set rstKeywordsThesaurus = Nothing
        End If

    
        iSelectedKeywordID = 0
        iToRemoveKeywordID = 0
        sSelectedKeyword = ""
        sToRemoveKeyword = ""
        Me.lstThesaurus.Clear
        Me.optConvert = False
        Me.optDelete = False
        cnDatabase.CommitTrans
    End If
data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "BuildRecordset"
             cnDatabase.RollbackTrans
        End Select
    End If
End Sub

Private Sub Command1_Click()
    frmMain.Show
End Sub

Private Sub Form_Load()
    Dim sConnectionString As String
    
    Set cnDatabase = New Connection
    sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\database\NCPL.mdb"
    cnDatabase.Open (sConnectionString)
    Set rstKeywords = New ADODB.Recordset
    'Set rstThesaurus = New ADODB.Recordset
    'rstKeywords.Open , , , adLockBatchOptimistic
    With rstKeywords
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywords")
    End With
    
    Set rstThesaurus = New Recordset
            With rstThesaurus
                .ActiveConnection = cnDatabase
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblThesaurusEquivalent")
            End With
    Set rstThesaurusFill = New ADODB.Recordset
'    Set cKeywordID = New Collection
    With rstThesaurusFill
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusEquivalent")
    End With
    Call fill_list
    Call Fill_Keyword_List
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstKeywords = Nothing
    Set rstThesaurus = Nothing
    Set cKeywordID = Nothing
    Set cnDatabase = Nothing
    
    Set rstThesaurusFill = Nothing

End Sub

Private Sub lstKeywords_Click()
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    iItemnumber = lstKeywords.ListIndex
    iKeywordID = cKeywordID.Item(iItemnumber + 1)
    sItem = lstKeywords.List(iItemnumber)
    Me.txtKeywordID = iKeywordID
    iSelectedKeywordID = iKeywordID
    Call BuildRecordset(Str(iKeywordID))
    'Set rstThesaurus = Nothing
    'Set rstThesaurus = New ADODB.Recordset
    
    'With rstThesaurus
    '    .ActiveConnection = cnDatabase
    '    .CursorType = adOpenKeyset
    '    .LockType = adLockOptimistic
    '    .Open ("SELECT * from qryThesaurus WHERE KEYWORDID=" & iKeywordID)
    'End With
    'rstKeywords.MoveFirst
    'rstKeywords.Find ("rstKeywords!keywordorcodesection = sItem")
    'sItem = rstKeywords!KeywordID
    'MsgBox rstkeywords!KeywordID lstKeywords.List(itemnumber)
    'lstThesaurus.Clear
    'If Not rstThesaurus.EOF Then
    '    rstThesaurus.MoveFirst
    '    Do While Not rstThesaurus.EOF
    '        lstThesaurus.AddItem rstThesaurus!ThesaurusEquivalent
    '        rstThesaurus.MoveNext
    '    Loop
    'End If
End Sub

Private Sub lstKeywords_DblClick()
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    If Me.lstThesaurus.ListCount = 0 Then
        iItemnumber = lstKeywords.ListIndex
        iKeywordID = cKeywordID.Item(iItemnumber + 1)
        sItem = lstKeywords.List(iItemnumber)
        'Me.txtKeywordID = iKeywordID
        lstKeywords.RemoveItem (iItemnumber)
        iRemovedItemNumber = iItemnumber
        cKeywordID.Remove (iItemnumber + 1)
        lstThesaurus.AddItem (sItem)
        Me.txtOldKeywordID = iKeywordID
        iToRemoveKeywordID = iKeywordID
        sToRemoveKeyword = sItem
    End If
End Sub

Private Sub lstThesaurus_DblClick()
    Dim iItemnumber As Integer
    Dim sItem As String
    Dim iRemovedID As Integer
    If lstThesaurus.SelCount > 0 Then
        iItemnumber = lstThesaurus.ListIndex
        sItem = lstThesaurus.List(iItemnumber)
        iRemovedID = Me.txtOldKeywordID
        Me.txtOldKeywordID = ""
        lstThesaurus.RemoveItem (iItemnumber)

        lstKeywords.AddItem sItem, iRemovedItemNumber
        cKeywordID.Add iRemovedID, , , iRemovedItemNumber
    End If
End Sub

Public Sub Fill_Keyword_List()
    Dim iIndex As Integer
    iIndex = 0
    If Not rstKeywords.EOF Then
        Set cKeywordID = New Collection
        lstKeywords.Clear
        rstKeywords.MoveFirst
        Do While Not rstKeywords.EOF
            lstKeywords.AddItem rstKeywords!KeywordOrCodeSection
            iIndex = rstKeywords!KeywordID
            cKeywordID.Add iIndex ', iIndex
            rstKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If

End Sub

Private Sub BuildRecordset(KeywordID As String)
    Dim rst As ADODB.Recordset
    Dim lngRow As Long
    Dim lngrows As Long
    Dim intcols As Integer
    
    Dim lrecnum As Long
    Dim sRowString As String
    Dim iRecordID As Integer
    Dim icounter As Integer
    Dim sSource As String
    
    Me.lboResults.Clear
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = cnDatabase
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open ("SELECT RecordID, Title from qryRecordsKeywords where KeywordID=" & KeywordID)
        
    End With

    If rst.EOF Then
        Exit Sub
    End If
            
    If Not rst.EOF Then rst.MoveLast
        
    'intcols = 14
    lngrows = rst.RecordCount

    ReDim avardata(0 To lngrows, 0 To 1)

    If Not rst.EOF Then rst.MoveFirst
    lngRow = 0
    
    
    
    Do While Not rst.EOF
        sRowString = ""
        
        'add "Rowstring" to the listbox
        If rst!Title <> "" Then sRowString = rst!Title
        Me.lboResults.AddItem ((lngRow + 1) & ". " & sRowString)
        avardata(lngRow, 0) = rst!recordID
        
        rst.MoveNext
        lngRow = lngRow + 1
        
    Loop

    'End If
    'Set cnx = Nothing
    Set rst = Nothing
    
BuildRecordsetDone:
    On Error GoTo 0
    Exit Sub

BuildRecordsetErr:
    Select Case Err
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
         vbOKOnly + vbCritical, "BuildRecordset"
    End Select
    Resume BuildRecordsetDone
End Sub

Public rstThesaurusFill As ADODB.Recordset
Dim rstThesaurusLookup As ADODB.Recordset
Dim iCurrentRecord As String


Private Sub cmdAddThesaurus_Click()
    frmAddThesaurus.txtThesaurus.Text = ""
    frmAddThesaurus.Caption = "Add Thesaurus Equivalent"
    frmAddThesaurus.Show
End Sub

Private Sub cmdDeleteThesaurus_Click()
    If Me.lstThesaurusEquivalents.SelCount = 0 Then
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
    If Me.lstThesaurusEquivalents.SelCount = 0 Then
        MsgBox "You must select a word to edit."
    Else
        frmAddThesaurus.Caption = "Edit Thesaurus Equivalent"
        frmAddThesaurus.txtThesaurus.Text = Me.lstThesaurusEquivalents.Text
        Call Get_Current_RecNum
        frmAddThesaurus.txtRecNum.Text = iCurrentRecord
        frmAddThesaurus.Show
    End If
'
End Sub


Public Sub fill_list()

'    Dim iIndex As Integer
'    iIndex = 0
    lstThesaurusEquivalents.Clear
    If Not rstThesaurusFill.EOF Then
        rstThesaurusFill.MoveFirst
        Do While Not rstThesaurusFill.EOF
            lstThesaurusEquivalents.AddItem rstThesaurusFill!ThesaurusEquivalent
            'iIndex = rstKeywords!KeywordID
'            cKeywordID.Add iIndex ', iIndex
            rstThesaurusFill.MoveNext
'            'iIndex = iIndex + 1
        Loop
    End If
End Sub


Private Sub Get_Current_RecNum()
    Dim sTE As String
    
    sTE = Me.lstThesaurusEquivalents.Text
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

