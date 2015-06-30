VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmKeywordChange 
   Caption         =   "Modify Keywords and Thesaurus Equivalents"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   4800
      TabIndex        =   18
      Top             =   8520
      Width           =   3255
   End
   Begin VB.CommandButton cmdGenThesaurus 
      Caption         =   "Thesaurus Table Entry"
      Height          =   735
      Left            =   7380
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtNumRecords 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteKeyword 
      Caption         =   "Delete Keyword Completely"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ListBox lboResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   360
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   6720
      Width           =   13335
   End
   Begin VB.CommandButton cmdAddNewKeyword 
      Caption         =   "Add New Keyword"
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdEditKeyword 
      Caption         =   "Edit Selected Keyword"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton optDelete 
      Caption         =   "Simply delete keyword; no conversion"
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   5280
      Width           =   3135
   End
   Begin VB.OptionButton optConvert 
      Caption         =   "Convert Deleted Keyword into Thesaurus Equivalent"
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Keyword/Thesaurus Entry"
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtOldKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Conversion"
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Top             =   5880
      Width           =   3255
   End
   Begin VB.TextBox txtKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstThesaurus 
      Height          =   2790
      Left            =   6600
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.ListBox lstKeywords 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   4215
   End
   Begin MSForms.ToggleButton tglMain 
      Height          =   735
      Left            =   360
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      VariousPropertyBits=   746588185
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2990;1296"
      Value           =   "1"
      Caption         =   "Keyword Folding"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblOldKeywordID 
      Caption         =   "Old Keyword ID"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblKeywordID 
      Caption         =   "Keyword ID"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblKeywords 
      Caption         =   "Double Click to remove from list and add to Thesaurus. Then select which keyword should remain and be master keyword"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
End
Attribute VB_Name = "frmKeywordChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstKeywords As ADODB.Recordset
Public rstThesaurus As ADODB.Recordset
Public rstStacks As ADODB.Recordset
Public rstStackJunction As ADODB.Recordset

Public rstKeywordsThesaurus As ADODB.Recordset

Dim rstRecordsKeywords As ADODB.Recordset
Public cnDatabase As ADODB.Connection
Public cnRemoteDatabase As ADODB.Connection
Dim iSelectedKeywordID As Integer
Dim iToRemoveKeywordID As Integer
Dim sSelectedKeyword As String
Dim sToRemoveKeyword As String
Dim iRemovedItemNumber As Integer
Public avardata As Variant

Private Sub cmdAddNewKeyword_Click()
    frmEditKeywords.cmdAdd.Caption = "Add"
    frmEditKeywords.Caption = "Add Keywords"
    'frmEditKeywords.txtKeyword = Me.lstKeywords.List(lstKeywords.ListIndex)
    frmEditKeywords.Show
    frmEditKeywords.txtKeyword.SetFocus
End Sub

Private Sub cmdCheck_Click()
    Dim iSelIndex As Integer
    'Me.Hide
    frmKeywordThesaurusChange.Show
    Call frmKeywordThesaurusChange.Fill_KT_List
    'Call frmKeywordThesaurusChange.Fill_TT_List
    iSelIndex = Me.lstKeywords.ListIndex
    frmKeywordThesaurusChange.lstThesaurusKeywords.ListIndex = iSelIndex
End Sub

Private Sub cmdCombine_Click()
    'Me.Hide
    
    frmCombine.Show
    Call frmCombine.Fill_New_Keywords_List
    Call frmCombine.Fill_Old_Keyword_List
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteKeyword_Click()
    Dim iMsgBoxResult As Integer
    Dim iKeywordNumber As Integer
    iMsgBoxResult = MsgBox("Are you sure? This will delete the keyword and all associations with records will be lost.", vbOKCancel, "Confirm")
    'MsgBox iMsgBoxResult
    If iMsgBoxResult = 1 Then
        iKeywordNumber = Me.txtKeywordID
        On Error GoTo data_Error
        cnDatabase.BeginTrans
            rstKeywords.MoveFirst
            Do While rstKeywords!KeywordID <> iKeywordNumber
                rstKeywords.MoveNext
            Loop
            rstKeywords.Delete
            rstKeywords.Update
        cnDatabase.CommitTrans
        
        Call Me.Fill_Keyword_List
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

Private Sub cmdGenThesaurus_Click()
    'Me.Hide
    frmThesaurusEntry.Show
    Call frmThesaurusEntry.fill_list
End Sub

Private Sub cmdKeywordRebuild_Click()
    frmKeywordRebuild.Show
End Sub

Private Sub cmdStack_Click()
    'Me.Hide
    frmStackEntry.Show
    Call frmStackEntry.Fill_Stack_List
    Call frmStackEntry.Fill_ST_List
    
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
    Dim rstDeleteKeyword As ADODB.Recordset
    
'check to see if all sections made
    If (Me.lstKeywords.SelCount = 0) Or (Me.lstThesaurus.ListCount = 0) Or _
        ((Me.optConvert = False) And (Me.optDelete = False)) Then
        MsgBox "You did not make all necessary selections", vbCritical, "Error"
        'Cancel = True
    Else
    
'put thesaurus equivalents of keyword to be deleted in a collection
        Set rstOldKeywordEquivs = New Recordset
        With rstOldKeywordEquivs
            .CursorLocation = adUseClient
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
        'On Error GoTo data_Error
        
'open a rstRecordsKeywords recordset of to get record numbers that will be affected by deletion
        Set rstRecordsKeywords = New Recordset
        With rstRecordsKeywords
            .CursorLocation = adUseClient
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords where KeywordID=" & iToRemoveKeywordID)
        End With
        
'set up a test to see if any of the converted thesaurus equivalents would duplicate a current thesaurus quivalent of the selected Keyword
        Set rstRKTest = New Recordset
        Set cRecNums = New Collection
        With rstRKTest
            .CursorLocation = adUseClient
            .ActiveConnection = cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords where KeywordID=" & iSelectedKeywordID)
        End With
        
        Do While Not rstRKTest.EOF
            'If rstRecordsKeywords!KeywordID = iSelectedKeywordID Then bNoDuplicate = False
            itmpInt = rstRKTest!RecordID
            cRecNums.Add itmpInt
            rstRKTest.MoveNext
        Loop
        
        Set rstRKTest = Nothing
        bNoDuplicate = True
        'If Not rstRecordsKeywords.EOF Then
            Do While Not rstRecordsKeywords.EOF
                For i = 1 To cRecNums.Count
                    If rstRecordsKeywords!RecordID = cRecNums.Item(i) Then bNoDuplicate = False
                Next
                If bNoDuplicate Then
                    rstRecordsKeywords!KeywordID = iSelectedKeywordID
                    rstRecordsKeywords.Update
                End If
                bNoDuplicate = True
                rstRecordsKeywords.MoveNext
            Loop
            If Not rstRecordsKeywords.EOF Then rstRecordsKeywords.MoveFirst
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
        
'put the thesaurus equivalents of the deleted keyword as thesaurus equivalents of selected keyword
        Set rstKeyWordThesaurusOld = New Recordset
        With rstKeyWordThesaurusOld
            .CursorLocation = adUseClient
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
            'Set rstThesaurus = Nothing
            
            'Set rstKeywordsThesaurus = New Recordset
            rstKeywordsThesaurus.Requery
            rstKeywordsThesaurus.MoveFirst
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
            'Set rstKeywordsThesaurus = Nothing
        End If

    
        'iSelectedKeywordID = 0
        'iToRemoveKeywordID = 0
        'sSelectedKeyword = ""
        'sToRemoveKeyword = ""
        'Me.lstThesaurus.Clear
        'Me.optConvert = False
        'Me.optDelete = False
        cnDatabase.CommitTrans
        
        
        Set rstRecordsKeywords = Nothing

        cnDatabase.BeginTrans
    'delete the keyword itself from the keyword table
            Set rstDeleteKeyword = New ADODB.Recordset
            With rstDeleteKeyword
                .CursorLocation = adUseClient
                .ActiveConnection = cnDatabase
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open ("SELECT * from tblKeywords where KeywordID=" & iToRemoveKeywordID)
            End With
    
            
            'rstKeywords.MoveFirst
            'Do Until rstKeywords!KeywordID = iToRemoveKeywordID
            '    rstKeywords.MoveNext
            'Loop
            rstDeleteKeyword.Delete
            rstDeleteKeyword.Update
            iSelectedKeywordID = 0
        iToRemoveKeywordID = 0
        sSelectedKeyword = ""
        sToRemoveKeyword = ""
        Me.lstThesaurus.Clear
        Me.optConvert = False
        Me.optDelete = False
        
        cnDatabase.CommitTrans
        Set rstDeleteKeyword = Nothing
        'rstKeywords.Requery
        Call Me.Fill_Keyword_List
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
    Dim sremoteConnectionString As String
    Set cnDatabase = New Connection
    Set cnRemoteDatabase = New Connection
    'sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\database\NCPL.mdb"
    'sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLBETA;Data Source=128.122.192.28"
    'sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLBETA;Data Source=NCPL"
    sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLLive;Data Source=NCPL"
    sremoteConnectionString = "Provider=SQLOLEDB.1;Data Source=awssqldev.nyulaw.me;Initial Catalog=NCPLLive;User Id=barnesw;Password=philly"
    
    
    
    cnDatabase.Open (sConnectionString)
    Set rstKeywords = New ADODB.Recordset
    With rstKeywords
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywords")
    End With
    
    Set rstThesaurus = New Recordset
    With rstThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusEquivalent")
    End With
    
    Set rstKeywordsThesaurus = New Recordset
    With rstKeywordsThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywordThesaurus")
    End With
    
    Set rstThesaurus = New Recordset
    With rstThesaurus
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusEquivalent")
    End With
    
    Set rstStacks = New Recordset
    With rstStacks
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusStack")
    End With
        
    Set rstStackJunction = New Recordset
    With rstStackJunction
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblThesaurusStackJunction")
    End With
        
        
    Call Fill_Keyword_List
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rstKeywords = Nothing
    Set rstThesaurus = Nothing
    Set cnDatabase = Nothing
    Set cnRemoteDatabase = Nothing
    Set rstKeywordsThesaurus = Nothing
    Call frmMain.Populate_Keyword_List
End Sub


Private Sub lstKeywords_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    
    
    
    sKeytext = Me.lstKeywords.Text
    sKeytext = Bill_Replace(sKeytext, "'", "''")
    Replace sKeytext, "'", "''"
    Set rstGetKeyNum = New ADODB.Recordset
    With rstGetKeyNum
        .CursorLocation = adUseClient
        .ActiveConnection = cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT KeywordID from tblKeywords where KeywordOrCodeSection='" & sKeytext & "'")
    End With
    
    iKeywordID = rstGetKeyNum!KeywordID
    Set rstGetKeyNum = Nothing
    
    'iItemnumber = lstKeywords.ListIndex
    'iKeywordID = cKeywordID.Item(iItemnumber + 1)
    'sItem = lstKeywords.List(iItemnumber)
    Me.txtKeywordID = iKeywordID
    iSelectedKeywordID = iKeywordID
    Call BuildRecordset(Str(iKeywordID))
    
End Sub

Private Sub lstKeywords_DblClick()
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    If Me.lstThesaurus.ListCount = 0 Then
        iItemnumber = lstKeywords.ListIndex
        'iKeywordID = cKeywordID.Item(iItemnumber + 1)
        iKeywordID = Me.txtKeywordID
        
        sItem = lstKeywords.List(iItemnumber)
        'Me.txtKeywordID = iKeywordID
        lstKeywords.RemoveItem (iItemnumber)
        iRemovedItemNumber = iItemnumber
        'cKeywordID.Remove (iItemnumber + 1)
        lstThesaurus.AddItem (sItem)
        Me.txtOldKeywordID = iKeywordID
        iToRemoveKeywordID = iKeywordID
        sToRemoveKeyword = sItem
        If iItemnumber >= Me.lstKeywords.ListCount Then
            Me.lstKeywords.ListIndex = (Me.lstKeywords.ListCount - 1)
        Else
            Me.lstKeywords.ListIndex = iItemnumber
        End If
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
        'cKeywordID.Add iRemovedID, , , iRemovedItemNumber
    End If
End Sub

Public Sub Fill_Keyword_List()
    'Set cKeywordID = Nothing
    'Set cKeywordID = New Collection
    Dim i As Integer
    Dim sSelectedItem As String
    Dim iSelectedIndex As String
    If Me.lstKeywords.SelCount > 0 Then sSelectedItem = Me.lstKeywords.Text
    rstKeywords.Requery
    
    If Not rstKeywords.EOF Then
        'Set cKeywordID = New Collection
        lstKeywords.Clear
        rstKeywords.MoveFirst
        Do While Not rstKeywords.EOF
            lstKeywords.AddItem rstKeywords!KeywordOrCodeSection
            'iIndex = rstKeywords!KeywordID
            'cKeywordID.Add iIndex ', iIndex
            rstKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If
    iSelectedIndex = 0
    If sSelectedItem <> "" Then
        For i = 0 To (Me.lstKeywords.ListCount - 1)
            If Me.lstKeywords.List(i) = sSelectedItem Then
                iSelectedIndex = i
            End If
        Next
    End If
    Me.lstKeywords.ListIndex = iSelectedIndex
    'Set rstKeywords = Nothing
End Sub

Private Sub BuildRecordset(KeywordID As String)
    'Dim rst As Recordset
' Private Static Sub Timer1_Timer()
    'If bRunProgram = False Then Exit Sub
    
    'Dim ctlSQL As Control
    'Dim cnx As ADODB.Connection
    Dim rst As ADODB.Recordset
    'Dim rstAuthors As ADODB.Recordset
    'Dim rstKeyword As ADODB.Recordset
    'Dim rstEditor As ADODB.Recordset
    
    'Dim iKeywordCount As Integer
    'Dim iKeywordCounter As Integer
    'Dim sKeywordSQL As String
    
    Dim lngRow As Long
    'Dim intCol As Integer
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
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open ("SELECT RecordID, Title from qryRecordsKeywords where KeywordID=" & KeywordID)
        
    End With
    'On Error GoTo BuildRecordsetErr

    If rst.EOF Then
        Me.txtNumRecords.Text = "0 records."
        Exit Sub
    End If
    
    ' Switch to in-line error handling
    'On Error Resume Next

    'If Err <> 0 Then
        'strMsg = "Error " & Err.Number & ": " & Err.Description
        
        ' Extra help for common errors
        'Select Case Err
        'Case adhcErrorActionQuery
        '    strMsg = strMsg & vbCrLf & _
        '    "(This error will occur if the SQL entered is an Action query.)"
        'Case adhcErrorDDLQuery
        '    strMsg = strMsg & vbCrLf & _
        '    "(This error will occur if the SQL entered is a DDL query.)"
        'Case adhcErrorTooFewParameters
        '    strMsg = strMsg & vbCrLf & _
        '    "(This error is often caused by a misspelled table or field name.)"
        'Case Else
            ' Do nothing
        'End Select
        
        'On Error GoTo BuildRecordsetErr

    '    intcols = 0
    '    lngrows = 0
    'Else
    '    On Error GoTo BuildRecordsetErr
        
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
        avardata(lngRow, 0) = rst!RecordID
        
        rst.MoveNext
        lngRow = lngRow + 1
        
    Loop

    'End If
    'Set cnx = Nothing
    Set rst = Nothing
    Me.txtNumRecords.Text = lngRow & " records."
    If lngRow = 1 Then Me.txtNumRecords.Text = "1 record."
    
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

Public Function Bill_Replace(sString As String, sStringToReplace As String, sReplacementString _
    As String) As String
        Dim lReplacePos As Long
        Dim lReplaceLength As Long
        Dim sLeftString As String
        Dim sRightString As String
        lReplacePos = 0
        lReplaceLength = 1
        Do While InStr((lReplacePos + lReplaceLength), sString, sStringToReplace) <> 0
            lReplacePos = InStr((lReplacePos + lReplaceLength), sString, sStringToReplace)
            sLeftString = Left(sString, lReplacePos - 1)
            sRightString = Right(sString, (Len(sString) - (lReplacePos - 1 + Len(sStringToReplace))))
            sString = sLeftString & sReplacementString & sRightString
            lReplaceLength = Len(sReplacementString)
        Loop
        Bill_Replace = sString
End Function


