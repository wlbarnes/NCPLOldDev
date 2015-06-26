VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCombine 
   Caption         =   "Combine Old and New Keywords"
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   10965
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optConvert 
      Caption         =   "Keep existing article association and add new association"
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   5520
      Width           =   3015
   End
   Begin VB.OptionButton optDelete 
      Caption         =   "Remove article association from left column and add to right"
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox txtNewKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtOldKeywordID 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   495
      Left            =   10320
      TabIndex        =   9
      Top             =   10200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFolding 
      Caption         =   "Keyword Folding"
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox lstOldAndNew 
      Height          =   2790
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
   End
   Begin VB.ListBox lstNew 
      Height          =   2790
      Left            =   6840
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Begin"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   6120
      Width           =   3255
   End
   Begin VB.ListBox lboResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   6960
      Width           =   13335
   End
   Begin VB.TextBox txtNumRecords 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Keyword/Thesaurus Entry"
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenThesaurus 
      Caption         =   "Thesaurus Table Entry"
      Height          =   735
      Left            =   7380
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblNewKeywordID 
      Caption         =   "New Keyword ID"
      Height          =   495
      Left            =   6840
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblOldKeywordID 
      Caption         =   "Old Keyword ID"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblNew 
      Caption         =   "Select a keyword in new system to move selected articles"
      Height          =   495
      Left            =   6840
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblOldAndNew 
      Caption         =   "Select to See Articles from Old or New Keyword System"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   4560
      Width           =   2295
   End
   Begin MSForms.ToggleButton tglMain 
      Height          =   735
      Left            =   2700
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      VariousPropertyBits=   746588185
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "2990;1296"
      Value           =   "1"
      Caption         =   "Combine Old/New Keywords"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstOldKeywords As ADODB.Recordset
Public rstOldRecordsKeywords As ADODB.Recordset
Dim cRecordNumbers As Collection


Private Sub cmdCheck_Click()
    Me.Hide
    frmKeywordThesaurusChange.Show
    Call frmKeywordThesaurusChange.Fill_KT_List

End Sub

Private Sub cmdFolding_Click()
    Me.Hide
    frmKeywordChange.Show
    Call frmKeywordChange.Fill_Keyword_List
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    For i = 0 To Me.lboResults.ListCount - 1
        Me.lboResults.Selected(i) = True
    Next
End Sub

Private Sub cmdStart_Click()
    Dim i As Integer
    Dim j As Integer
    Dim iTempRecNum As Integer
    Dim cTempRecNumColl As Collection
    Dim cToRemoveRecNum As Collection
    Dim iRecNumAffected As Integer
    Dim iOldKeywordID As Integer
    Dim iNewKeywordID As Integer
    Dim rstAddRecordsKeywords As ADODB.Recordset
    Dim rstRemoveRecordsKeywords As ADODB.Recordset
    Dim rstTestRecordsKeywords As ADODB.Recordset
    Dim sOpenString As String
'check to see if all sections made
    If (Me.lstNew.SelCount = 0) Or (Me.lstOldAndNew.ListCount = 0) Or _
        ((Me.optConvert = False) And (Me.optDelete = False)) Or _
        (Me.lboResults.SelCount = 0) Then
        MsgBox "You did not make all necessary selections", vbCritical, "Error"
        'Cancel = True
    Else
'process request
        Set cTempRecNumColl = New Collection
        For i = 1 To Me.lboResults.ListCount
            If Me.lboResults.Selected(i - 1) = True Then
                iTempRecNum = cRecordNumbers.Item(i)
                cTempRecNumColl.Add iTempRecNum
            End If
            
        Next
        
'First, look at Records affected in new Keywords to remove duplicates
        'If Me.lblOldKeywordID.Caption = "Old Keyword ID" Then
        iNewKeywordID = Me.txtNewKeywordID.Text
        Set rstTestRecordsKeywords = New Recordset
        With rstTestRecordsKeywords
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords where KeywordID=" & iNewKeywordID)
        End With
'copy collection for possible deletion of records later
        Set cToRemoveRecNum = New Collection
        For i = 1 To cTempRecNumColl.Count
            cToRemoveRecNum.Add cTempRecNumColl.Item(i)
        Next
'take this recordset, see if any duplicates in collection
        Do While Not rstTestRecordsKeywords.EOF
            iRecNumAffected = rstTestRecordsKeywords!RecordID
                j = cTempRecNumColl.Count
                For i = 1 To j
                    If Not (i > cTempRecNumColl.Count) Then
                        If iRecNumAffected = cTempRecNumColl.Item(i) Then
                            cTempRecNumColl.Remove (i)
                            i = i - 1
                            j = j - 1
                        End If
                    End If
                Next
            rstTestRecordsKeywords.MoveNext
        Loop
        Set rstAddRecordsKeywords = New Recordset
        With rstAddRecordsKeywords
            .CursorLocation = adUseClient
            .ActiveConnection = frmKeywordChange.cnDatabase
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open ("SELECT * from tblRecordsKeywords")
        End With
        iNewKeywordID = Me.txtNewKeywordID.Text
        
        frmKeywordChange.cnDatabase.BeginTrans
            On Error GoTo data_Error
            For i = 1 To cTempRecNumColl.Count
                rstAddRecordsKeywords.AddNew
                    rstAddRecordsKeywords!RecordID = cTempRecNumColl.Item(i)
                    rstAddRecordsKeywords!KeywordID = iNewKeywordID
                    
                rstAddRecordsKeywords.Update
            Next
        frmKeywordChange.cnDatabase.CommitTrans
        
        
'now check to see if need to delete
        If Me.optDelete = True Then
            If Me.lblOldKeywordID.Caption = "Old Keyword ID" Then sOpenString = "SELECT * from tblRecordsKeywordsOld"
            If Me.lblOldKeywordID.Caption = "New Keyword ID" Then sOpenString = "SELECT * from tblRecordsKeywords"
            iOldKeywordID = Me.txtOldKeywordID.Text
            
            Set rstRemoveRecordsKeywords = New Recordset
            With rstRemoveRecordsKeywords
                .CursorLocation = adUseClient
                .ActiveConnection = frmKeywordChange.cnDatabase
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open (sOpenString)
            End With
            
            frmKeywordChange.cnDatabase.BeginTrans
                
                Do While Not rstRemoveRecordsKeywords.EOF
                    If rstRemoveRecordsKeywords!KeywordID = iOldKeywordID Then
                        For i = 1 To cToRemoveRecNum.Count
                             If rstRemoveRecordsKeywords!RecordID = cToRemoveRecNum.Item(i) Then
                                rstRemoveRecordsKeywords.Delete
                                rstRemoveRecordsKeywords.Update
                                i = cToRemoveRecNum.Count
                             End If
                        Next
                    End If
                    rstRemoveRecordsKeywords.MoveNext
                Loop
            frmKeywordChange.cnDatabase.CommitTrans
        End If
        
        Set cTempRecNumColl = Nothing
        Set cToRemoveRecNum = Nothing
        Set rstTestRecordsKeywords = Nothing
        Set rstAddRecordsKeywords = Nothing
        Set rstRemoveRecordsKeywords = Nothing
        'End If
        Call Me.Fill_Old_Keyword_List
    End If
data_Error:
    If Err <> 0 Then
        Select Case Err
        Case Else
            MsgBox "Error#" & Err.Number & ": " & Err.Description, _
             vbOKOnly + vbCritical, "BuildRecordset"
             frmKeywordChange.cnDatabase.RollbackTrans
        End Select
    End If
End Sub

Private Sub Form_Load()
    Set rstOldKeywords = New ADODB.Recordset
    With rstOldKeywords
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblKeywordsOld")
    End With
    Call Me.Fill_Old_Keyword_List
    Call Me.Fill_New_Keywords_List
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstOldKeywords = Nothing
    Set rstOldRecordsKeywords = Nothing
End Sub

Public Sub Fill_New_Keywords_List()
    Dim i As Integer
    Dim sSelectedItem As String
    Dim iSelectedIndex As String
    If Me.lstNew.SelCount > 0 Then sSelectedItem = Me.lstNew.Text
    frmKeywordChange.rstKeywords.Requery
    
    If Not frmKeywordChange.rstKeywords.EOF Then
        lstNew.Clear
        frmKeywordChange.rstKeywords.Requery
        frmKeywordChange.rstKeywords.MoveFirst
        Do While Not frmKeywordChange.rstKeywords.EOF
            lstNew.AddItem frmKeywordChange.rstKeywords!KeywordOrCodeSection
            frmKeywordChange.rstKeywords.MoveNext
        Loop
    End If
    iSelectedIndex = 0
    If sSelectedItem <> "" Then
        For i = 0 To (Me.lstNew.ListCount - 1)
            If Me.lstNew.List(i) = sSelectedItem Then
                iSelectedIndex = i
            End If
        Next
    End If
    Me.lstNew.ListIndex = iSelectedIndex
End Sub


Public Sub Fill_Old_Keyword_List()
    'Set cKeywordID = Nothing
    'Set cKeywordID = New Collection
    Dim i As Integer
    Dim sSelectedItem As String
    Dim iSelectedIndex As String
    If Me.lstOldAndNew.SelCount > 0 Then sSelectedItem = Me.lstOldAndNew.Text
    
    frmKeywordChange.rstKeywords.Requery
    Me.lstOldAndNew.Clear
    If Not frmKeywordChange.rstKeywords.EOF Then
        'Set cKeywordID = New Collection
        
        frmKeywordChange.rstKeywords.MoveFirst
        Do While Not frmKeywordChange.rstKeywords.EOF
            Me.lstOldAndNew.AddItem frmKeywordChange.rstKeywords!KeywordOrCodeSection & " (new)"
            'iIndex = frmkeywordchange.rstkeywords!KeywordID
            'cKeywordID.Add iIndex ', iIndex
            frmKeywordChange.rstKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If
    
    rstOldKeywords.Requery
    
    If Not rstOldKeywords.EOF Then
        'Set cKeywordID = New Collection
        rstOldKeywords.MoveFirst
        Do While Not rstOldKeywords.EOF
            Me.lstOldAndNew.AddItem rstOldKeywords!KeywordOrCodeSection & " (old)"
            'iIndex = frmkeywordchange.rstkeywords!KeywordID
            'cKeywordID.Add iIndex ', iIndex
            rstOldKeywords.MoveNext
            'iIndex = iIndex + 1
        Loop
    End If
    
    
    iSelectedIndex = 0
    If sSelectedItem <> "" Then
        For i = 0 To (Me.lstOldAndNew.ListCount - 1)
            If Me.lstOldAndNew.List(i) = sSelectedItem Then
                iSelectedIndex = i
            End If
        Next
    End If
    Me.lstOldAndNew.ListIndex = iSelectedIndex
    'Set rstKeywords = Nothing
End Sub

Private Sub lstNew_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    
    
    
    sKeytext = Me.lstNew.Text
    sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
    Replace sKeytext, "'", "''"
    Set rstGetKeyNum = New ADODB.Recordset
    With rstGetKeyNum
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open ("SELECT KeywordID from tblKeywords where KeywordOrCodeSection='" & sKeytext & "'")
    End With
    
    iKeywordID = rstGetKeyNum!KeywordID
    Set rstGetKeyNum = Nothing
    
    'iItemnumber = lstKeywords.ListIndex
    'iKeywordID = cKeywordID.Item(iItemnumber + 1)
    'sItem = lstKeywords.List(iItemnumber)
    Me.txtNewKeywordID = iKeywordID
End Sub

Private Sub lstOldAndNew_Click()
    Dim rstGetKeyNum As ADODB.Recordset
    Dim iItemnumber As Integer
    Dim iKeywordID As Integer
    Dim sItem As String
    Dim sKeytext As String
    Dim iLeftLen As Integer
    Dim sOldNew As String
    Dim sOpenString As String
    Dim sPassString As String
    
    iLeftLen = Len(Me.lstOldAndNew.Text) - 5
    sOldNew = Right(Me.lstOldAndNew.Text, 5)
    
    sKeytext = Left(Me.lstOldAndNew.Text, iLeftLen)
    sKeytext = frmKeywordChange.Bill_Replace(sKeytext, "'", "''")
    
    If sOldNew = "(new)" Then
        Me.lblOldKeywordID.Caption = "New Keyword ID"
        sOpenString = "SELECT KeywordID from tblKeywords where KeywordOrCodeSection='" & sKeytext & "'"
    End If
    If sOldNew = "(old)" Then
        Me.lblOldKeywordID.Caption = "Old Keyword ID"
        sOpenString = "SELECT KeywordID from tblKeywordsOld where KeywordOrCodeSection='" & sKeytext & "'"
    End If
            
    Set rstGetKeyNum = New ADODB.Recordset
    With rstGetKeyNum
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open sOpenString
    End With
    
    iKeywordID = rstGetKeyNum!KeywordID
    Set rstGetKeyNum = Nothing
    If sOldNew = "(new)" Then
        sPassString = "SELECT RecordID, Title from qryRecordsKeywords where KeywordID=" & iKeywordID
    End If
    If sOldNew = "(old)" Then
        sPassString = "SELECT RecordID, Title from qryRecordsKeywordsOld where KeywordID=" & iKeywordID
    End If
    Me.txtOldKeywordID = iKeywordID
    'iSelectedKeywordID = iKeywordID
    'Call BuildRecordset(Str(iKeywordID))
    Call BuildRecordset(sPassString)

End Sub

Private Sub BuildRecordset(sOpenString As String)
    Dim rst As ADODB.Recordset
    Dim lngRow As Long
    Dim lngrows As Long
    Dim intcols As Integer
    
    Dim lrecnum As Long
    Dim sRowString As String
    Dim iRecordID As Integer
    Dim icounter As Integer
    Dim sSource As String
    Dim iRecNum As Integer
    
    Set cRecordNumbers = New Collection
    Me.lboResults.Clear
    Set rst = New ADODB.Recordset
    With rst
        .CursorLocation = adUseClient
        .ActiveConnection = frmKeywordChange.cnDatabase
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        '.Open ("SELECT RecordID, Title from qryRecordsKeywords where KeywordID=" & KeywordID)
        .Open (sOpenString)
        
    End With
    'On Error GoTo BuildRecordsetErr

    If rst.EOF Then
        Me.txtNumRecords.Text = "0 records."
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
        avardata(lngRow, 0) = rst!RecordID
        iRecNum = rst!RecordID
        cRecordNumbers.Add iRecNum
        
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


