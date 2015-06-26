VERSION 5.00
Begin VB.Form frmNewJournal 
   Caption         =   "New Journal"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtJournalID 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbPagination 
      Height          =   315
      ItemData        =   "frmNewJournal.frx":0000
      Left            =   2760
      List            =   "frmNewJournal.frx":0002
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtPlaceOfPublication 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtCallNumber 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtNewJournalShortForm 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.TextBox txtNewJournal 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblPagination 
      Caption         =   "Pagination"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblPlaceOfPublication 
      Caption         =   "Place of Publication"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblCallNumber 
      Caption         =   "Call Number"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblnewJournalShortForm 
      Caption         =   "Journal Short Form"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Journal Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bEdit As Boolean

Private Sub cmbPagination_Validate(Cancel As Boolean)
    If (cmbPagination.Text <> "Consecutive") And (cmbPagination.Text <> "Nonconsecutive") Then
        MsgBox "Not a valid pagination type."
        Cancel = True
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Call Clear_Form
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sJournalTitle As String
    Dim sJournalTitleShortForm As String
    Dim sPagination As String
    Dim sCallNumber As String
    Dim sPlaceOfPublication As String
    Dim iJournalID As Integer
    Dim rstJournalCheck As ADODB.Recordset
    Dim sSource As String
    
    If (Me.txtNewJournal = "") Or (Me.txtNewJournalShortForm = "") Or (Me.cmbPagination = "") Then
        MsgBox "You did not enter all required fields."
        'Cancel = True
    Else
        sJournalTitle = Me.txtNewJournal.Text
        sSource = "SELECT * FROM tblJournals"
        Set rstJournalCheck = New ADODB.Recordset
        rstJournalCheck.CursorLocation = adUseClient
        rstJournalCheck.Open sSource, frmMain.cnWriteDatabase, adOpenKeyset, adLockOptimistic
        If Me.Caption = "New Journal" Then
            rstJournalCheck.MoveFirst
            Do Until rstJournalCheck.EOF
                If rstJournalCheck!JournalTitle = sJournalTitle Then
                    MsgBox "Journal Already Exists in Database."
                    Call Clear_Form
                    GoTo Duplicate_Record
                End If
                rstJournalCheck.MoveNext
            Loop
        End If
        sJournalTitleShortForm = Me.txtNewJournalShortForm.Text
        sPagination = Me.cmbPagination.Text
        sCallNumber = Me.txtCallNumber.Text
        sPlaceOfPublication = Me.txtPlaceOfPublication.Text
        If Me.Caption = "Edit Journal" Then
            rstJournalCheck.MoveFirst
            Do Until rstJournalCheck!JournalTitle = sJournalTitle
                rstJournalCheck.MoveNext
            Loop
        End If
        If Me.Caption = "New Journal" Then rstJournalCheck.AddNew
            If sJournalTitle <> "" Then rstJournalCheck!JournalTitle = sJournalTitle
            If sJournalTitleShortForm <> "" Then rstJournalCheck!JournalTitleShortFOrm = sJournalTitleShortForm
            If sPagination <> "" Then rstJournalCheck!Pagination = sPagination
            If sCallNumber <> "" Then rstJournalCheck!CallNumber = sCallNumber
            rstJournalCheck!PlaceOfPublication = sPlaceOfPublication
        rstJournalCheck.Update
        'If Me.Caption = "New Journal" Then iJournalID = rstJournalCheck!JournalID
        iJournalID = rstJournalCheck!JournalID
        rstJournalCheck.Requery
        Call frmMain.Populate_Journal_Combobox
        'frmMain.cmbJournalTitle.AddItem sJournalTitle
        frmMain.cmbJournalTitle.Text = sJournalTitle
        frmMain.txtJournalID = iJournalID
        'frmMain.txtJournalTitleShortForm.Text = sJournalTitleShortForm
        frmMain.cmbPagination = sPagination
        frmMain.txtJournaTitleShortForm = sJournalTitleShortForm
        'frmMain.txtCallNumber = sCallNumber
        'frmMain.txtPlaceOfPublication = sPlaceOfPublication
        Unload Me
        Call Clear_Form
        rstJournalCheck.Close
        Set rstJournalCheck = Nothing

    End If
Duplicate_Record:
End Sub

Private Sub Form_Load()
    Me.cmbPagination.AddItem "Consecutive"
    Me.cmbPagination.AddItem "Nonconsecutive"
    If bEdit Then Me.Caption = "Edit Journal" Else Me.Caption = "New Journal"
    If Me.Caption = "Edit Journal" Then Call Fill_Form
End Sub

Private Sub Clear_Form()
        Me.txtCallNumber = ""
        Me.txtNewJournal = ""
        Me.txtNewJournalShortForm = ""
        Me.txtPlaceOfPublication = ""
        Me.cmbPagination = ""
End Sub

Private Sub Fill_Form()
    Dim sJournalTitle As String
    Dim sJournalTitleShortForm As String
    Dim sPagination As String
    Dim sCallNumber As String
    Dim sPlaceOfPublication As String
    Dim iJournalID As Integer
    Dim rstJournalCheck As ADODB.Recordset
    Dim sSource As String
    
    Me.txtJournalID = frmMain.txtJournalID
    If Me.txtJournalID.Text <> "" Then
        iJournalID = Me.txtJournalID.Text
    Else
        iJournalID = 0
    End If
    
    
    sSource = "SELECT * FROM tblJournals WHERE JournalID=" & iJournalID
    Set rstJournalCheck = New ADODB.Recordset
    rstJournalCheck.CursorLocation = adUseClient
    rstJournalCheck.Open sSource, frmMain.cnReadDatabase, adOpenKeyset, adLockOptimistic
    
    'rstJournalCheck.MoveFirst
    'Do Until rstJournalCheck!JournalID = iJournalID
    '    rstJournalCheck.MoveNext
    'Loop
    If rstJournalCheck!CallNumber <> "" Then Me.txtCallNumber.Text = rstJournalCheck!CallNumber
    If rstJournalCheck!JournalTitle <> "" Then Me.txtNewJournal.Text = rstJournalCheck!JournalTitle
    If rstJournalCheck!Pagination <> "" Then Me.cmbPagination.Text = rstJournalCheck!Pagination
    If rstJournalCheck!JournalTitleShortFOrm <> "" Then Me.txtNewJournalShortForm.Text = rstJournalCheck!JournalTitleShortFOrm
    If rstJournalCheck!PlaceOfPublication <> "" Then Me.txtPlaceOfPublication.Text = rstJournalCheck!PlaceOfPublication
End Sub
