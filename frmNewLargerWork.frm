VERSION 5.00
Begin VB.Form frmNewLargerWork 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllChaptersBySameAuthor 
      Caption         =   "All Chapters By Same Author?"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtTitleOfSeriesIfNotIssuedByAuthor 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   3360
      Width           =   7455
   End
   Begin VB.TextBox txtSeriesVolume 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtOriginalPublicationDate 
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtPublisher 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox txtCallNumber 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox lblOriginalPublicationDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Text            =   "Original Publication Date"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox lblSeriesVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Text            =   "Series Volume"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox lblPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Text            =   "Publisher"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox lblTitleOfSeriesIfNotIssuedByAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmNewLargerWork.frx":0000
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox lblCallNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Text            =   "Call Number"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtLargerWorkID 
      Enabled         =   0   'False
      Height          =   495
      Left            =   10560
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLargerWorkTitle 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   7335
   End
   Begin VB.TextBox txtEditionandPrinting 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10560
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLargerWorkName 
      Caption         =   "Larger Work Name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblEditionandPrinting 
      Caption         =   "Edition and Printing"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmNewLargerWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Save_LargerWork(rstLargerWorkCheck As Recordset, cnConnection As Connection, sLargerWork As String, _
        sEditionAndPrinting As String, sPublisher As String, sCallNumber As String, sTitleOfSeriesIfNotIssuedByAuthor As String, _
        sSeriesVolume As String, sOriginalPublicationDate As String, bAllChaptersBySameAuthor As Boolean)
            
    With rstLargerWorkCheck
        .ActiveConnection = cnConnection
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblLargerWorks")
    End With
            
    rstLargerWorkCheck.AddNew
        If sLargerWork <> "" Then rstLargerWorkCheck!LargerWorkTitle = sLargerWork
        If sEditionAndPrinting <> "" Then rstLargerWorkCheck!EditionAndPrinting = sEditionAndPrinting
        If sPublisher <> "" Then rstLargerWorkCheck!Publisher = sPublisher
        If sCallNumber <> "" Then rstLargerWorkCheck!CallNumber = sCallNumber
        If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstLargerWorkCheck!TitleOfSeriesIfNotIssuedByAuthor = sTitleOfSeriesIfNotIssuedByAuthor
        If sSeriesVolume <> "" Then rstLargerWorkCheck!SeriesVolume = sSeriesVolume
        If sOriginalPublicationDate <> "" Then rstLargerWorkCheck!OriginalPublicationDate = sOriginalPublicationDate
        rstLargerWorkCheck!AllChaptersBySameAuthor = bAllChaptersBySameAuthor
    rstLargerWorkCheck.Update
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sLargerWork As String
    Dim sEditionAndPrinting As String
    Dim sPublisher As String
    Dim sCallNumber As String
    Dim sOriginalPublicationDate As String
    Dim sSeriesVolume As String
    Dim sTitleOfSeriesIfNotIssuedByAuthor As String
    Dim bAllChaptersBySameAuthor As Boolean
    Dim iLargerWordID As Integer
    Dim rstLargerWorkCheck As ADODB.Recordset
    Dim sSource As String
    
    If (Me.txtLargerWorkTitle) = "" Then
        MsgBox "You did not enter all required fields."
        Cancel = True
    Else
        Set rstLargerWorkCheck = New ADODB.Recordset
        sSource = "SELECT * FROM tblLargerWorks"
        rstLargerWorkCheck.CursorLocation = adUseClient
        rstLargerWorkCheck.Open sSource, frmMain.cnReadDatabase, adOpenKeyset, adLockOptimistic
        sLargerWork = Me.txtLargerWorkTitle.Text
        rstLargerWorkCheck.MoveFirst
        Do Until rstLargerWorkCheck.EOF
            If rstLargerWorkCheck!LargerWorkTitle = sLargerWork Then
                MsgBox "Larger Work Already Exists in Database."
                Call Clear_Form
                GoTo Duplicate_Record
            End If
            rstLargerWorkCheck.MoveNext
        Loop
                    
        sEditionAndPrinting = Me.txtEditionAndPrinting.Text
        sPublisher = Me.txtPublisher.Text
        sCallNumber = Me.txtCallNumber.Text
        sOriginalPublicationDate = Me.txtOriginalPublicationDate.Text
        sSeriesVolume = Me.txtSeriesVolume.Text
        sTitleOfSeriesIfNotIssuedByAuthor = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
        bAllChaptersBySameAuthor = Me.chkAllChaptersBySameAuthor.Value
        
        Set rstLargerWorkCheck = Nothing
        Set rstLargerWorkCheck = New ADODB.Recordset
        
        
        Call Save_LargerWork(rstLargerWorkCheck, frmMain.cnRemoteWriteDatabase, sLargerWork, _
        sEditionAndPrinting, sPublisher, sCallNumber, sTitleOfSeriesIfNotIssuedByAuthor, _
        sSeriesVolume, sOriginalPublicationDate, bAllChaptersBySameAuthor)
        
        Set rstLargerWorkCheck = Nothing
        Set rstLargerWorkCheck = New ADODB.Recordset
        
        Call Save_LargerWork(rstLargerWorkCheck, frmMain.cnWriteDatabase, sLargerWork, _
        sEditionAndPrinting, sPublisher, sCallNumber, sTitleOfSeriesIfNotIssuedByAuthor, _
        sSeriesVolume, sOriginalPublicationDate, bAllChaptersBySameAuthor)
        
        
        iLargerWorkID = rstLargerWorkCheck!LargerWorkID
        rstLargerWorkCheck.Requery
        
        frmMain.cmbLargerWorkTitle.AddItem sLargerWork
        frmMain.cmbLargerWorkTitle.Text = sLargerWork
        frmMain.txtLargerWorkID = iLargerWorkID
        
        Unload Me
        Call Clear_Form
        rstLargerWorkCheck.Close
        Set rstLargerWorkCheck = Nothing
    End If
Duplicate_Record:
End Sub


Private Sub Clear_Form()
        Me.txtCallNumber.Text = ""
        Me.txtEditionAndPrinting.Text = ""
        Me.txtLargerWorkID = ""
        Me.txtLargerWorkTitle = ""
        Me.txtOriginalPublicationDate = ""
        Me.txtPublisher = ""
        Me.txtSeriesVolume = ""
        Me.txtTitleOfSeriesIfNotIssuedByAuthor = ""
        Me.chkAllChaptersBySameAuthor = 0
End Sub


