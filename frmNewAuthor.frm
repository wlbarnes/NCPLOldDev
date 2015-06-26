VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewAuthor 
   Caption         =   "Add New Author"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAETID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtSuffix 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtInstitutionalEntity 
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
      Caption         =   "Cancel"
      Size            =   "2566;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   480
      Width           =   1455
      Caption         =   "Save"
      Size            =   "2566;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblType 
      Caption         =   "Author, Ed., or Trans."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblSuffix 
      Caption         =   "Suffix"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblMiddleName 
      Caption         =   "Middle Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblFirstName 
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblInstitutionalEntity 
      Caption         =   "Institutional Entity"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmNewAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sFirstName As String
    Dim sMiddleName As String
    Dim sLastName As String
    Dim sInstitutionalEntity As String
    Dim sSuffix As String
    Dim sType As String
    Dim sFirstNameTest As String
    Dim sMiddleNameTest As String
    Dim sLastNameTest As String
    Dim sInstitutionalEntityTest As String
    Dim sSuffixTest As String
    Dim sTypeTest As String
    Dim sItem As String
    Dim sFullName As String
    Dim iAETID As Integer
    Dim rstAuthorTest As Recordset
    Dim iCurrentListItem As Integer
    Dim i As Integer
    
    If ((Me.txtInstitutionalEntity = "") And ((Me.txtFirstName = "") Or (Me.txtLastName = ""))) Or (Me.cmbType = "") Then
        MsgBox "You did not enter all required fields."
        Cancel = True
    Else
    
    Set rstAuthorTest = New ADODB.Recordset
    With rstAuthorTest
        .ActiveConnection = frmMain.cnWriteDatabase
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open ("SELECT * from tblAuthorsEditorsTranslators")
    End With
        sFirstName = Me.txtFirstName.Text
        sMiddleName = Me.txtMiddleName.Text
        sLastName = Me.txtLastName.Text
        sInstitutionalEntity = Me.txtInstitutionalEntity.Text
        sSuffix = Me.txtSuffix.Text
        sType = Me.cmbType.Text
        
        Do Until rstAuthorTest.EOF
            If IsNull(rstAuthorTest!InstitutionalEntity) Then sInstitutionalEntityTest = "" _
                Else sInstitutionalEntityTest = rstAuthorTest!InstitutionalEntity
            
            If IsNull(rstAuthorTest!FirstName) Then sFirstNameTest = "" _
                Else sFirstNameTest = rstAuthorTest!FirstName
            If IsNull(rstAuthorTest!MiddleName) Then sMiddleNameTest = "" _
                Else sMiddleNameTest = rstAuthorTest!MiddleName
            If IsNull(rstAuthorTest!LastName) Then sLastNameTest = "" _
                Else sLastNameTest = rstAuthorTest!LastName
            If IsNull(rstAuthorTest!Suffix) Then sSuffixTest = "" _
                Else sSuffixTest = rstAuthorTest!Suffix
            If IsNull(rstAuthorTest!AETType) Then sTypeTest = "" _
                Else sTypeTest = rstAuthorTest!AETType
            
            If (sInstitutionalEntityTest = sInstitutionalEntity) And _
            (sFirstNameTest = sFirstName) And _
            (sMiddleNameTest = sMiddleName) And _
            (sLastNameTest = sLastName) And _
            (sSuffixTest = sSuffix) And _
            (sTypeTest = sType) Then

                MsgBox "Author Already Exists in Database."
                'Call Clear_Form
                GoTo Duplicate_Record
            End If
            rstAuthorTest.MoveNext
        Loop
        rstAuthorTest.AddNew
            If sInstitutionalEntity <> "" Then rstAuthorTest!InstitutionalEntity = sInstitutionalEntity
            If sFirstName <> "" Then rstAuthorTest!FirstName = sFirstName
            If sMiddleName <> "" Then rstAuthorTest!MiddleName = sMiddleName
            If sLastName <> "" Then rstAuthorTest!LastName = sLastName
            If sSuffix <> "" Then rstAuthorTest!Suffix = sSuffix
            If sType <> "" Then rstAuthorTest!AETType = sType
        rstAuthorTest.Update
        iAETID = rstAuthorTest!AETID
        Select Case sType
            Case "Author"
                'frmMain.rstAuthors.Requery
                'frmMain.rstAuthors.MoveFirst
                'frmMain.rstAuthors.Find ("AETID = " & iAETID)
                sFullName = frmMain.Full_AET_Name(rstAuthorTest)
                'sItem = frmMain.rstAuthors.Fields("FullName").Value & " (ID: " & iAETID & ")"
                sItem = sFullName & " (ID: " & iAETID & ")"
                
                frmMain.lstAuthors.AddItem sItem
                For i = 1 To (frmMain.lstAuthors.ListCount - 1)
                    If sItem = frmMain.lstAuthors.List(i) Then
                        iCurrentListItem = i
                        frmMain.lstAuthors.Selected(i) = True
                        GoTo ExitHere
                    End If
                Next
ExitHere:
                Call frmMain.Manage_Lists(frmMain.lstCurrentAuthors, frmMain.lstAuthors, frmMain.cAuthors)

        
            Case "Editor"
                'frmMain.rstEditors.Requery
                'frmMain.rstEditors.MoveFirst
                'frmMain.rstEditors.Find ("AETID = " & iAETID)
                'sItem = frmMain.rstEditors.Fields("FullName").Value & " (ID: " & frmMain.rstEditors!AETID & ")"
                sFullName = frmMain.Full_AET_Name(rstAuthorTest)
                sItem = sFullName & " (ID: " & iAETID & ")"
                frmMain.lstEditors.AddItem sItem
                For i = 1 To (frmMain.lstEditors.ListCount - 1)
                    If sItem = frmMain.lstEditors.List(i) Then
                        iCurrentListItem = i
                        frmMain.lstEditors.Selected(i) = True
                        GoTo ExitHereEditor
                    End If
                Next
ExitHereEditor:
                Call frmMain.Manage_Lists(frmMain.lstCurrentEditors, frmMain.lstEditors, frmMain.cEditors)

                
            Case "Translator"
                sFullName = frmMain.Full_AET_Name(rstAuthorTest)
                sItem = sFullName & " (ID: " & iAETID & ")"
                frmMain.lstTranslators.AddItem sItem
                For i = 1 To (frmMain.lstTranslators.ListCount - 1)
                    If sItem = frmMain.lstTranslators.List(i) Then
                        iCurrentListItem = i
                        frmMain.lstTranslators.Selected(i) = True
                        GoTo ExitHereTranslator
                    End If
                Next
ExitHereTranslator:
                Call frmMain.Manage_Lists(frmMain.lstCurrentTranslators, frmMain.lstTranslators, frmMain.cTranslators)

            
            
                'frmMain.rstTranslators.Requery
        
        End Select
        
        
        'frmMain.cmbJournalTitle.AddItem sJournalTitle
        'frmMain.cmbJournalTitle.Text = sJournalTitle
        'frmMain.txtJournalID = iJournalID
        'frmMain.txtJournalTitleShortForm.Text = sJournalTitleShortForm
        'frmMain.cmbPagination = sPagination
        'frmMain.txtCallNumber = sCallNumber
        'frmMain.txtPlaceOfPublication = sPlaceOfPublication
        Unload Me
        Call Clear_Form
    rstAuthorTest.Close

    End If
Duplicate_Record:
    
    Set rstAuthorTest = Nothing
    
End Sub

Private Sub Form_Load()
    Me.cmbType.AddItem "Author"
    Me.cmbType.AddItem "Editor"
    Me.cmbType.AddItem "Translator"
    Select Case frmMain.cmdNewAuthor.Caption
        Case "New Author"
            Me.cmbType.Text = "Author"
        Case "New Editor"
            Me.cmbType.Text = "Editor"
        Case "New Translator"
            Me.cmbType.Text = "Translator"
    End Select
End Sub

Private Sub Clear_Form()
        Me.txtFirstName = ""
        Me.txtInstitutionalEntity = ""
        Me.txtLastName = ""
        Me.txtMiddleName = ""
        Me.txtSuffix = ""
        Me.cmbType = ""
End Sub
