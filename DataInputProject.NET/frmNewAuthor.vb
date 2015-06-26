Option Strict Off
Option Explicit On
Friend Class frmNewAuthor
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.ClickEvent
		Me.Close()
	End Sub
	
	Private Sub cmdSave_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.ClickEvent
		Dim Cancel As Object
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
		Dim iAETID As Short
		Dim rstAuthorTest As ADODB.Recordset
		Dim iCurrentListItem As Short
		Dim i As Short
		
		If ((Me.txtInstitutionalEntity.Text = "") And ((Me.txtFirstName.Text = "") Or (Me.txtLastName.Text = ""))) Or (Me.cmbType.Text = "") Then
			MsgBox("You did not enter all required fields.")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cancel = True
		Else
			
			rstAuthorTest = New ADODB.Recordset
			With rstAuthorTest
				.let_ActiveConnection(frmMain.cnWriteDatabase)
				.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
				.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				.LockType = ADODB.LockTypeEnum.adLockOptimistic
				.Open(("SELECT * from tblAuthorsEditorsTranslators"))
			End With
			sFirstName = Me.txtFirstName.Text
			sMiddleName = Me.txtMiddleName.Text
			sLastName = Me.txtLastName.Text
			sInstitutionalEntity = Me.txtInstitutionalEntity.Text
			sSuffix = Me.txtSuffix.Text
			sType = Me.cmbType.Text
			
			Do Until rstAuthorTest.EOF
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("InstitutionalEntity").Value) Then sInstitutionalEntityTest = "" Else sInstitutionalEntityTest = rstAuthorTest.Fields("InstitutionalEntity").Value
				
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("FirstName").Value) Then sFirstNameTest = "" Else sFirstNameTest = rstAuthorTest.Fields("FirstName").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("MiddleName").Value) Then sMiddleNameTest = "" Else sMiddleNameTest = rstAuthorTest.Fields("MiddleName").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("LastName").Value) Then sLastNameTest = "" Else sLastNameTest = rstAuthorTest.Fields("LastName").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("Suffix").Value) Then sSuffixTest = "" Else sSuffixTest = rstAuthorTest.Fields("Suffix").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(rstAuthorTest.Fields("AETType").Value) Then sTypeTest = "" Else sTypeTest = rstAuthorTest.Fields("AETType").Value
				
				If (sInstitutionalEntityTest = sInstitutionalEntity) And (sFirstNameTest = sFirstName) And (sMiddleNameTest = sMiddleName) And (sLastNameTest = sLastName) And (sSuffixTest = sSuffix) And (sTypeTest = sType) Then
					
					MsgBox("Author Already Exists in Database.")
					'Call Clear_Form
					GoTo Duplicate_Record
				End If
				rstAuthorTest.MoveNext()
			Loop 
			rstAuthorTest.AddNew()
			If sInstitutionalEntity <> "" Then rstAuthorTest.Fields("InstitutionalEntity").Value = sInstitutionalEntity
			If sFirstName <> "" Then rstAuthorTest.Fields("FirstName").Value = sFirstName
			If sMiddleName <> "" Then rstAuthorTest.Fields("MiddleName").Value = sMiddleName
			If sLastName <> "" Then rstAuthorTest.Fields("LastName").Value = sLastName
			If sSuffix <> "" Then rstAuthorTest.Fields("Suffix").Value = sSuffix
			If sType <> "" Then rstAuthorTest.Fields("AETType").Value = sType
			rstAuthorTest.Update()
			iAETID = rstAuthorTest.Fields("AETID").Value
			Select Case sType
				Case "Author"
					'frmMain.rstAuthors.Requery
					'frmMain.rstAuthors.MoveFirst
					'frmMain.rstAuthors.Find ("AETID = " & iAETID)
					sFullName = frmMain.Full_AET_Name(rstAuthorTest)
					'sItem = frmMain.rstAuthors.Fields("FullName").Value & " (ID: " & iAETID & ")"
					sItem = sFullName & " (ID: " & iAETID & ")"
					
					frmMain.lstAuthors.Items.Add(sItem)
					For i = 1 To (frmMain.lstAuthors.Items.Count - 1)
						If sItem = VB6.GetItemString(frmMain.lstAuthors, i) Then
							iCurrentListItem = i
							frmMain.lstAuthors.SetSelected(i, True)
							GoTo ExitHere
						End If
					Next 
ExitHere: 
					Call frmMain.Manage_Lists((frmMain.lstCurrentAuthors), (frmMain.lstAuthors), (frmMain.cAuthors))
					
					
				Case "Editor"
					'frmMain.rstEditors.Requery
					'frmMain.rstEditors.MoveFirst
					'frmMain.rstEditors.Find ("AETID = " & iAETID)
					'sItem = frmMain.rstEditors.Fields("FullName").Value & " (ID: " & frmMain.rstEditors!AETID & ")"
					sFullName = frmMain.Full_AET_Name(rstAuthorTest)
					sItem = sFullName & " (ID: " & iAETID & ")"
					frmMain.lstEditors.Items.Add(sItem)
					For i = 1 To (frmMain.lstEditors.Items.Count - 1)
						If sItem = VB6.GetItemString(frmMain.lstEditors, i) Then
							iCurrentListItem = i
							frmMain.lstEditors.SetSelected(i, True)
							GoTo ExitHereEditor
						End If
					Next 
ExitHereEditor: 
					Call frmMain.Manage_Lists((frmMain.lstCurrentEditors), (frmMain.lstEditors), (frmMain.cEditors))
					
					
				Case "Translator"
					sFullName = frmMain.Full_AET_Name(rstAuthorTest)
					sItem = sFullName & " (ID: " & iAETID & ")"
					frmMain.lstTranslators.Items.Add(sItem)
					For i = 1 To (frmMain.lstTranslators.Items.Count - 1)
						If sItem = VB6.GetItemString(frmMain.lstTranslators, i) Then
							iCurrentListItem = i
							frmMain.lstTranslators.SetSelected(i, True)
							GoTo ExitHereTranslator
						End If
					Next 
ExitHereTranslator: 
					Call frmMain.Manage_Lists((frmMain.lstCurrentTranslators), (frmMain.lstTranslators), (frmMain.cTranslators))
					
					
					
					'frmMain.rstTranslators.Requery
					
			End Select
			
			
			'frmMain.cmbJournalTitle.AddItem sJournalTitle
			'frmMain.cmbJournalTitle.Text = sJournalTitle
			'frmMain.txtJournalID = iJournalID
			'frmMain.txtJournalTitleShortForm.Text = sJournalTitleShortForm
			'frmMain.cmbPagination = sPagination
			'frmMain.txtCallNumber = sCallNumber
			'frmMain.txtPlaceOfPublication = sPlaceOfPublication
			Me.Close()
			Call Clear_Form()
			rstAuthorTest.Close()
			
		End If
Duplicate_Record: 
		
		'UPGRADE_NOTE: Object rstAuthorTest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthorTest = Nothing
		
	End Sub
	
	Private Sub frmNewAuthor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.cmbType.Items.Add("Author")
		Me.cmbType.Items.Add("Editor")
		Me.cmbType.Items.Add("Translator")
		Select Case frmMain.cmdNewAuthor.Text
			Case "New Author"
				Me.cmbType.Text = "Author"
			Case "New Editor"
				Me.cmbType.Text = "Editor"
			Case "New Translator"
				Me.cmbType.Text = "Translator"
		End Select
	End Sub
	
	Private Sub Clear_Form()
		Me.txtFirstName.Text = ""
		Me.txtInstitutionalEntity.Text = ""
		Me.txtLastName.Text = ""
		Me.txtMiddleName.Text = ""
		Me.txtSuffix.Text = ""
		Me.cmbType.Text = ""
	End Sub
End Class