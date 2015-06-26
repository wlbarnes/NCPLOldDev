Option Strict Off
Option Explicit On
Friend Class frmNewJournal
	Inherits System.Windows.Forms.Form
	Public bEdit As Boolean
	
	Private Sub cmbPagination_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbPagination.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		If (cmbPagination.Text <> "Consecutive") And (cmbPagination.Text <> "Nonconsecutive") Then
			MsgBox("Not a valid pagination type.")
			Cancel = True
			
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Call Clear_Form()
		Me.Close()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim sJournalTitle As String
		Dim sJournalTitleShortForm As String
		Dim sPagination As String
		Dim sCallNumber As String
		Dim sPlaceOfPublication As String
		Dim iJournalID As Short
		Dim rstJournalCheck As ADODB.Recordset
		Dim sSource As String
		
		If (Me.txtNewJournal.Text = "") Or (Me.txtNewJournalShortForm.Text = "") Or (Me.cmbPagination.Text = "") Then
			MsgBox("You did not enter all required fields.")
			'Cancel = True
		Else
			sJournalTitle = Me.txtNewJournal.Text
			sSource = "SELECT * FROM tblJournals"
			rstJournalCheck = New ADODB.Recordset
			rstJournalCheck.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstJournalCheck.Open(sSource, frmMain.cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			If Me.Text = "New Journal" Then
				rstJournalCheck.MoveFirst()
				Do Until rstJournalCheck.EOF
					If rstJournalCheck.Fields("JournalTitle").Value = sJournalTitle Then
						MsgBox("Journal Already Exists in Database.")
						Call Clear_Form()
						GoTo Duplicate_Record
					End If
					rstJournalCheck.MoveNext()
				Loop 
			End If
			sJournalTitleShortForm = Me.txtNewJournalShortForm.Text
			sPagination = Me.cmbPagination.Text
			sCallNumber = Me.txtCallNumber.Text
			sPlaceOfPublication = Me.txtPlaceOfPublication.Text
			If Me.Text = "Edit Journal" Then
				rstJournalCheck.MoveFirst()
				Do Until rstJournalCheck.Fields("JournalTitle").Value = sJournalTitle
					rstJournalCheck.MoveNext()
				Loop 
			End If
			If Me.Text = "New Journal" Then rstJournalCheck.AddNew()
			If sJournalTitle <> "" Then rstJournalCheck.Fields("JournalTitle").Value = sJournalTitle
			If sJournalTitleShortForm <> "" Then rstJournalCheck.Fields("JournalTitleShortFOrm").Value = sJournalTitleShortForm
			If sPagination <> "" Then rstJournalCheck.Fields("Pagination").Value = sPagination
			If sCallNumber <> "" Then rstJournalCheck.Fields("CallNumber").Value = sCallNumber
			rstJournalCheck.Fields("PlaceOfPublication").Value = sPlaceOfPublication
			rstJournalCheck.Update()
			'If Me.Caption = "New Journal" Then iJournalID = rstJournalCheck!JournalID
			iJournalID = rstJournalCheck.Fields("JournalID").Value
			rstJournalCheck.Requery()
			Call frmMain.Populate_Journal_Combobox()
			'frmMain.cmbJournalTitle.AddItem sJournalTitle
			frmMain.cmbJournalTitle.Text = sJournalTitle
			frmMain.txtJournalID.Text = CStr(iJournalID)
			'frmMain.txtJournalTitleShortForm.Text = sJournalTitleShortForm
			frmMain.cmbPagination.Text = sPagination
			frmMain.txtJournaTitleShortForm.Text = sJournalTitleShortForm
			'frmMain.txtCallNumber = sCallNumber
			'frmMain.txtPlaceOfPublication = sPlaceOfPublication
			Me.Close()
			Call Clear_Form()
			rstJournalCheck.Close()
			'UPGRADE_NOTE: Object rstJournalCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstJournalCheck = Nothing
			
		End If
Duplicate_Record: 
	End Sub
	
	Private Sub frmNewJournal_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.cmbPagination.Items.Add("Consecutive")
		Me.cmbPagination.Items.Add("Nonconsecutive")
		If bEdit Then Me.Text = "Edit Journal" Else Me.Text = "New Journal"
		If Me.Text = "Edit Journal" Then Call Fill_Form()
	End Sub
	
	Private Sub Clear_Form()
		Me.txtCallNumber.Text = ""
		Me.txtNewJournal.Text = ""
		Me.txtNewJournalShortForm.Text = ""
		Me.txtPlaceOfPublication.Text = ""
		Me.cmbPagination.Text = ""
	End Sub
	
	Private Sub Fill_Form()
		Dim sJournalTitle As String
		Dim sJournalTitleShortForm As String
		Dim sPagination As String
		Dim sCallNumber As String
		Dim sPlaceOfPublication As String
		Dim iJournalID As Short
		Dim rstJournalCheck As ADODB.Recordset
		Dim sSource As String
		
		Me.txtJournalID.Text = frmMain.txtJournalID.Text
		If Me.txtJournalID.Text <> "" Then
			iJournalID = CShort(Me.txtJournalID.Text)
		Else
			iJournalID = 0
		End If
		
		
		sSource = "SELECT * FROM tblJournals WHERE JournalID=" & iJournalID
		rstJournalCheck = New ADODB.Recordset
		rstJournalCheck.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstJournalCheck.Open(sSource, frmMain.cnReadDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		'rstJournalCheck.MoveFirst
		'Do Until rstJournalCheck!JournalID = iJournalID
		'    rstJournalCheck.MoveNext
		'Loop
		If rstJournalCheck.Fields("CallNumber").Value <> "" Then Me.txtCallNumber.Text = rstJournalCheck.Fields("CallNumber").Value
		If rstJournalCheck.Fields("JournalTitle").Value <> "" Then Me.txtNewJournal.Text = rstJournalCheck.Fields("JournalTitle").Value
		If rstJournalCheck.Fields("Pagination").Value <> "" Then Me.cmbPagination.Text = rstJournalCheck.Fields("Pagination").Value
		If rstJournalCheck.Fields("JournalTitleShortFOrm").Value <> "" Then Me.txtNewJournalShortForm.Text = rstJournalCheck.Fields("JournalTitleShortFOrm").Value
		If rstJournalCheck.Fields("PlaceOfPublication").Value <> "" Then Me.txtPlaceOfPublication.Text = rstJournalCheck.Fields("PlaceOfPublication").Value
	End Sub
End Class