Option Strict Off
Option Explicit On
Friend Class frmNewLargerWork
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim iLargerWorkID As Object
		Dim Cancel As Object
		Dim sLargerWork As String
		Dim sEditionAndPrinting As String
		Dim sPublisher As String
		Dim sCallNumber As String
		Dim sOriginalPublicationDate As String
		Dim sSeriesVolume As String
		Dim sTitleOfSeriesIfNotIssuedByAuthor As String
		Dim bAllChaptersBySameAuthor As Boolean
		Dim iLargerWordID As Short
		Dim rstLargerWorkCheck As ADODB.Recordset
		Dim sSource As String
		
		If (Me.txtLargerWorkTitle).Text = "" Then
			MsgBox("You did not enter all required fields.")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cancel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cancel = True
		Else
			rstLargerWorkCheck = New ADODB.Recordset
			sSource = "SELECT * FROM tblLargerWorks"
			rstLargerWorkCheck.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstLargerWorkCheck.Open(sSource, frmMain.cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			sLargerWork = Me.txtLargerWorkTitle.Text
			rstLargerWorkCheck.MoveFirst()
			Do Until rstLargerWorkCheck.EOF
				If rstLargerWorkCheck.Fields("LargerWorkTitle").Value = sLargerWork Then
					MsgBox("Larger Work Already Exists in Database.")
					Call Clear_Form()
					GoTo Duplicate_Record
				End If
				rstLargerWorkCheck.MoveNext()
			Loop 
			sEditionAndPrinting = Me.txtEditionAndPrinting.Text
			sPublisher = Me.txtPublisher.Text
			sCallNumber = Me.txtCallNumber.Text
			sOriginalPublicationDate = Me.txtOriginalPublicationDate.Text
			sSeriesVolume = Me.txtSeriesVolume.Text
			sTitleOfSeriesIfNotIssuedByAuthor = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
			bAllChaptersBySameAuthor = Me.chkAllChaptersBySameAuthor.CheckState
			rstLargerWorkCheck.AddNew()
			If sLargerWork <> "" Then rstLargerWorkCheck.Fields("LargerWorkTitle").Value = sLargerWork
			If sEditionAndPrinting <> "" Then rstLargerWorkCheck.Fields("EditionAndPrinting").Value = sEditionAndPrinting
			If sPublisher <> "" Then rstLargerWorkCheck.Fields("Publisher").Value = sPublisher
			If sCallNumber <> "" Then rstLargerWorkCheck.Fields("CallNumber").Value = sCallNumber
			If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstLargerWorkCheck.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value = sTitleOfSeriesIfNotIssuedByAuthor
			If sSeriesVolume <> "" Then rstLargerWorkCheck.Fields("SeriesVolume").Value = sSeriesVolume
			If sOriginalPublicationDate <> "" Then rstLargerWorkCheck.Fields("OriginalPublicationDate").Value = sOriginalPublicationDate
			rstLargerWorkCheck.Fields("AllChaptersBySameAuthor").Value = bAllChaptersBySameAuthor
			rstLargerWorkCheck.Update()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object iLargerWorkID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iLargerWorkID = rstLargerWorkCheck.Fields("LargerWorkID").Value
			rstLargerWorkCheck.Requery()
			
			frmMain.cmbLargerWorkTitle.Items.Add(sLargerWork)
			frmMain.cmbLargerWorkTitle.Text = sLargerWork
			'UPGRADE_WARNING: Couldn't resolve default property of object iLargerWorkID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmMain.txtLargerWorkID.Text = iLargerWorkID
			'frmMain.txtJournalTitleShortForm.Text = sJournalTitleShortForm
			'frmMain.cmbPagination = sPagination
			'frmMain.txtCallNumber = sCallNumber
			'frmMain.txtPlaceOfPublication = sPlaceOfPublication
			Me.Close()
			Call Clear_Form()
			rstLargerWorkCheck.Close()
			'UPGRADE_NOTE: Object rstLargerWorkCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstLargerWorkCheck = Nothing
		End If
Duplicate_Record: 
	End Sub
	
	
	Private Sub Clear_Form()
		Me.txtCallNumber.Text = ""
		Me.txtEditionAndPrinting.Text = ""
		Me.txtLargerWorkID.Text = ""
		Me.txtLargerWorkTitle.Text = ""
		Me.txtOriginalPublicationDate.Text = ""
		Me.txtPublisher.Text = ""
		Me.txtSeriesVolume.Text = ""
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = ""
		Me.chkAllChaptersBySameAuthor.CheckState = System.Windows.Forms.CheckState.Unchecked
	End Sub
End Class