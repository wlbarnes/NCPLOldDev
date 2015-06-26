Option Strict Off
Option Explicit On
Friend Class frmKeywordThesaurus
	Inherits System.Windows.Forms.Form
	Dim rstKeywords As ADODB.Recordset
	Dim rstThesaurus As ADODB.Recordset
	Dim cKeywordID As Collection
	
	
	Private Sub frmKeywordThesaurus_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim iIndex As Short
		rstKeywords = New ADODB.Recordset
		'Set rstThesaurus = New ADODB.Recordset
		cKeywordID = New Collection
		With rstKeywords
			.let_ActiveConnection(frmMain.cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from tblKeywords"))
		End With
		iIndex = 0
		If Not rstKeywords.EOF Then
			rstKeywords.MoveFirst()
			Do While Not rstKeywords.EOF
				lstKeywords.Items.Add(rstKeywords.Fields("keywordorcodesection").Value)
				iIndex = rstKeywords.Fields("KeywordID").Value
				cKeywordID.Add(iIndex) ', iIndex
				rstKeywords.MoveNext()
				'iIndex = iIndex + 1
			Loop 
		End If
	End Sub
	
	Private Sub frmKeywordThesaurus_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object rstKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstKeywords = Nothing
		'UPGRADE_NOTE: Object rstThesaurus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstThesaurus = Nothing
		'UPGRADE_NOTE: Object cKeywordID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cKeywordID = Nothing
	End Sub
	
	'UPGRADE_WARNING: Event lstKeywords.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstKeywords_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstKeywords.SelectedIndexChanged
		Dim iItemnumber As Short
		Dim iKeywordID As Short
		Dim sItem As String
		iItemnumber = lstKeywords.SelectedIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object cKeywordID.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iKeywordID = cKeywordID.Item(iItemnumber + 1)
		sItem = VB6.GetItemString(lstKeywords, iItemnumber)
		Me.txtKeywordID.Text = CStr(iKeywordID)
		'UPGRADE_NOTE: Object rstThesaurus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstThesaurus = Nothing
		rstThesaurus = New ADODB.Recordset
		
		With rstThesaurus
			.let_ActiveConnection(frmMain.cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from qryThesaurus WHERE KEYWORDID=" & iKeywordID))
		End With
		'rstKeywords.MoveFirst
		'rstKeywords.Find ("rstKeywords!keywordorcodesection = sItem")
		'sItem = rstKeywords!KeywordID
		'MsgBox rstkeywords!KeywordID lstKeywords.List(itemnumber)
		lstThesaurus.Items.Clear()
		If Not rstThesaurus.EOF Then
			rstThesaurus.MoveFirst()
			Do While Not rstThesaurus.EOF
				lstThesaurus.Items.Add(rstThesaurus.Fields("ThesaurusEquivalent").Value)
				rstThesaurus.MoveNext()
			Loop 
		End If
	End Sub
End Class