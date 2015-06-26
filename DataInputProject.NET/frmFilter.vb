Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmFilter
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClear.Click
		Me.txtQuery.Text = ""
		Me.txtQuery.Focus()
	End Sub
	
	Private Sub cmdConvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConvert.Click
		Me.txtQuery.Text = Build_SQL
	End Sub
	
	Private Sub cmdRecordID_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRecordID.Click
		Dim iREcordNum As Short
		Dim sSQLStatement As String
		iREcordNum = CShort(Me.txtQuery.Text)
		'Bill_Replace Me.txtQuery.Text, "*", "%"
		sSQLStatement = "Select * from qryrecordinfo where RecordID=" & iREcordNum
		Call Execute_SQL(sSQLStatement)
	End Sub
	
	Private Sub cmdSQL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSQL.Click
		Dim sSQLStatement As String
		Dim iNumRecords As Short
		sSQLStatement = Me.txtQuery.Text
		Call Execute_SQL(sSQLStatement)
		
	End Sub
	
	Private Sub Execute_SQL(ByRef sSQLStatement As String)
		Dim rstRecordID As Object
		Dim rstRecordTemp As ADODB.Recordset
		Dim sSource As String
		Dim i As Short
		
		If frmMain.rstRecords.State <> 0 Then frmMain.rstRecords.Close()
		On Error GoTo SQLerror
		rstRecordTemp = New ADODB.Recordset
		
		With rstRecordTemp
			.let_ActiveConnection(frmMain.cnReadDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockReadOnly
			.Open(sSQLStatement) '("SELECT * from tblRecords")
		End With
		i = 1
		sSource = "SELECT * FROM tblRecords WHERE "
		Do While Not rstRecordTemp.EOF
			
			'UPGRADE_WARNING: Couldn't resolve default property of object rstRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSource = sSource & "(RECORDID=" & rstRecordID & rstRecordTemp.Fields("RecordID").Value & ")"
			i = i + 1
			rstRecordTemp.MoveNext()
			If Not (i > rstRecordTemp.RecordCount) Then sSource = sSource & " OR "
		Loop 
		With frmMain.rstRecords
			.let_ActiveConnection(frmMain.cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(sSource) '("SELECT * from tblRecords")
		End With
		'iNumRecords = frmMain.rstRecords.RecordCount
		' MsgBox "Query Executed. " & iNumRecords & " records found."
		Call frmMain.populate_RecordID_List()
		frmMain.cmbRecordNumber.SelectedIndex = 0
		Me.Close()
SQLerror: 
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("Syntax error in search statement. Try again.")
				Exit Sub
		End Select
	End Sub
	
	Private Function Build_SQL() As String
		Dim sSearchRequest As String
		Dim cSearchRequest As Collection
		Dim sSQLStatement As String
		
		
		sSearchRequest = Me.txtQuery.Text
		sSearchRequest = UCase(sSearchRequest)
		Bill_Replace(sSearchRequest, "AU (", "AU(")
		Bill_Replace(sSearchRequest, "TI (", "TI(")
		Bill_Replace(sSearchRequest, "JN (", "JN(")
		Bill_Replace(sSearchRequest, "KW (", "KW(")
		Bill_Replace(sSearchRequest, "DA (", "DA(")
		Bill_Replace(sSearchRequest, "NOT (", "NOT(")
		
		'*********************************Extract words from the values ***************************
		
		cSearchRequest = New Collection
		Call Extract_Words(sSearchRequest, cSearchRequest)
		
		'join the words in the author fields so that they can do a full author search
		Call Join_Authors(cSearchRequest, "Advanced")
		
		
		'***************************************Build the SQL Statement***************************
		
		sSQLStatement = SQL_Statement(cSearchRequest)
		
		Build_SQL = sSQLStatement
		
	End Function
	
	Private Sub cmdStandard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStandard.Click
		
		'*******************************execute query and fill the array with values from recordset
		Call Execute_SQL(Build_SQL)
		
	End Sub
	
	Private Function SQL_Statement(ByRef cSearchRequest As Collection) As String
		Dim sSQLStatement As String
		Dim bTitle As Boolean
		Dim bKeyword As Boolean
		Dim bAuthor As Boolean
		Dim bJournal As Boolean
		Dim bDate As Boolean
		Dim bNoField As Boolean
		Dim bLeftParenCount As Short
		Dim bRightParenCount As Short
		Dim bNot As Boolean
		Dim sLikeString As String
		Dim bLeftNotParenCount As Short
		Dim bRightNotParenCount As Short
		Dim sAndOr As String
		Dim lcounter As Short
		Dim bNotNoParen As Boolean
		
		sSQLStatement = "SELECT * FROM qryRecordInfo"
		sSQLStatement = sSQLStatement & " WHERE "
		bAuthor = False
		bTitle = False
		bKeyword = False
		bJournal = False
		bDate = False
		bNot = False
		bNotNoParen = False
		bLeftParenCount = 0
		bRightParenCount = 0
		For lcounter = 1 To cSearchRequest.Count()
			If bNotNoParen Then
				If (Not bAuthor) And (Not bTitle) And (Not bKeyword) And (Not bJournal) And (Not bDate) Then
					bNot = False
					bNotNoParen = False
				End If
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "AU[" Then
				bAuthor = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "TI[" Then
				bTitle = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "KW[" Then
				bKeyword = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "JN[" Then
				bJournal = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "DA[" Then
				bDate = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "NOT[" Then
				bNot = True
				bLeftNotParenCount = 1
				bRightNotParenCount = 0
				lcounter = lcounter + 1
				'sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "NOT" Then
				bNot = True
				bNotNoParen = True
				lcounter = lcounter + 1
			End If
			
			If (bTitle = False) And (bAuthor = False) And (bDate = False) And (bJournal = False) And (bKeyword = False) Then
				bNoField = True
			End If
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "[" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do Until cSearchRequest.Item(lcounter) <> "["
					sSQLStatement = sSQLStatement & "("
					lcounter = lcounter + 1
					If (bTitle Or bAuthor Or bDate Or bJournal Or bKeyword) Then bLeftParenCount = bLeftParenCount + 1
					If ((bNot) And (bNotNoParen = False)) Then bLeftNotParenCount = bLeftNotParenCount + 1
					
				Loop 
			End If
			
			bNoField = False
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "AU[" Then
				bAuthor = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "TI[" Then
				bTitle = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "KW[" Then
				bKeyword = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "JN[" Then
				bJournal = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "DA[" Then
				bDate = True
				bLeftParenCount = 1
				bRightParenCount = 0
				lcounter = lcounter + 1
				sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "NOT[" Then
				bNot = True
				bLeftNotParenCount = 1
				bRightNotParenCount = 0
				lcounter = lcounter + 1
				'sSQLStatement = sSQLStatement & "("
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "NOT" Then
				bNot = True
				bNotNoParen = True
				lcounter = lcounter + 1
			End If
			
			
			If (bTitle = False) And (bAuthor = False) And (bDate = False) And (bJournal = False) And (bKeyword = False) Then
				bNoField = True
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If cSearchRequest.Item(lcounter) = "[" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Do Until cSearchRequest.Item(lcounter) <> "["
					sSQLStatement = sSQLStatement & "("
					lcounter = lcounter + 1
					If (bTitle Or bAuthor Or bDate Or bJournal Or bKeyword) Then bLeftParenCount = bLeftParenCount + 1
				Loop 
			End If
			
			If bNot Then sLikeString = " NOT LIKE '" Else sLikeString = " LIKE '"
			If bNot Then sAndOr = "AND" Else sAndOr = "OR"
			
			If bKeyword Then
				'         sSQLStatement = sSQLStatement _
				''                      & "(((KeywordOrCodeSection" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (KeywordOrCodeSection" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (KeywordOrCodeSection" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (KeywordOrCodeSection" & sLikeString & cSearchRequest.Item(lcounter) & " %')) " & sAndOr & " " _
				'''                       & "((ThesaurusEquivalent" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (ThesaurusEquivalent" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (ThesaurusEquivalent" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (ThesaurusEquivalent" & sLikeString & cSearchRequest.Item(lcounter) & " %')) " & sAndOr & " " _
				''
				
				'for now, leave title in as part of search; remove when keyword system more fully developed
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "((((AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				'& "((((Title" & sLikeString & cSearchRequest.Item(lcounter) & "') " _
				''& sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " _
				''& sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " _
				''& sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & " %') " & sAndOr _
				'' & " (AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr _
				'& " (AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr _
				'
				' & " (AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " _
				''  & sAndOr & " (AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " _
				''  & sAndOr & " (AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				If bNot Then
					sSQLStatement = sSQLStatement & " OR ((AllKeywords IS NULL)))"
					'& " OR (IsNull(AllKeywords)))"
				Else
					sSQLStatement = sSQLStatement & ")"
				End If
			End If
			
			If bTitle Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "(((Title" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ".%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ",%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ":%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ";%') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & ". %') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & ", %') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & ": %') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & "; %') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				'& "(((Title" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				
				'& "(((Title LIKE '" & cSearchRequest.Item(lcounter) & "') OR (Title LIKE '% " & cSearchRequest.Item(lcounter) & " %') OR (Title LIKE '% " & cSearchRequest.Item(lcounter) & "') OR (Title LIKE '" & cSearchRequest.Item(lcounter) & " %')))"
			End If
			
			If bAuthor Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, cSearchRequest.Item(lcounter), " ") Then 'a space,denoting multiple words for author
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQLStatement = sSQLStatement & "((((AllAuthors" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %')" & sAndOr & "(AllAuthors" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "')))"
					
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sSQLStatement = sSQLStatement & "((((AllAuthorLastNameOnly" & sLikeString & cSearchRequest.Item(lcounter) & " %')" & sAndOr & "(AllAuthorLastNameOnly" & sLikeString & cSearchRequest.Item(lcounter) & "')" & sAndOr & "(AllAuthorLastNameOnly" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %')" & sAndOr & "(AllAuthorLastNameOnly" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "')))"
					
				End If
				
				If bNot Then
					sSQLStatement = sSQLStatement & " OR ((AllAuthors IS NULL)))"
					'& " OR (IsNull(AllAuthors)))"
				Else
					sSQLStatement = sSQLStatement & ")"
				End If
			End If
			
			If bJournal Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "(((JournalTitle" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (JournalTitle" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (JournalTitle" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (JournalTitle" & sLikeString & cSearchRequest.Item(lcounter) & " %'))" & sAndOr & " ((JournalTitleShortForm" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (JournalTitleShortForm" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (JournalTitleShortForm" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				'& "(((JournalTitle LIKE '" & cSearchRequest.Item(lcounter) & "') OR (JournalTitle LIKE '% " & cSearchRequest.Item(lcounter) & " %') OR (JournalTitle LIKE '% " & cSearchRequest.Item(lcounter) & "') OR (JournalTitle LIKE '" & cSearchRequest.Item(lcounter) & " %')))"
			End If
			
			If bDate Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Mid(cSearchRequest.Item(lcounter), 1, 1) = ">" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If cSearchRequest.Item(lcounter) = ">" Then
						lcounter = lcounter + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) >= '" & cSearchRequest.Item(lcounter) & "')"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) >= '" & Mid(cSearchRequest.Item(lcounter), 2, Len(cSearchRequest.Item(lcounter)) - 1) & "')"
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Mid(cSearchRequest.Item(lcounter), 1, 1) = "<" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If cSearchRequest.Item(lcounter) = "<" Then
						lcounter = lcounter + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) <= '" & cSearchRequest.Item(lcounter) & "')"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) <= '" & Mid(cSearchRequest.Item(lcounter), 2, Len(cSearchRequest.Item(lcounter)) - 1) & "')"
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Mid(cSearchRequest.Item(lcounter), 1, 1) = "=" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If cSearchRequest.Item(lcounter) = "=" Then
						lcounter = lcounter + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) = '" & cSearchRequest.Item(lcounter) & "')"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) = '" & Mid(cSearchRequest.Item(lcounter), 2, Len(cSearchRequest.Item(lcounter)) - 1) & "')"
					End If
				End If
				'If sYearAfter <> "" Then
				'    sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) >= '" & sYearAfter & "')"
				'    If sYearBefore <> "" Then sSQLStatement = sSQLStatement & " AND "
				'End If
				
				'If sYearBefore <> "" Then sSQLStatement = sSQLStatement & "((qryRecordInfo.PublicationYear) <= '" & sYearBefore
				
			End If
			If bNoField Then 'if no specific field
				'title part
				sSQLStatement = sSQLStatement & "(" 'opening paren
				
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "(((Title" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ".%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ",%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ":%') " & sAndOr & " (Title" & sLikeString & "% " & cSearchRequest.Item(lcounter) & ";%') " & sAndOr & " (Title" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				
				sSQLStatement = sSQLStatement & " " & sAndOr & " "
				
				'keyword part
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "((((AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (AllKeywords" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "') " & sAndOr & " (AllKeywords" & sLikeString & cSearchRequest.Item(lcounter) & " %')))"
				If bNot Then
					sSQLStatement = sSQLStatement & " OR ((AllKeywords IS NULL)))"
					'& " OR (IsNull(AllKeywords)))"
				Else
					sSQLStatement = sSQLStatement & ")"
				End If
				
				'author part
				
				sSQLStatement = sSQLStatement & " " & sAndOr & " "
				
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sSQLStatement = sSQLStatement & "((((AllAuthors" & sLikeString & "% " & cSearchRequest.Item(lcounter) & " %') " & sAndOr & " (AllAuthors" & sLikeString & "% " & cSearchRequest.Item(lcounter) & "')))"
				
				If bNot Then
					sSQLStatement = sSQLStatement & " OR ((AllAuthors IS NULL)))"
					'& " OR (IsNull(AllAuthors)))"
				Else
					sSQLStatement = sSQLStatement & ")"
				End If
				sSQLStatement = sSQLStatement & ")" 'closing paren for nofield
				
			End If
			
			If Not (lcounter + 1) > cSearchRequest.Count() Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (cSearchRequest.Item(lcounter + 1) = "]") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do Until (lcounter + 1 > cSearchRequest.Count()) Or (cSearchRequest.Item(lcounter + 1) <> "]")
						If (bNotNoParen = False) And (Not bAuthor) And (Not bTitle) And (Not bKeyword) And (Not bJournal) And (Not bDate) Then
							If bNot Then bRightNotParenCount = bRightNotParenCount + 1
							
						End If
						
						If Not ((bNot = True) And (bRightNotParenCount = bLeftNotParenCount)) Then sSQLStatement = sSQLStatement & ")"
						
						'If Not (lcounter + 1) > cSearchRequest.Count Then
						lcounter = lcounter + 1
						If (bTitle Or bAuthor Or bDate Or bJournal Or bKeyword) Then bRightParenCount = bRightParenCount + 1
						If (bRightNotParenCount = bLeftNotParenCount) Then bNot = False
						
						If (bRightParenCount = bLeftParenCount) Then
							bRightParenCount = 0
							bLeftParenCount = 0
							bTitle = False
							bKeyword = False
							bAuthor = False
							bJournal = False
							bDate = False
						End If
						If (lcounter = cSearchRequest.Count()) Then GoTo loop_exit
					Loop 
loop_exit: 
					'lcounter = lcounter + 1
				End If
				
			End If
			
			If Not (lcounter + 1) > cSearchRequest.Count() Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(lcounter + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (cSearchRequest.Item(lcounter + 1) = "OR") Or (cSearchRequest.Item(lcounter + 1) = "AND") Then
					If (bNot = True) And (bNotNoParen = False) And (bAuthor Or bTitle Or bKeyword Or bJournal Or bDate) Then
						sSQLStatement = sSQLStatement & " " & "AND" & " "
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object cSearchRequest.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sSQLStatement = sSQLStatement & " " & cSearchRequest.Item(lcounter + 1) & " "
						
					End If
					
					lcounter = lcounter + 1
				Else
					sSQLStatement = sSQLStatement & " " & "AND" & " "
				End If
				
			End If
			
		Next 
		'sSQLStatement = sSQLStatement & " ORDER BY qryRecordInfo.PublicationYear DESC, " _
		''    & " [InstitutionalEntity], [LastName], [Title]"
		SQL_Statement = sSQLStatement
	End Function
	
	'function to put all words in a collection, and put quotes around the words if they are not connectors
	Public Sub Extract_Words(ByVal pString As String, ByVal pCollection As Collection)
		Dim lSpacePos As Integer
		Dim lNextSpacePos As Integer
		Dim sWorkingString As String
		Dim sCurrentItem As String
		Dim lcounter As Integer
		Dim cSpacePos As Collection
		Dim cTempWordPos As Collection
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim iReplaceCounter As Short
		Dim sTestChar As String
		Dim bMultipleWords As Boolean
		Dim bPartOfPhrase As Boolean
		Dim cTempColl As Collection
		Dim cTempCollForBrack As Collection
		Dim cLBracketPos As Collection
		Dim cRBracketPos As Collection
		Dim sWorkingString2 As String
		Dim cQuotePosColl As Collection
		'    Dim bPhrase As Boolean
		'    Dim sPhraseWord As String
		'    Dim lPhraseWordCount As Long
		'
		'
		'    bPhrase = False
		'    lSpacePos = 0
		'    lNextSpacePos = 0
		'    lPhraseWordCount = 0
		bPartOfPhrase = False
		bMultipleWords = True
		sWorkingString = pString
		cTempColl = New Collection
		cQuotePosColl = New Collection
		cTempCollForBrack = New Collection
		cLBracketPos = New Collection
		cRBracketPos = New Collection
		
		'remove quotations, and change characters in working string
		sWorkingString = UCase(sWorkingString)
		'Bill_Replace sWorkingString, Chr(34), "'"
		Bill_Replace(sWorkingString, "{", "[")
		Bill_Replace(sWorkingString, "}", "]")
		Bill_Replace(sWorkingString, "'", "''")
		Bill_Replace(sWorkingString, "BUT NOT", "AND NOT")
		
		
		'Bill_Replace sWorkingString, ",", ", "
		
		'Bill_Replace sWorkingString, "AU", "AU "
		'Bill_Replace sWorkingString, "au", "AU "
		'Bill_Replace sWorkingString, "aU", "AU "
		'Bill_Replace sWorkingString, "Au", "AU "
		
		sWorkingString2 = sWorkingString
		sWorkingString = ""
		'new procedure for checking parentheses; change to brackets with space to accomodate existing procedure
		'rebuild workingstring character by character
		'For i = 1 To Len(sWorkingString2)
		'    If Mid(sWorkingString2, i, 1) = ")" Then
		'        sWorkingString = sWorkingString + " ]"
		'        i = i + 1
		'    End If
		'    If Mid(sWorkingString2, i, 1) = "(" Then
		'        If Mid(sWorkingString2, i + 2, 1) = ")" Then
		'            sWorkingString = sWorkingString + Mid(sWorkingString2, i, 1)
		'            sWorkingString = sWorkingString + Mid(sWorkingString2, i + 1, 1)
		'            sWorkingString = sWorkingString + Mid(sWorkingString2, i + 2, 1)
		'            i = i + 2
		'        Else
		'            sWorkingString = sWorkingString + "[ "
		'        End If
		'    Else
		'    If i <= Len(sWorkingString2) Then sWorkingString = sWorkingString + Mid(sWorkingString2, i, 1)
		'    End If
		'Next
		'Cut and pasted from advanced search; test to see if works
		For i = 1 To Len(sWorkingString2)
			Select Case (Mid(sWorkingString2, i, 1))
				'If Mid(sWorkingString2, i, 1) = ")" Then
				Case ")"
					sWorkingString = sWorkingString & " ]"
					'i = i + 1
					'End If
					'If Mid(sWorkingString2, i, 1) = "(" Then
				Case "("
					If Mid(sWorkingString2, i + 2, 1) = ")" Then
						sWorkingString = sWorkingString & Mid(sWorkingString2, i, 1)
						sWorkingString = sWorkingString & Mid(sWorkingString2, i + 1, 1)
						sWorkingString = sWorkingString & Mid(sWorkingString2, i + 2, 1)
						i = i + 2
					Else
						sWorkingString = sWorkingString & "[ "
					End If
					
				Case Else
					If i <= Len(sWorkingString2) Then sWorkingString = sWorkingString & Mid(sWorkingString2, i, 1)
			End Select
		Next 
		'Bill_Replace sWorkingString, "[", "[ "
		'Bill_Replace sWorkingString, "]", " ]"
		Bill_Replace(sWorkingString, "*", "%")
		Bill_Replace(sWorkingString, "!", "%")
		Bill_Replace(sWorkingString, "?", "_")
		
		
		
		cSpacePos = New Collection
		cTempWordPos = New Collection
		'Set cConnectorPos = New Collection
		
		lSpacePos = 0
		lNextSpacePos = 0
		Do 
			lSpacePos = InStr(lSpacePos + 1, sWorkingString, " ")
			If lSpacePos <> 0 Then cSpacePos.Add(lSpacePos)
		Loop Until lSpacePos = 0
		
		'look at spacepositions to see if there are double spaces
		'If cSpacePos.Count > 0 Then
		'    For i = 1 To cSpacePos.Count
		'    Next
		'End If
		
		'add all words to a temporary collection
		
		'first word
		
		If cSpacePos.Count() = 0 Then
			bMultipleWords = False
			lSpacePos = Len(sWorkingString) + 1
			cSpacePos.Add(lSpacePos)
		End If
		
		
		sCurrentItem = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object cSpacePos.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCurrentItem = VB.Left(sWorkingString, cSpacePos.Item(1) - 1)
		'If Left(sCurrentItem, 1) = "'" Then sCurrentItem = Left(sCurrentItem, Len(sCurrentItem) - 1)
		cTempColl.Add(sCurrentItem)
		
		'middle words
		If cSpacePos.Count() > 1 Then
			For i = 1 To cSpacePos.Count() - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object cSpacePos.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSpacePos.Item(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object cSpacePos.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sCurrentItem = Mid(sWorkingString, cSpacePos.Item(i) + 1, cSpacePos.Item(i + 1) - cSpacePos.Item(i) - 1)
				cTempColl.Add(sCurrentItem)
				
			Next 
		End If
		
		'last word
		If bMultipleWords Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cSpacePos.Item(cSpacePos.Count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sCurrentItem = VB.Right(sWorkingString, Len(sWorkingString) - cSpacePos.Item(cSpacePos.Count()))
			cTempColl.Add(sCurrentItem)
		End If
		sCurrentItem = ""
		For i = 1 To cTempColl.Count()
			'If Left(cTempColl.Item(i), 1) = "'" Then cQuotePosColl.Add i
			'If Right(cTempColl.Item(i), 1) = "'" Then cQuotePosColl.Add i
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If VB.Left(cTempColl.Item(i), 1) = Chr(34) Then cQuotePosColl.Add(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If VB.Right(cTempColl.Item(i), 1) = Chr(34) Then cQuotePosColl.Add(i)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If VB.Left(cTempColl.Item(i), 1) = "[" Then cLBracketPos.Add(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If VB.Left(cTempColl.Item(i), 1) = "]" Then cRBracketPos.Add(i)
			
		Next 
		
		For k = 1 To cTempColl.Count()
			bPartOfPhrase = False
			
			For i = 1 To cQuotePosColl.Count() - 1 Step 2 'step 2 because of matching quotes; odd quote will be beginning
				'UPGRADE_WARNING: Couldn't resolve default property of object cQuotePosColl.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If cQuotePosColl.Item(i) = k Then
					bPartOfPhrase = True
					'UPGRADE_WARNING: Couldn't resolve default property of object cQuotePosColl.Item(i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object cQuotePosColl.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For j = cQuotePosColl.Item(i) To cQuotePosColl.Item(i + 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If VB.Left(cTempColl.Item(j), 1) = Chr(34) Then
							sCurrentItem = VB.Right(cTempColl.Item(j), Len(cTempColl.Item(j)) - 1)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCurrentItem = sCurrentItem & " " & cTempColl.Item(j)
						End If 'get rid of quotation
						k = k + 1
					Next 
					If VB.Right(sCurrentItem, 1) = Chr(34) Then sCurrentItem = VB.Left(sCurrentItem, Len(sCurrentItem) - 1) 'get rid of quotation
					k = k - 1
				End If
			Next 
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not (bPartOfPhrase) Then sCurrentItem = cTempColl.Item(k)
			'Bill_Replace sWorkingString, "'", "''" ' to take care of O'leary, etc, and journal title short form
			
			pCollection.Add(sCurrentItem)
		Next 
	End Sub
	
	Public Function Bill_Replace(ByRef sString As String, ByRef sStringToReplace As String, ByRef sReplacementString As String) As String
		Dim lReplacePos As Integer
		Dim lReplaceLength As Integer
		Dim sLeftString As String
		Dim sRightString As String
		lReplacePos = 0
		lReplaceLength = 1
		Do While InStr(lReplacePos + lReplaceLength, sString, sStringToReplace) <> 0
			lReplacePos = InStr(lReplacePos + lReplaceLength, sString, sStringToReplace)
			sLeftString = VB.Left(sString, lReplacePos - 1)
			sRightString = VB.Right(sString, Len(sString) - (lReplacePos - 1 + Len(sStringToReplace)))
			sString = sLeftString & sReplacementString & sRightString
			lReplaceLength = Len(sReplacementString)
		Loop 
		Bill_Replace = sString
	End Function
	
	Public Sub Join_Authors(ByRef pAuthorCollection As Collection, ByRef pSearchType As String)
		Dim i As Short
		Dim sCurrentAuthor As String
		Dim cTempColl As Collection
		Dim iLeftParenCount As Short
		Dim iRightParenCount As Short
		
		sCurrentAuthor = ""
		cTempColl = New Collection
		
		Select Case pSearchType
			Case "Guided"
				For i = 1 To pAuthorCollection.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (pAuthorCollection.Item(i) <> "AND") And (pAuthorCollection.Item(i) <> "OR") And (pAuthorCollection.Item(i) <> "NOT") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If sCurrentAuthor = "" Then
							sCurrentAuthor = sCurrentAuthor & pAuthorCollection.Item(i)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCurrentAuthor = sCurrentAuthor & " " & pAuthorCollection.Item(i)
						End If
					Else
						cTempColl.Add(sCurrentAuthor)
						cTempColl.Add(pAuthorCollection.Item(i))
						sCurrentAuthor = ""
					End If
					If i = pAuthorCollection.Count() Then cTempColl.Add(sCurrentAuthor)
				Next 
				
			Case "Advanced"
				iLeftParenCount = 0
				iRightParenCount = 0
				sCurrentAuthor = ""
				For i = 1 To pAuthorCollection.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If pAuthorCollection.Item(i) = "AU[" Then
						cTempColl.Add(pAuthorCollection.Item(i))
						iLeftParenCount = 1
						iRightParenCount = 0
						Do Until (iRightParenCount = iLeftParenCount)
							i = i + 1
							Select Case pAuthorCollection.Item(i)
								Case "AND"
									cTempColl.Add(sCurrentAuthor)
									cTempColl.Add(pAuthorCollection.Item(i))
									sCurrentAuthor = ""
								Case "OR"
									cTempColl.Add(sCurrentAuthor)
									cTempColl.Add(pAuthorCollection.Item(i))
									sCurrentAuthor = ""
									
								Case "NOT["
									'cTempColl.Add sCurrentAuthor
									cTempColl.Add(pAuthorCollection.Item(i))
									sCurrentAuthor = ""
									
								Case "["
									iLeftParenCount = iLeftParenCount + 1
									cTempColl.Add(pAuthorCollection.Item(i))
								Case "]"
									iRightParenCount = iRightParenCount + 1
									If iLeftParenCount <> iRightParenCount Then cTempColl.Add(pAuthorCollection.Item(i))
								Case Else
									'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If sCurrentAuthor = "" Then
										sCurrentAuthor = sCurrentAuthor & pAuthorCollection.Item(i)
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object pAuthorCollection.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										sCurrentAuthor = sCurrentAuthor & " " & pAuthorCollection.Item(i)
									End If
									
							End Select
						Loop 
						cTempColl.Add(sCurrentAuthor)
						cTempColl.Add(pAuthorCollection.Item(i))
						sCurrentAuthor = ""
					Else
						cTempColl.Add(pAuthorCollection.Item(i))
					End If
				Next 
		End Select
		pAuthorCollection = cTempColl
	End Sub
End Class