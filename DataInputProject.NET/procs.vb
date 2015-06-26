Option Strict Off
Option Explicit On
Friend Class procs
	Public smessage As String
	'Dim wDocument As Word.Application
	
	Public Function dll_msgbx() As Object
		MsgBox(smessage)
	End Function
	
	'Public Sub Process_Word_Line(avardata As Variant, lcounter As Long, bChkRecNum As Boolean, _
	'pDocument As Word.Application)
	
	Public Sub Process_Word_Line(ByRef iSourceType As Short, ByRef sAuthor As String, ByRef iRecordID As Short, ByRef sArticleDesignation As String, ByRef sTitle As String, ByRef sVolume As String, ByRef sJournalTitle As String, ByRef sPage As String, ByRef sMonth As String, ByRef sDay As String, ByRef sYear As String, ByRef sSeriesTitle As String, ByRef sEditor As String, ByRef sEdition As String, ByRef iEditorCount As Short, ByRef bChkRecNum As Boolean, ByRef pDocument As Word.Application, ByRef sLegislativeMaterialType As String, ByRef sNameOfHouse As String, ByRef sNumberOfCongress As String, ByRef SessionOfCongress As String, ByRef sStateLegislativeSession As String, ByRef sUSCCANCitation As String, ByRef sReportOrDocumentNumber As String, ByRef sSuDocNumber As String, ByRef sLocation As String)
		Dim rRange As Word.Range
		
		
		
		'Set wDocument = pDocument
		If bChkRecNum Then pDocument.Selection.TypeText(" (RecNum: " & Str(iRecordID) & ") ")
		If iSourceType <> 3 Then pDocument.Selection.Font.SmallCaps = False Else pDocument.Selection.Font.SmallCaps = True
		pDocument.Selection.Font.Italic = False
		pDocument.Selection.Font.Bold = False
		pDocument.Selection.Font.SmallCaps = False
		'
		Select Case iSourceType
			Case 1 'consecutively paginated
				If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
				If (sArticleDesignation <> "") And (sArticleDesignation <> "book review") And (sArticleDesignation <> "Book Review") Then pDocument.Selection.TypeText(sArticleDesignation & ", ")
				
				If sTitle <> "" Then
					pDocument.Selection.Font.Italic = True
					pDocument.Selection.TypeText(sTitle)
					pDocument.Selection.Font.Italic = False
					pDocument.Selection.TypeText(", ")
				End If
				pDocument.Selection.TypeText(sVolume)
				If sVolume <> "" Then pDocument.Selection.TypeText(" ")
				pDocument.Selection.Font.SmallCaps = True
				pDocument.Selection.TypeText(sJournalTitle)
				'MsgBox sJournalTitle
				pDocument.Selection.Font.SmallCaps = False
				If sPage <> "" Then pDocument.Selection.TypeText(" " & sPage)
				If sYear <> sVolume Then pDocument.Selection.TypeText(" (" & sYear & ")")
				If (sArticleDesignation = "book review") Or (sArticleDesignation = "Book Review") Then pDocument.Selection.TypeText(" (book review).") Else  : pDocument.Selection.TypeText(".")
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
				
				
			Case 2 'nonconsecutively paginated
				If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
				If (sArticleDesignation <> "") And (sArticleDesignation <> "book review") And (sArticleDesignation <> "Book Review") Then pDocument.Selection.TypeText(sArticleDesignation & ", ")
				
				If sTitle <> "" Then
					pDocument.Selection.Font.Italic = True
					pDocument.Selection.TypeText(sTitle)
					pDocument.Selection.Font.Italic = False
					pDocument.Selection.TypeText(", ")
				End If
				
				pDocument.Selection.Font.SmallCaps = True
				pDocument.Selection.TypeText(sJournalTitle & ", ")
				pDocument.Selection.Font.SmallCaps = False
				pDocument.Selection.TypeText(sMonth & " ")
				If sDay <> "" Then pDocument.Selection.TypeText(sDay & ", ")
				pDocument.Selection.TypeText(sYear)
				If sPage <> "" Then pDocument.Selection.TypeText(", at " & sPage)
				If (sArticleDesignation = "book review") Or (sArticleDesignation = "Book Review") Then pDocument.Selection.TypeText(" (book review).") Else  : pDocument.Selection.TypeText(".")
				
				'pDocument.Selection.TypeText "."
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
				
				
				
			Case 3 'treatise
				pDocument.Selection.Font.SmallCaps = True
				If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
				
				If sTitle <> "" Then
					pDocument.Selection.Font.SmallCaps = True
					pDocument.Selection.TypeText(sTitle)
					pDocument.Selection.Font.SmallCaps = False
					pDocument.Selection.TypeText(" ")
				End If
				
				If sEdition <> "" Then sYear = sEdition & " " & sYear
				If sEditor <> "" Then
					
					If iEditorCount = 1 Then sYear = sEditor & " ed., " & sYear
					If iEditorCount > 1 Then sYear = sEditor & " eds., " & sYear
				End If
				
				'sRowString = sRowString & "("
				'     If sSeriesTitle <> "" Then sRowString = sRowString & sSeriesTitle & ", "
				'     If sYear <> "" Then sRowString = sRowString & sYear
				'sRowString = sRowString & ")."
				'pdocument.Selection.TypeText "(" & sYear & ")."
				pDocument.Selection.TypeText("(")
				If sSeriesTitle <> "" Then pDocument.Selection.TypeText(sSeriesTitle & ", ")
				If sYear <> "" Then pDocument.Selection.TypeText(sYear)
				pDocument.Selection.TypeText(").")
				
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
				
			Case 4 'chapter in treatise
				
				If sJournalTitle <> "" Then 'if there is a title for the larger work
					If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
					If sTitle <> "" Then
						pDocument.Selection.Font.Italic = True
						pDocument.Selection.TypeText(sTitle)
						pDocument.Selection.Font.Italic = False
						pDocument.Selection.TypeText(", ")
						pDocument.Selection.Font.Italic = True
						pDocument.Selection.TypeText("in ")
						pDocument.Selection.Font.Italic = False
					End If
					If sSeriesTitle = "" And sVolume <> "" Then
						pDocument.Selection.TypeText(sVolume & " ")
					End If
					pDocument.Selection.Font.SmallCaps = True
					pDocument.Selection.TypeText(sJournalTitle)
					pDocument.Selection.Font.SmallCaps = False
					pDocument.Selection.TypeText(" ")
					If sPage <> "" Then pDocument.Selection.TypeText(sPage & " ")
					pDocument.Selection.TypeText("(")
					If sSeriesTitle <> "" Then
						'If sJournalTitle <> "" Then
						
						pDocument.Selection.TypeText(sSeriesTitle)
						If sVolume <> "" Then pDocument.Selection.TypeText(" No. " & sVolume)
						pDocument.Selection.TypeText(", ")
					End If
					If sEdition <> "" Then sYear = sEdition & " " & sYear
					If sEditor <> "" Then
						
						If iEditorCount = 1 Then sYear = sEditor & " ed., " & sYear
						If iEditorCount > 1 Then sYear = sEditor & " eds., " & sYear
					End If
					
					If sYear <> "" Then pDocument.Selection.TypeText(sYear)
					pDocument.Selection.TypeText(").")
				Else
					pDocument.Selection.Font.SmallCaps = True
					If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
					'If sSeriesTitle <> "" Then pDocument.Selection.TypeText sSeriesTitle & " No. " & sVolume & ", "
					If sSeriesTitle <> "" Then
						pDocument.Selection.TypeText(sSeriesTitle)
						If sVolume <> "" Then pDocument.Selection.TypeText(" No. " & sVolume)
						pDocument.Selection.TypeText(", ")
					End If
					
					If sTitle <> "" Then pDocument.Selection.TypeText(sTitle)
					pDocument.Selection.Font.SmallCaps = False
					
					pDocument.Selection.TypeText(" ")
					
					If sPage <> "" Then pDocument.Selection.TypeText(sPage & " ")
					If sEdition <> "" Then sYear = sEdition & " " & sYear
					If sEditor <> "" Then
						
						If iEditorCount = 1 Then sYear = sEditor & " ed., " & sYear
						If iEditorCount > 1 Then sYear = sEditor & " eds., " & sYear
					End If
					
					If sYear <> "" Then pDocument.Selection.TypeText("(" & sYear & ")") Else 
					pDocument.Selection.TypeText(".")
				End If
				
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
				
				
				'
				
			Case 5 'legislative hearing
				rRange = pDocument.ActiveDocument.Range(Start:=0, End:=pDocument.ActiveDocument.Characters.Count)
				
				If SessionOfCongress <> "" Then sYear = SessionOfCongress & " " & sYear
				Select Case sLegislativeMaterialType
					
					Case "Committee Hearing"
						If sAuthor <> "" Then rRange.InsertAfter(sAuthor & ", ")
						
						If sTitle <> "" Then rRange.InsertAfter("#ITALICS " & sTitle & " ITALICS#" & ", ")
						If sNumberOfCongress <> "" Then rRange.InsertAfter(sNumberOfCongress & " ")
						If sPage <> "" Then rRange.InsertAfter(sPage & " ")
						If sYear <> "" Then rRange.InsertAfter("(" & sYear & ")")
						rRange.InsertAfter(".")
						
					Case "Report", "Executive Document", "Miscellaneous Document", "Conference Report"
						rRange.InsertAfter("#SMALLCAPS ")
						If sAuthor <> "" Then rRange.InsertAfter(sAuthor & ", ")
						rRange.InsertAfter(sTitle & " ")
						
						If sReportOrDocumentNumber <> "" Then
							If sLegislativeMaterialType = "Report" Then
								If sNameOfHouse = "Senate" Then rRange.InsertAfter("S. Rep. No. ")
								If sNameOfHouse = "House" Then rRange.InsertAfter("H.R. Rep. No. ")
								If sNameOfHouse = "" Then rRange.InsertAfter("Rep. No. ")
							End If
							If sLegislativeMaterialType = "Conference Report" Then
								If sNameOfHouse = "Senate" Then rRange.InsertAfter("S. Conf. Rep. No. ")
								If sNameOfHouse = "House" Then rRange.InsertAfter("H.R. Conf. Rep. No. ")
								If sNameOfHouse = "" Then rRange.InsertAfter("Conf. Rep. No. ")
							End If
							If sLegislativeMaterialType = "Executive Document" Then
								If sNameOfHouse = "Senate" Then rRange.InsertAfter("S. Exec. Doc. No. ")
								If sNameOfHouse = "House" Then rRange.InsertAfter("H.R. Exec. Doc. No. ")
								If sNameOfHouse = "" Then rRange.InsertAfter("Exec. Doc. No. ")
							End If
							If sLegislativeMaterialType = "Miscellaneous Document" Then
								If sNameOfHouse = "Senate" Then rRange.InsertAfter("S. Misc. Doc. No. ")
								If sNameOfHouse = "House" Then rRange.InsertAfter("H.R. Misc. Doc. No. ")
								If sNameOfHouse = "" Then rRange.InsertAfter("Misc. Doc. No. ")
							End If
							rRange.InsertAfter(sReportOrDocumentNumber)
							
						End If
						rRange.InsertAfter(" SMALLCAPS#")
						If sPage <> "" Then rRange.InsertAfter(", at " & sPage & " ")
						If sYear <> "" Then rRange.InsertAfter("(" & sYear & ")")
						If sUSCCANCitation <> "" Then rRange.InsertAfter(", #ITALICS reprinted in ITALICS# " & sUSCCANCitation)
						rRange.InsertAfter(".")
						
					Case "Committee Print"
						sYear = "Comm. Print " & sYear
						rRange.InsertAfter("#SMALLCAPS ")
						If sAuthor <> "" Then rRange.InsertAfter(sAuthor & ", ")
						If sNumberOfCongress <> "" Then rRange.InsertAfter(sNumberOfCongress & ", ")
						rRange.InsertAfter(sTitle & " SMALLCAPS#")
						
						
						If sPage <> "" Then rRange.InsertAfter(sPage)
						rRange.InsertAfter(" (" & sYear & ").")
					Case Else
						
						'others can be added later
						
						
						'Case "State Material"
						
				End Select
				pDocument.Selection.Find.ClearFormatting()
				pDocument.Selection.Find.Replacement.ClearFormatting()
				pDocument.Selection.Find.Replacement.Font.SmallCaps = True
				With pDocument.Selection.Find
					.Text = "(#SMALLCAPS) (*) (SMALLCAPS#)"
					.Replacement.Text = "\2"
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = True
					.MatchCase = False
					.MatchWholeWord = False
					.MatchAllWordForms = False
					.MatchSoundsLike = False
					.MatchWildcards = True
				End With
				pDocument.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
				
				
				pDocument.Selection.Find.ClearFormatting()
				pDocument.Selection.Find.Replacement.ClearFormatting()
				pDocument.Selection.Find.Replacement.Font.Italic = True
				With pDocument.Selection.Find
					.Text = "(#ITALICS) (*) (ITALICS#)"
					.Replacement.Text = "\2"
					.Forward = True
					.Wrap = Word.WdFindWrap.wdFindContinue
					.Format = True
					.MatchCase = False
					.MatchWholeWord = False
					.MatchAllWordForms = False
					.MatchSoundsLike = False
					.MatchWildcards = True
				End With
				pDocument.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
				
				
			Case 6 'nonprint material
				If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
				If sTitle <> "" Then
					pDocument.Selection.Font.Italic = True
					pDocument.Selection.TypeText(sTitle)
					pDocument.Selection.Font.Italic = False
					'pDocument.Selection.TypeText ", "
					pDocument.Selection.TypeText((" ("))
				End If
				If sMonth <> "" Then pDocument.Selection.TypeText(sMonth & " ")
				If sDay <> "" Then pDocument.Selection.TypeText(sDay & ", ")
				pDocument.Selection.TypeText(sYear & ")")
				
				'have put journaltitle here, instead of making avardata bigger, because these won't have a journaltitle field
				
				If sJournalTitle = "" Then
					pDocument.Selection.TypeText(".")
				Else
					pDocument.Selection.TypeText(", ")
					pDocument.Selection.Font.Italic = True
					pDocument.Selection.TypeText("available at ")
					pDocument.Selection.Font.Italic = False
					'have put journaltitle here, instead of making avardata bigger, because these won't have a journaltitle field
					pDocument.Selection.TypeText((sJournalTitle) & ".")
				End If
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
			Case 7 'unpublished work
				If sAuthor <> "" Then pDocument.Selection.TypeText(sAuthor & ", ")
				
				pDocument.Selection.TypeText(sTitle & " (")
				
				If sMonth <> "" Then pDocument.Selection.TypeText(sMonth & " ")
				If sDay <> "" Then pDocument.Selection.TypeText(sDay & ", ")
				pDocument.Selection.TypeText(sYear & ")")
				'pdocument.Selection.TypeText " (unpublished work, on file with author)."
				If iSourceType = 7 Then
					pDocument.Selection.TypeText(" (unpublished work).")
					'add ability for location later
				End If
				pDocument.Selection.TypeParagraph()
				pDocument.Selection.TypeParagraph()
				
				
			Case Else
				' Do nothing
		End Select
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'Set wDocument = New Word.Application
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error Resume Next
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function Word_Report(ByRef pReportType As String, ByRef iRecords As Short, ByRef avardata As Object, ByRef bChkRecNum As Boolean, ByRef sSessionID As String) As String
		Dim ACRODISTXLib As Object
		
		Dim wDocument As Word.Application
		Dim vFileFilter As Object
		Dim sFileName As String
		Dim sPDFFileName As String
        'Dim PDFMaker As ACRODISTXLib
		Dim dDate As Date
		Dim sDate As String
		'Dim sctest As wcPrepareReport
		Dim lcounter As Integer
		'Dim bChkRecNum As Boolean
		Dim i As Short
		Dim iSourceType As Short
		Dim sTitle As String
		Dim sVolume As String
		Dim sPage As String
		Dim sMonth As String
		Dim sYear As String
		Dim sJournalTitle As String
		Dim sDay As String
		'Dim iRecords As Integer
		Dim sCongress As String
		Dim sNotes As String
		Dim sAuthor As String
		Dim sArticleDesignation As String
		Dim iRecordID As Short
		Dim sSeriesTitle As String
		Dim sEditor As String
		Dim sEdition As String
		Dim iEditorCount As Short
		Dim sLocation As String
		wDocument = New Word.Application
		wDocument.Visible = False
		wDocument.Documents.Add() ', , , False
		
		' Switch to in-line error handling
		On Error Resume Next
		
		wDocument.Selection.Font.Size = 12
		wDocument.Selection.Font.Name = "Times New Roman"
		'Set sctest = New wcPrepareReport
		'tmpPrepare.WriteTemplate
		'iRecords = lngRow
		
		bChkRecNum = False 'if you want recnum to be displayed before each record in report
		
		
		For lcounter = 0 To iRecords
			'If ((lcounter And 3) = 3) Then 'Response.Write ". "
			sTitle = ""
			sJournalTitle = ""
			sVolume = ""
			sMonth = ""
			sYear = ""
			sPage = ""
			sDay = ""
			sCongress = ""
			sNotes = ""
			iSourceType = 0
			sAuthor = ""
			iRecordID = 0
			sLocation = ""
			sSeriesTitle = ""
			
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iSourceType = avardata(lcounter, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sAuthor = avardata(lcounter, 2)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iRecordID = avardata(lcounter, 3)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sArticleDesignation = avardata(lcounter, 4)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sYear = avardata(lcounter, 11)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sJournalTitle = Trim(avardata(lcounter, 7))
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sPage = avardata(lcounter, 8)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sMonth = avardata(lcounter, 9)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sDay = avardata(lcounter, 10)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTitle = avardata(lcounter, 5)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sVolume = avardata(lcounter, 6)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSeriesTitle = avardata(lcounter, 14)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sEditor = avardata(lcounter, 12)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sEdition = avardata(lcounter, 13)
			'UPGRADE_WARNING: Couldn't resolve default property of object avardata(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iEditorCount = avardata(lcounter, 15)
			
            'Call Process_Word_Line(iSourceType, sAuthor, iRecordID, sArticleDesignation, sTitle, sVolume, sJournalTitle, sPage, sMonth, sDay, sYear, sSeriesTitle, sEditor, sEdition, iEditorCount, False, wDocument, sLocation)
			
		Next 
		
		
		wDocument.Selection.WholeStory()
		wDocument.Selection.Find.ClearFormatting()
		wDocument.Selection.Find.Replacement.ClearFormatting()
		With wDocument.Selection.Find
			.Text = "'"
			.Replacement.Text = "'"
			.Forward = True
			.Wrap = Word.WdFindWrap.wdFindAsk
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
		End With
		wDocument.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
		With wDocument.Selection.Find
			.Text = """"
			.Replacement.Text = """"
			.Forward = True
			.Wrap = Word.WdFindWrap.wdFindAsk
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
		End With
		wDocument.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
		wDocument.Selection.Find.ClearFormatting()
		wDocument.Selection.Find.Replacement.ClearFormatting()
		With wDocument.Selection.Find
			.Text = "--"
			.Replacement.Text = "^+"
			.Forward = True
			.Wrap = Word.WdFindWrap.wdFindAsk
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
		End With
		wDocument.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
		
		'wDocument.Visible = True
		
		dDate = Now
		sDate = CStr(dDate)
		Bill_Replace(sDate, "/", "_")
		Bill_Replace(sDate, " ", "_")
		Bill_Replace(sDate, ":", "_")
		If pReportType = "WordPerfect" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object vFileFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vFileFilter = wDocument.FileConverters.Item(3).SaveFormat
			sFileName = sDate & sSessionID & ".wpd"
			wDocument.Documents.Item(1).SaveAs("d:\NCPL\Search\OutputDocs\" & sFileName, vFileFilter)
			'sFileName = "c:\ctest\" & sDate & sSessionID & ".wpd"
			'wDocument.Documents(1).SaveAs sFileName, vFileFilter
			Word_Report = sFileName
		End If
		If pReportType = "Word" Then
			'sFileName = "file:///c:/ctest/" & sDate & ".doc"
			'sFileName = "..\c:\ctest" & sDate & ".doc"
			'sFileName = sDate &  sSessionID & ".doc"
			sFileName = sDate & sSessionID & ".doc"
			
			wDocument.Documents.Item(1).SaveAs("d:\NCPL\Search\OutputDocs\" & sFileName)
			Word_Report = sFileName
		End If
		If pReportType = "Adobe" Then
			'printout
			sPDFFileName = "d:\NCPL\Search\OutputDocs\" & sDate & sSessionID & ".ps"
			sFileName = "d:\NCPL\Search\OutputDocs\" & sDate & sSessionID & ".doc"
			wDocument.ActiveDocument.SaveAs(sFileName)
			
			wDocument.ActivePrinter = "Generic PostScript Printer"
			
			wDocument.ActiveDocument.PrintOut(False,  ,  , sPDFFileName,  ,  ,  ,  ,  ,  , True,  , sFileName)
			'Do Until wDocument.BackgroundPrintingStatus = 0
			'    For i = 1 To 1000
			'    Next
			'Loop
			'wDocument.ActiveDocument.Close
			'Set wDocument = Nothing
			'convert
			'wDocument.Documents(1).SaveAs sPDFFilename
			'wDocument.PrintOut , , , sFileName, , , , , , , False
			sFileName = sPDFFileName
			
			sPDFFileName = "d:\NCPL\Search\OutputDocs\" & sDate & sSessionID & ".pdf"
            'PDFMaker = New ACRODISTXLib.PdfDistiller
			'UPGRADE_WARNING: Couldn't resolve default property of object PDFMaker.FileToPDF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'PDFMaker.FileToPDF(sFileName, sPDFFileName, "")
			'UPGRADE_NOTE: Object PDFMaker may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'PDFMaker = Nothing
			'Word_Report = sPDFFileName
			Word_Report = sDate & sSessionID & ".pdf"
			'PDFMaker.FileToPDF sPDFFilename, sFileName, ""
			'Set PDFMaker = Nothing
		End If
		For i = 1 To wDocument.Documents.Count
			wDocument.Documents.Item(i).Close()
		Next 
		
		wDocument.Quit((0))
		'UPGRADE_NOTE: Object wDocument may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		wDocument = Nothing
		
	End Function
	
	
	Public Function BuildRecordset(ByRef sSource As String, ByRef cnx As ADODB.Connection, ByRef intcols As Short, ByRef lngrows As Short, ByRef lngRow As Short) As Object
		On Error GoTo BuildRecordsetErr
		
		Dim rst As ADODB.Recordset
		Dim rstAuthors As ADODB.Recordset
		Dim rstKeyword As ADODB.Recordset
		Dim rstEditor As ADODB.Recordset
		
		Dim iKeywordCount As Short
		Dim iKeywordCounter As Short
		Dim sKeywordSQL As String
		
		Dim strMsg As String
		Dim ctlMessage As System.Windows.Forms.Control
		Dim ctlRecords As System.Windows.Forms.Control
		Dim bMultipleRecord As Boolean
		
		Dim cAuthors As Collection
		Dim cEditors As Collection
		Dim cTranslators As Collection
		
		Dim iSourceType As Short
		Dim sTitle As String
		Dim sVolume As String
		Dim sAuthorFirstName As String
		Dim sAuthorMiddleName As String
		Dim sAuthorLastName As String
		Dim sPage As String
		Dim sMonth As String
		Dim sYear As String
		Dim sJournalTitle As String
		Dim sInstitutionalAuthor As String
		Dim sInstitutionalEditor As String
		Dim sInstitutionalTranslator As String
		Dim sKeywords As String
		Dim lrecnum As Integer
		Dim sDay As String
		Dim sRowString As String
		Dim sCongress As String
		Dim sNotes As String
		Dim sAuthor As String
		Dim iRecordID As Short
		Dim sAuthorSQLString As String
		Dim iAuthorCount As Short
		Dim icounter As Short
		Dim sArticleDesignation As String
		Dim sEditor As String
		Dim sTranslator As String
		Dim iEditorCount As Short
		Dim iProgressCounter As Short
		Dim sEdition As String
		Dim sConnectionString As String
		Dim iRecordsAETID As Short
        Dim sSeriesTitle As String
        'Dim avardata As Array
		
		Const adhcErrorTooFewParameters As Short = 3061
		Const adhcErrorActionQuery As Short = 3219
		Const adhcErrorDDLQuery As Short = 3324
		
		rst = New ADODB.Recordset
		rstAuthors = New ADODB.Recordset
		rstKeyword = New ADODB.Recordset
		
		rst.Open(sSource, cnx, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
		
		
		If rst.EOF Then GoTo BuildRecordsetErr
		
		
		
		' Switch to in-line error handling
		On Error Resume Next
		
		If Err.Number <> 0 Then
			'Response.Write "Error " & Err.Number & ": " & Err.Description
			
			' Extra help for common errors
			Select Case Err.Number
				Case adhcErrorActionQuery
					'Response.Write "(This error will occur if the SQL entered is an Action query.)"
				Case adhcErrorDDLQuery
					'Response.Write "(This error will occur if the SQL entered is a DDL query.)"
				Case adhcErrorTooFewParameters
					'Response.Write "(This error is often caused by a misspelled table or field name.)"
				Case Else
					' Do nothing
			End Select
			
			On Error GoTo BuildRecordsetErr
			
			intcols = 0
			lngrows = 0
		Else
			On Error GoTo BuildRecordsetErr
			
			If Not rst.EOF Then rst.MoveLast()
			'***changed to 15 for keywordprintout
			'intcols = 14
			intcols = 16
			
			lngrows = rst.RecordCount
			
            'Dim avardata(lngrows, intcols) As Array
			
			If Not rst.EOF Then rst.MoveFirst()
			lngRow = 0
			
			Do While Not rst.EOF
				
				'If ((lngRow And 3) = 0) Then Response.Write ". "
				
				'reset all values
				sTitle = ""
				sJournalTitle = ""
				sVolume = ""
				sMonth = ""
				sYear = ""
				sPage = ""
				sAuthorFirstName = ""
				sAuthorMiddleName = ""
				sAuthorLastName = ""
				sDay = ""
				sRowString = ""
				sCongress = ""
				sNotes = ""
				sAuthor = ""
				sInstitutionalAuthor = ""
				sInstitutionalEditor = ""
				sInstitutionalTranslator = ""
				sKeywords = ""
				sSeriesTitle = ""
				
				iRecordID = 0
				sAuthorSQLString = ""
				iAuthorCount = 0
				iSourceType = 0
				iRecordsAETID = 0
				sArticleDesignation = ""
				sEditor = ""
				sEdition = ""
				iEditorCount = 0
				cAuthors = New Collection
				cEditors = New Collection
				cTranslators = New Collection
				
				'check to see source type
				If rst.Fields("DocumentType").Value = "Journal Article" Then
					If rst.Fields("Pagination").Value = "Consecutive" Then iSourceType = 1
					If rst.Fields("Pagination").Value = "Nonconsecutive" Then iSourceType = 2
				End If
				If rst.Fields("DocumentType").Value = "Treatise" Then iSourceType = 3
				If rst.Fields("DocumentType").Value = "Chapter in Treatise" Then iSourceType = 4
				If rst.Fields("DocumentType").Value = "Legislative Material" Then iSourceType = 5
				'If rst.Fields("DocumentType").Value = "Legislative Report" Then iSourceType = 6
				If rst.Fields("DocumentType").Value = "Unpublished Work" Then iSourceType = 7
				If rst.Fields("DocumentType").Value = "Nonprint Material" Then iSourceType = 6
				iRecordID = rst.Fields(0).Value
				
				'check to see if the query returned duplicate records; if it did, move to next record
				bMultipleRecord = False
				If lngRow > 0 Then
					For iKeywordCounter = 0 To (lngRow - 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(iKeywordCounter, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'If avardata(iKeywordCounter, 3) = iRecordID Then bMultipleRecord = True
					Next 
				End If
				If bMultipleRecord Then
					GoTo Here
				End If
				
				
				
				
				Call Get_AET_String(iRecordID, cnx, sAuthor, sEditor, iAuthorCount, iEditorCount)
				
				
				'fill in year and page number values, check for article designation
				If rst.Fields("PublicationYear").Value <> "" Then sYear = rst.Fields("PublicationYear").Value
				If rst.Fields("PageNumber").Value <> "" Then sPage = rst.Fields("PageNumber").Value
				If rst.Fields("ArticleDesignationForCitation").Value <> "" Then sArticleDesignation = Trim(rst.Fields("ArticleDesignationForCitation").Value)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 11). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 11) = sYear
				
				'begin building "Rowstring," which will vary according to the source type
				
				'If (iSourceType <> 5) And (iSourceType <> 6) Then 'don't do anything with legislative materi
				If (iSourceType <> 5) Then 'don't do anything with legislative materi
					If sAuthor <> "" Then sRowString = sRowString & sAuthor & ", "
					If (sArticleDesignation <> "") And (sArticleDesignation <> "book review") And (sArticleDesignation <> "Book Review") Then sRowString = sRowString & sArticleDesignation & ", "
					
				End If
				Select Case iSourceType
					Case 1 'consecutively paginated
						
						If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						sTitle = Trim(sTitle)
						If rst.Fields("JournalTitleShortForm").Value <> "" Then sJournalTitle = rst.Fields("JournalTitleShortForm").Value Else If rst.Fields("JournalTitle").Value <> "" Then sJournalTitle = rst.Fields("JournalTitle").Value
						
						If rst.Fields("Volume").Value <> "" Then sVolume = rst.Fields("Volume").Value
						
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 2) = sAuthor
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 5) = sTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 7) = sJournalTitle
						
						If sTitle <> "" Then sRowString = sRowString & "<i>" & sTitle & "</i>" & ", "
						
						sRowString = sRowString & sVolume
						If sVolume <> "" Then sRowString = sRowString & " "
						sJournalTitle = Make_Smallcaps(sJournalTitle)
						sRowString = sRowString & sJournalTitle
						'"<span style=" & Chr(34) & "font-variant: small-caps" & Chr(34) & "> " & sJournalTitle & "</span>" & " "
						If sPage <> "" Then sRowString = sRowString & " " & sPage & " "
						
						If sYear <> sVolume Then sRowString = sRowString & " (" & sYear & ")"
						If (sArticleDesignation = "book review") Or (sArticleDesignation = "Book Review") Then sRowString = sRowString & " (book review)." Else sRowString = sRowString & "."
						
						
						
					Case 2 'nonconsecutively paginated
						If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						sTitle = Trim(sTitle)
						
						If rst.Fields("JournalTitleShortForm").Value <> "" Then sJournalTitle = rst.Fields("JournalTitleShortForm").Value Else If rst.Fields("JournalTitle").Value <> "" Then sJournalTitle = rst.Fields("JournalTitle").Value
						If rst.Fields("PublicationMonthOrSeason").Value <> "" Then sMonth = rst.Fields("PublicationMonthOrSeason").Value
						If rst.Fields("ArticlePublicationDay").Value <> "" Then sDay = rst.Fields("ArticlePublicationDay").Value
						
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 2) = sAuthor
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 5) = sTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 7) = sJournalTitle
						
						If sTitle <> "" Then sRowString = sRowString & "<i>" & sTitle & "</i>" & ", "
						sJournalTitle = Make_Smallcaps(sJournalTitle)
						sRowString = sRowString & sJournalTitle & ", "
						sRowString = sRowString & sMonth & " "
						If sDay <> "" Then sRowString = sRowString & sDay & ", "
						sRowString = sRowString & sYear
						If sPage <> "" Then sRowString = sRowString & ", at " & sPage
						If (sArticleDesignation = "Book Review") Or (sArticleDesignation = "book review") Then sRowString = sRowString & " (book review)." Else sRowString = sRowString & "."
						
						'sRowString = sRowString & "."
					Case 3 'treatise
						If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						sTitle = Trim(sTitle)
						
						If rst.Fields("TreatiseEditionAndPrinting").Value <> "" Then sEdition = rst.Fields("TreatiseEditionAndPrinting").Value
						If rst.Fields("TreatiseTitleOfSeriesIfNotIssuedByAuthor").Value <> "" Then sSeriesTitle = rst.Fields("TreatiseTitleOfSeriesIfNotIssuedByAuthor").Value
						
						'If sInstitutionalAuthor <> "" Then sRowString = sRowString & sInstitutionalAuthor & ", "
						
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 2) = sAuthor
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 5) = sTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 7) = sJournalTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 14) = sSeriesTitle
						
						
						sRowString = Make_Smallcaps(sRowString)
						sTitle = Make_Smallcaps(sTitle)
						If sTitle <> "" Then sRowString = sRowString & sTitle & " "
						'sRowString = sRowString & sPage
						If sEdition <> "" Then sYear = sEdition & " " & sYear
						If sEditor <> "" Then
							If iEditorCount = 1 Then sYear = sEditor & " ed., " & sYear
							If iEditorCount > 1 Then sYear = sEditor & " eds., " & sYear
						End If
						
						If Not ((sSeriesTitle = "") And (sYear = "")) Then
							'               If Not (sSeriesTitle = "") Then
							
							sRowString = sRowString & "("
							If sSeriesTitle <> "" Then sRowString = sRowString & sSeriesTitle
							If Not sVolume = "" Then sRowString = sRowString & " " & sVolume
							If Not sSeriesTitle = "" And sVolume = "" Then sRowString = sRowString & ", "
							If sYear <> "" Then sRowString = sRowString & sYear
							
							sRowString = sRowString & ")."
						End If
						'If sSeriesTitle <> "" Then sRowString = sRowString & sSeriesTitle & " No. " & sVolume & ", "
						'If sYear <> "" Then sRowString = sRowString & sYear
						
						'If sYear <> "" Then sRowString = sRowString & "(" & sYear & ")" Else
						
						'sRowString = sRowString & "."
						
					Case 4 'chapter in treatise
						'If TitleOfSeriesIfNotIssuedByAuthor <> "" Then
						If rst.Fields("LargerWorkSeriesVolume").Value <> "" Then sVolume = rst.Fields("LargerWorkSeriesVolume").Value
						If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						sTitle = Trim(sTitle)
						
						If rst.Fields("LargerWorkTitle").Value <> "" Then sJournalTitle = rst.Fields("LargerWorkTitle").Value
						If rst.Fields("LargerWorkTitleOfSeriesIfNotIssuedByAuthor").Value <> "" Then sSeriesTitle = rst.Fields("LargerWorkTitleOfSeriesIfNotIssuedByAuthor").Value
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 2) = sAuthor
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 5) = sTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 7) = sJournalTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 14). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 14) = sSeriesTitle
						
						If sEdition <> "" Then sYear = sEdition & " " & sYear
						If sEditor <> "" Then
							If iEditorCount = 1 Then sYear = sEditor & " ed., " & sYear
							If iEditorCount > 1 Then sYear = sEditor & " eds., " & sYear
						End If
						
						If sJournalTitle <> "" Then 'if there is a title for the larger work
							If sTitle <> "" Then sRowString = sRowString & "<i>" & sTitle & "</i>" & ", <i>in</i> "
							sJournalTitle = Make_Smallcaps(sJournalTitle)
							If sSeriesTitle = "" And sVolume <> "" Then
								sRowString = sRowString & sVolume & " "
							End If
							If sJournalTitle <> "" Then sRowString = sRowString & sJournalTitle & " "
							If sPage <> "" Then sRowString = sRowString & sPage & " "
							sRowString = sRowString & "("
							If sSeriesTitle <> "" Then
								sRowString = sRowString & sSeriesTitle
								If sVolume <> "" Then sRowString = sRowString & " No. " & sVolume
								sRowString = sRowString & ", "
							End If
							If sYear <> "" Then sRowString = sRowString & sYear
							sRowString = sRowString & ")."
						Else
							'sRowString = Make_Smallcaps(sRowString)
							'sTitle = Make_Smallcaps(sTitle)
							'sSeriesTitle = Make_Smallcaps(sSeriesTitle)
							If sSeriesTitle <> "" Then
								sRowString = sRowString & sSeriesTitle
								If sVolume <> "" Then sRowString = sRowString & " No. " & sVolume
								sRowString = sRowString & ", "
							End If
							If sTitle <> "" Then sRowString = sRowString & sTitle & " "
							sRowString = Make_Smallcaps(sRowString)
							
							If sPage <> "" Then sRowString = sRowString & sPage & " "
							If sYear <> "" Then sRowString = sRowString & "(" & sYear & ")" Else 
							
							sRowString = sRowString & "."
						End If
						
						
						
						'Else
						'If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						'If rst.Fields("LargerWorkTitle").Value <> "" Then sJournalTitle = rst.Fields("LargerWorkTitle").Value
						
						'If sTitle <> "" Then sRowString = sRowString & sTitle & ", in "
						'sJournalTitle = Make_Smallcaps(sJournalTitle)
						'
						'If sJournalTitle <> "" Then sRowString = sRowString & sJournalTitle & " "
						'If sPage <> "" Then sRowString = sRowString & sPage & " "
						'If sYear <> "" Then sRowString = sRowString & "(" & sYear & ")"
						'sRowString = sRowString & "."
						' En'd If
						
						
						'Case 5 'legislative hearing
						'    If rst.Fields("LegislativeTitle").Value <> "" Then sTitle = rst.Fields("LegislativeTitle").Value
						'    If rst.Fields("NumberofCongress").Value <> "" Then sCongress = rst.Fields("NumberofCongress").Value
						'    If rst.Fields("Notes").Value <> "" Then sNotes = rst.Fields("Notes").Value
						'
						'    If sTitle <> "" Then sRowString = sRowString & sTitle & ", "
						'    If sCongress <> "" Then sRowString = sRowString & sCongress & " "
						'    If sPage <> "" Then sRowString = sRowString & sPage
						'    If sYear <> "" Then sRowString = sRowString & " (" & sYear & ")"
						'    If sNotes <> "" Then sRowString = sRowString & " (" & sNotes & ")"
						'    sRowString = sRowString & "."
						
						'Case 6 'legislative report
						'we need to add functionality to the database to handle these. Fields are missing
						
						
					Case 6, 7 'unpublished work or nonprint
						If rst.Fields("Title").Value <> "" Then sTitle = rst.Fields("Title").Value
						sTitle = Trim(sTitle)
						
						If rst.Fields("UnpublishedMonth").Value <> "" Then sMonth = rst.Fields("UnpublishedMonth").Value
						If rst.Fields("UnpublishedDay").Value <> "" Then sDay = rst.Fields("UnpublishedDay").Value
						
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 2) = sAuthor
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 5) = sTitle
						'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 7). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'avardata(lngRow, 7) = sJournalTitle
						
						If sTitle <> "" Then sRowString = sRowString & sTitle & " ("
						
						If sMonth <> "" Then sRowString = sRowString & sMonth & " "
						If sDay <> "" Then sRowString = sRowString & sDay & ", "
						If sYear <> "" Then sRowString = sRowString & sYear
						If iSourceType = 7 Then sRowString = sRowString & ") (unpublished work)."
					Case Else
						' Do nothing
				End Select
				
				'add "Rowstring" to the listbox
				'sRowString = "(Recnum: " & rst!RecordID & ") " & sRowString
				'Response.Write sRowString & "<P>"
				'Me.lboResults.AddItem ((lngRow + 1) & ". " & sRowString)
				
				'add values to avardata array to use later for MS Word report
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 0) = sRowString
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 1) = iSourceType
				'avarData(lngRow, 2) = sAuthor
				'avarData(lngRow, 5) = sTitle
				'avarData(lngRow, 7) = sJournalTitle
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 3) = iRecordID
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 4) = sArticleDesignation
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 6) = sVolume
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 8). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 8) = sPage
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 9). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 9) = sMonth
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 10). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 10) = sDay
				'avardata(lngRow, 11) = sYear
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 12). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 12) = sEditor
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 13). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 13) = sEdition
				'UPGRADE_WARNING: Couldn't resolve default property of object avardata(lngRow, 15). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'avardata(lngRow, 15) = iEditorCount
				
				'add code for keywordcount
				'avardata(lngRow, 14) = sKeywords
				
				lngRow = lngRow + 1
				'ctlRecords = lngRow
				'update progress bar
				'Me.Refresh
				'Me.txtRecords.Refresh
Here: 'used for skipping duplicate records
				rst.MoveNext()
				
			Loop 
			
		End If
		'Set cnx = Nothing
		'Set rst = Nothing
		'UPGRADE_NOTE: Object rstAuthors may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthors = Nothing
		'UPGRADE_NOTE: Object rstKeyword may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstKeyword = Nothing
		If lngRow > 0 Then
			'    strMsg = "Done."
			'    Call Enabler(True)
			'    Else:
			'        strMsg = "No matching records found."
			'        MsgBox strMsg
		End If
		'ctlMessage = strMsg
		'Timer1.Interval = 0
		'bRunProgram = False
		'frmSearch.Hide
		'frmSearch.ProgressBar1.Enabled = False
		'frmSearch.ProgressBar1.Visible = False
		
		'UPGRADE_WARNING: Couldn't resolve default property of object BuildRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'BuildRecordset = VB6.CopyArray(avardata)
BuildRecordsetDone: 
		On Error GoTo 0
		
		Exit Function
		
BuildRecordsetErr: 
		Select Case Err.Number
			Case Else
				'Response.Write "Error#" & Err.Number & ": " & Err.Description
				
		End Select
		Resume BuildRecordsetDone
		
	End Function
	Public Function Make_Smallcaps(ByRef pString As String) As String
		Dim i As Short
		Dim iLength As Short
		Dim sChar As String
		Dim tempstring As String
		
		tempstring = ""
		iLength = Len(pString)
		
		For i = 1 To iLength
			sChar = Mid(pString, i, 1)
			'MsgBox sChar & " " & Asc(sChar)
			If ((Asc(sChar) > 47) And (Asc(sChar) < 92)) Or (sChar = "(") Or (sChar = ")") Then
				tempstring = tempstring & "<span style=" & Chr(34) & "text-transform: capitalize" & Chr(34) & ">" & sChar & "</span>"
			Else
				'tempstring = tempstring & "<span style=" & Chr(34) & "font-variant: small-caps" & Chr(34) & ">" & sChar & "</span>"
				tempstring = tempstring & "<span style=" & Chr(34) & "text-transform: uppercase" & Chr(34) & "><font size=" & Chr(34) & "2" & Chr(34) & ">" & sChar & "</font></span>"
				
				
			End If
		Next 
		Make_Smallcaps = tempstring
	End Function
	
	Public Function Bill_Replace(ByRef sString As String, ByRef sStringToReplace As String, ByRef sReplacementString As String) As String
		Dim lReplacePos As Integer
		Dim lReplaceLength As Integer
		Dim sLeftString As String
		Dim sRightString As String
		lReplacePos = 0
		lReplaceLength = 1
		Do While InStr(lReplacePos + lReplaceLength, sString, sStringToReplace) <> 0
			lReplacePos = InStr(lReplacePos + lReplaceLength, sString, sStringToReplace)
			sLeftString = Left(sString, lReplacePos - 1)
			sRightString = Right(sString, Len(sString) - (lReplacePos - 1 + Len(sStringToReplace)))
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
	
	Public Function MakeSmallcaps(ByRef pString As String) As String
		Dim i As Short
		Dim iLength As Short
		Dim sChar As String
		Dim tempstring As String
		
		tempstring = ""
		iLength = Len(pString)
		
		For i = 1 To iLength
			sChar = Mid(pString, i, 1)
			'MsgBox sChar & " " & Asc(sChar)
			If ((Asc(sChar) > 47) And (Asc(sChar) < 92)) Or (sChar = "(") Or (sChar = ")") Then
				tempstring = tempstring & "<span style=" & Chr(34) & "text-transform: capitalize" & Chr(34) & ">" & sChar & "</span>"
			Else
				'tempstring = tempstring & "<span style=" & Chr(34) & "font-variant: small-caps" & Chr(34) & ">" & sChar & "</span>"
				tempstring = tempstring & "<span style=" & Chr(34) & "text-transform: uppercase" & Chr(34) & "><font size=" & Chr(34) & "2" & Chr(34) & ">" & sChar & "</font></span>"
				
				
			End If
		Next 
		MakeSmallcaps = tempstring
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
		sCurrentItem = Left(sWorkingString, cSpacePos.Item(1) - 1)
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
			sCurrentItem = Right(sWorkingString, Len(sWorkingString) - cSpacePos.Item(cSpacePos.Count()))
			cTempColl.Add(sCurrentItem)
		End If
		sCurrentItem = ""
		For i = 1 To cTempColl.Count()
			'If Left(cTempColl.Item(i), 1) = "'" Then cQuotePosColl.Add i
			'If Right(cTempColl.Item(i), 1) = "'" Then cQuotePosColl.Add i
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(cTempColl.Item(i), 1) = Chr(34) Then cQuotePosColl.Add(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Right(cTempColl.Item(i), 1) = Chr(34) Then cQuotePosColl.Add(i)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(cTempColl.Item(i), 1) = "[" Then cLBracketPos.Add(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(cTempColl.Item(i), 1) = "]" Then cRBracketPos.Add(i)
			
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
						If Left(cTempColl.Item(j), 1) = Chr(34) Then
							sCurrentItem = Right(cTempColl.Item(j), Len(cTempColl.Item(j)) - 1)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCurrentItem = sCurrentItem & " " & cTempColl.Item(j)
						End If 'get rid of quotation
						k = k + 1
					Next 
					If Right(sCurrentItem, 1) = Chr(34) Then sCurrentItem = Left(sCurrentItem, Len(sCurrentItem) - 1) 'get rid of quotation
					k = k - 1
				End If
			Next 
			'UPGRADE_WARNING: Couldn't resolve default property of object cTempColl.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not (bPartOfPhrase) Then sCurrentItem = cTempColl.Item(k)
			'Bill_Replace sWorkingString, "'", "''" ' to take care of O'leary, etc, and journal title short form
			
			pCollection.Add(sCurrentItem)
		Next 
	End Sub
	
	Public Function Build_AET_String(ByRef cAET As Collection, ByRef sInstitutionalEntity As Object) As String
		Dim iAETCount As Short
		Dim sAET As String
		Dim icounter As Short
		iAETCount = cAET.Count()
		
		If (iAETCount = 1) Or (iAETCount = 2) Then
			For icounter = 1 To iAETCount
				'UPGRADE_WARNING: Couldn't resolve default property of object cAET.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sAET = sAET & cAET.Item(icounter)
				If iAETCount = 2 Then If icounter = 1 Then sAET = sAET & " & "
			Next 
		End If
		
		If iAETCount > 2 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object cAET.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sAET = cAET.Item(1)
			sAET = sAET & " et al."
			
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object sInstitutionalEntity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (sAET = "") And (sInstitutionalEntity <> "") Then
			sAET = sInstitutionalEntity
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sInstitutionalEntity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sInstitutionalEntity <> "" Then sAET = sAET & ", " & sInstitutionalEntity
		End If
		Build_AET_String = sAET
		
	End Function
	
	Public Sub Get_AET_String(ByRef iRecordID As Short, ByRef cnx As ADODB.Connection, ByRef sAuthor As Object, ByRef sEditor As Object, ByRef iAuthorCount As Object, ByRef iEditorCount As Object, Optional ByRef cAETIDs As Collection = Nothing)
		Dim sAuthorSQLString As String
		Dim rstAuthors As ADODB.Recordset
		Dim cAuthors As Collection
		Dim cEditors As Collection
		Dim cTranslators As Collection
		Dim sAuthorFirstName As String
		Dim sAuthorMiddleName As String
		Dim sAuthorLastName As String
		Dim sAuthorSuffix As String
		Dim sInstitutionalAuthor As String
		Dim sInstitutionalEditor As String
		Dim sInstitutionalTranslator As String
		Dim i As Short
		
		cAuthors = New Collection
		cEditors = New Collection
		cTranslators = New Collection
		
		rstAuthors = New ADODB.Recordset
		If iRecordID <> 0 Then
			sAuthorSQLString = "select * from qryAET WHERE qryAET.RecordID=" & iRecordID '& " ORDER BY tblRecordsAET.RecordsAETID"
		Else
			sAuthorSQLString = "select * from tblAuthorsEditorsTranslators WHERE "
			For i = 1 To cAETIDs.Count()
				If i > 1 Then sAuthorSQLString = sAuthorSQLString & " OR "
				'UPGRADE_WARNING: Couldn't resolve default property of object cAETIDs.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sAuthorSQLString = sAuthorSQLString & "(AETID=" & cAETIDs.Item(i) & ")"
				
			Next 
		End If
		'rstAuthors.Open sAuthorSQLString, cnx, adopenforwardonly,adLockReadOnly, adLockReadOnly, adCmdText
		rstAuthors.Open(sAuthorSQLString, cnx, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		Do While Not rstAuthors.EOF 'loop through once and add any institutional Entities first
			If rstAuthors.Fields("InstitutionalEntity").Value <> "" Then
				
				Select Case rstAuthors.Fields("AETType").Value
					Case "Author"
						If sInstitutionalAuthor = "" Then
							sInstitutionalAuthor = rstAuthors.Fields("InstitutionalEntity").Value
						Else
							sInstitutionalAuthor = sInstitutionalAuthor & ", " & rstAuthors.Fields("InstitutionalEntity").Value
						End If
					Case "Editor"
						If sInstitutionalEditor = "" Then
							sInstitutionalEditor = rstAuthors.Fields("InstitutionalEntity").Value
						Else
							sInstitutionalEditor = sInstitutionalEditor & ", " & rstAuthors.Fields("InstitutionalEntity").Value
						End If
					Case "Translator"
						If sInstitutionalTranslator = "" Then
							sInstitutionalTranslator = rstAuthors.Fields("InstitutionalEntity").Value
						Else
							sInstitutionalTranslator = sInstitutionalTranslator & ", " & rstAuthors.Fields("InstitutionalEntity").Value
						End If
				End Select
			End If
			rstAuthors.MoveNext()
			
		Loop 
		
		If rstAuthors.RecordCount > 0 Then rstAuthors.MoveFirst()
		Do While Not rstAuthors.EOF
			sAuthorFirstName = ""
			sAuthorMiddleName = ""
			sAuthorLastName = ""
			sAuthorSuffix = ""
			'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sAuthor = ""
			If rstAuthors.Fields("FirstName").Value <> "" Then sAuthorFirstName = rstAuthors.Fields("FirstName").Value
			If rstAuthors.Fields("MiddleName").Value <> "" Then sAuthorMiddleName = rstAuthors.Fields("MiddleName").Value
			If rstAuthors.Fields("LastName").Value <> "" Then sAuthorLastName = rstAuthors.Fields("LastName").Value
			If rstAuthors.Fields("Suffix").Value <> "" Then sAuthorSuffix = rstAuthors.Fields("Suffix").Value
			
			'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sAuthorFirstName <> "" Then sAuthor = sAuthor & sAuthorFirstName & " "
			'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sAuthorMiddleName <> "" Then sAuthor = sAuthor & sAuthorMiddleName & " "
			'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If sAuthorLastName <> "" Then sAuthor = sAuthor & sAuthorLastName
			If sAuthorSuffix <> "" Then
				If sAuthorSuffix = "Jr." Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sAuthor = sAuthor & ", " & sAuthorSuffix
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sAuthor = sAuthor & " " & sAuthorSuffix
				End If
			End If
			Select Case rstAuthors.Fields("AETType").Value
				Case "Author"
					'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If sAuthor <> "" Then cAuthors.Add(sAuthor)
					
				Case "Editor"
					'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If sAuthor <> "" Then cEditors.Add(sAuthor)
					
				Case "Translator"
					'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If sAuthor <> "" Then cTranslators.Add(sAuthor)
			End Select
			rstAuthors.MoveNext()
		Loop 
		
		rstAuthors.Close()
		'UPGRADE_NOTE: Object rstAuthors may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthors = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object iAuthorCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iAuthorCount = cAuthors.Count()
		'UPGRADE_WARNING: Couldn't resolve default property of object iEditorCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iEditorCount = cEditors.Count()
		'UPGRADE_WARNING: Couldn't resolve default property of object sAuthor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sAuthor = Build_AET_String(cAuthors, sInstitutionalAuthor)
		'UPGRADE_WARNING: Couldn't resolve default property of object sEditor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sEditor = Build_AET_String(cEditors, sInstitutionalEditor)
	End Sub
End Class