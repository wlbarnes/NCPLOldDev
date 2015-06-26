<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _mnuFile_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mneNewAuthor_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuNewJournal_4 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuNewKeyword_5 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuAdd_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents chkRepublished As System.Windows.Forms.CheckBox
	Public WithEvents txtJournaTitleShortForm As System.Windows.Forms.TextBox
	Public WithEvents cmdPreview As System.Windows.Forms.Button
	Public WithEvents lblRecordNumber As System.Windows.Forms.Button
	Public WithEvents chkLibraryCollection As System.Windows.Forms.CheckBox
	Public WithEvents lblArrow2 As System.Windows.Forms.TextBox
	Public WithEvents lblDblClicktoAdd2 As System.Windows.Forms.TextBox
	Public WithEvents lblStatus As System.Windows.Forms.TextBox
	Public WithEvents cmdDelete As System.Windows.Forms.Button
	Public WithEvents cmdEditJournal As System.Windows.Forms.Button
	Public WithEvents txtCallNumber As System.Windows.Forms.TextBox
	Public WithEvents cmdNewLargerWork As System.Windows.Forms.Button
	Public WithEvents lblOriginalPublicationDate As System.Windows.Forms.TextBox
	Public WithEvents lblPublisher As System.Windows.Forms.TextBox
	Public WithEvents lblCallNumber As System.Windows.Forms.TextBox
	Public WithEvents lblEditionAndPrinting As System.Windows.Forms.TextBox
	Public WithEvents lblMiscType As System.Windows.Forms.TextBox
	Public WithEvents lblLocation As System.Windows.Forms.TextBox
	Public WithEvents lblThesisDissertationType As System.Windows.Forms.TextBox
	Public WithEvents lblUnpublishedType As System.Windows.Forms.TextBox
	Public WithEvents lblUSCCANCitation As System.Windows.Forms.TextBox
	Public WithEvents lblReportOrDocumentNumber As System.Windows.Forms.TextBox
	Public WithEvents lblLegislativeHouse As System.Windows.Forms.TextBox
	Public WithEvents lblNumberOfCongress As System.Windows.Forms.TextBox
	Public WithEvents lblSessionOfCongress As System.Windows.Forms.TextBox
	Public WithEvents lblStateLegislativeSession As System.Windows.Forms.TextBox
	Public WithEvents lblSuDocNumber As System.Windows.Forms.TextBox
	Public WithEvents lblLegislativeType As System.Windows.Forms.TextBox
	Public WithEvents lblSeriesVolume As System.Windows.Forms.TextBox
	Public WithEvents lblTitleOfSeriesIfNotIssuedByAuthor As System.Windows.Forms.TextBox
	Public WithEvents lblLargerWorkTitle As System.Windows.Forms.TextBox
	Public WithEvents cmbPagination As System.Windows.Forms.ComboBox
	Public WithEvents lblVolume As System.Windows.Forms.TextBox
	Public WithEvents lblPublicationMonthOrSeason As System.Windows.Forms.TextBox
	Public WithEvents lblPage As System.Windows.Forms.TextBox
	Public WithEvents chkSource As System.Windows.Forms.CheckBox
	Public WithEvents lblPublicationDay As System.Windows.Forms.TextBox
	Public WithEvents chkYear As System.Windows.Forms.CheckBox
	Public WithEvents lblKeywords As System.Windows.Forms.TextBox
	Public WithEvents lblJournalTitle As System.Windows.Forms.TextBox
	Public WithEvents lblSourceType As System.Windows.Forms.TextBox
	Public WithEvents lblArticleDesignation As System.Windows.Forms.TextBox
	Public WithEvents lblInputInitials As System.Windows.Forms.TextBox
	Public WithEvents lblDateUpdated As System.Windows.Forms.TextBox
	Public WithEvents lblPublicationYear As System.Windows.Forms.TextBox
	Public WithEvents txtStatus As System.Windows.Forms.TextBox
	Public WithEvents lstNewKeywords As System.Windows.Forms.ListBox
	Public WithEvents cmdGetNewKeywords As System.Windows.Forms.Button
	Public WithEvents cmdNewAuthor As System.Windows.Forms.Button
	Public WithEvents cmdNewJournal As System.Windows.Forms.Button
	Public WithEvents txtMiscID As System.Windows.Forms.TextBox
	Public WithEvents txtUnpublishedID As System.Windows.Forms.TextBox
	Public WithEvents txtLegislativeID As System.Windows.Forms.TextBox
	Public WithEvents txtTreatiseID As System.Windows.Forms.TextBox
	Public WithEvents txtChapterID As System.Windows.Forms.TextBox
	Public WithEvents txtArticleID As System.Windows.Forms.TextBox
	Public WithEvents cmbRecordNumber As System.Windows.Forms.ComboBox
	Public WithEvents cmdNextRecord As System.Windows.Forms.Button
	Public WithEvents cmdPreviousRecord As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents lstKeywords As System.Windows.Forms.ListBox
	Public WithEvents lstCurrentKeywords As System.Windows.Forms.ListBox
	Public WithEvents cmbAETChoice As System.Windows.Forms.ComboBox
	Public WithEvents txtSuDocNumber As System.Windows.Forms.TextBox
	Public WithEvents txtLargerWorkID As System.Windows.Forms.TextBox
	Public WithEvents cmbLargerWorkTitle As System.Windows.Forms.ComboBox
	Public WithEvents txtReportOrDocumentNumber As System.Windows.Forms.TextBox
	Public WithEvents txtUSCCANCitation As System.Windows.Forms.TextBox
	Public WithEvents txtStateLegislativeSession As System.Windows.Forms.TextBox
	Public WithEvents txtSessionOfCongress As System.Windows.Forms.TextBox
	Public WithEvents txtNumberOfCongress As System.Windows.Forms.TextBox
	Public WithEvents txtLegislativeHouse As System.Windows.Forms.TextBox
	Public WithEvents cmbLegislativeType As System.Windows.Forms.ComboBox
	Public WithEvents cmbMiscType As System.Windows.Forms.ComboBox
	Public WithEvents txtLocation As System.Windows.Forms.TextBox
	Public WithEvents cmbUnpublishedType As System.Windows.Forms.ComboBox
	Public WithEvents cmbThesisDissertationType As System.Windows.Forms.ComboBox
	Public WithEvents chkAllChaptersBySameAuthor As System.Windows.Forms.CheckBox
	Public WithEvents txtTitleOfSeriesIfNotIssuedByAuthor As System.Windows.Forms.TextBox
	Public WithEvents txtSeriesVolume As System.Windows.Forms.TextBox
	Public WithEvents txtOriginalPublicationDate As System.Windows.Forms.TextBox
	Public WithEvents txtPublisher As System.Windows.Forms.TextBox
	Public WithEvents txtEditionAndPrinting As System.Windows.Forms.TextBox
	Public WithEvents txtOrganizationIssuingNewsletter As System.Windows.Forms.TextBox
	Public WithEvents txtNotes As System.Windows.Forms.TextBox
	Public WithEvents txtPage As System.Windows.Forms.TextBox
	Public WithEvents cmbPublicationMonthOrSeason As System.Windows.Forms.ComboBox
	Public WithEvents txtVolume As System.Windows.Forms.TextBox
	Public WithEvents txtPublicationDay As System.Windows.Forms.TextBox
	Public WithEvents txtJournalID As System.Windows.Forms.TextBox
	Public WithEvents cmbJournalTitle As System.Windows.Forms.ComboBox
	Public WithEvents cmbArticleDesignation As System.Windows.Forms.ComboBox
	Public WithEvents txtTitle As System.Windows.Forms.TextBox
	Public WithEvents txtYear As System.Windows.Forms.TextBox
	Public WithEvents txtInputInitials As System.Windows.Forms.TextBox
	Public WithEvents txtDateUpdated As System.Windows.Forms.TextBox
	Public WithEvents txtDateAdded As System.Windows.Forms.TextBox
	Public WithEvents cmbSourceType As System.Windows.Forms.ComboBox
	Public WithEvents lstAuthors As System.Windows.Forms.ListBox
	Public WithEvents lstTranslators As System.Windows.Forms.ListBox
	Public WithEvents lstCurrentAuthors As System.Windows.Forms.ListBox
	Public WithEvents lstCurrentTranslators As System.Windows.Forms.ListBox
	Public WithEvents lstCurrentEditors As System.Windows.Forms.ListBox
	Public WithEvents lstEditors As System.Windows.Forms.ListBox
	Public WithEvents lblT As System.Windows.Forms.TextBox
	Public WithEvents lblE As System.Windows.Forms.TextBox
	Public WithEvents lblA As System.Windows.Forms.TextBox
	Public WithEvents lblAETChoice As System.Windows.Forms.TextBox
	Public WithEvents chkKeepSelected As System.Windows.Forms.CheckBox
	Public WithEvents lblDoubleClickToAdd As System.Windows.Forms.TextBox
	Public WithEvents lblArrow As System.Windows.Forms.TextBox
	Public WithEvents lblTitle As System.Windows.Forms.TextBox
	Public WithEvents lblYear As System.Windows.Forms.TextBox
	Public WithEvents frmEntryInfo As System.Windows.Forms.GroupBox
	Public WithEvents frmRecordInfo As System.Windows.Forms.GroupBox
	Public WithEvents frmCitationInfo As System.Windows.Forms.GroupBox
	Public WithEvents frmAuthorInfo As System.Windows.Forms.GroupBox
	Public WithEvents frmKeywordInfo As System.Windows.Forms.GroupBox
	Public WithEvents frmNotes As System.Windows.Forms.GroupBox
	Public WithEvents lblSeparateBottom As System.Windows.Forms.Label
	Public WithEvents lblMiscID As System.Windows.Forms.Label
	Public WithEvents lblTreatiseID As System.Windows.Forms.Label
	Public WithEvents lblUnpublishedID As System.Windows.Forms.Label
	Public WithEvents lblChapterID As System.Windows.Forms.Label
	Public WithEvents lblArticleID As System.Windows.Forms.Label
	Public WithEvents lblLegisID As System.Windows.Forms.Label
	Public WithEvents tglNewRecords As AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
	Public WithEvents tglUpdateRecords As AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
	Public WithEvents tglImportRecords As AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
	Public WithEvents lblLargerWorkID As System.Windows.Forms.Label
	Public WithEvents lblNotes As System.Windows.Forms.Label
	Public WithEvents lblSeparateTop As System.Windows.Forms.Label
	Public WithEvents mneNewAuthor As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuAdd As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuFile As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuNewJournal As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuNewKeyword As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me._mnuFile_1 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuAdd_2 = New System.Windows.Forms.ToolStripMenuItem
		Me._mneNewAuthor_3 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuNewJournal_4 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuNewKeyword_5 = New System.Windows.Forms.ToolStripMenuItem
		Me.chkRepublished = New System.Windows.Forms.CheckBox
		Me.txtJournaTitleShortForm = New System.Windows.Forms.TextBox
		Me.cmdPreview = New System.Windows.Forms.Button
		Me.lblRecordNumber = New System.Windows.Forms.Button
		Me.chkLibraryCollection = New System.Windows.Forms.CheckBox
		Me.lblArrow2 = New System.Windows.Forms.TextBox
		Me.lblDblClicktoAdd2 = New System.Windows.Forms.TextBox
		Me.lblStatus = New System.Windows.Forms.TextBox
		Me.cmdDelete = New System.Windows.Forms.Button
		Me.cmdEditJournal = New System.Windows.Forms.Button
		Me.txtCallNumber = New System.Windows.Forms.TextBox
		Me.cmdNewLargerWork = New System.Windows.Forms.Button
		Me.lblOriginalPublicationDate = New System.Windows.Forms.TextBox
		Me.lblPublisher = New System.Windows.Forms.TextBox
		Me.lblCallNumber = New System.Windows.Forms.TextBox
		Me.lblEditionAndPrinting = New System.Windows.Forms.TextBox
		Me.lblMiscType = New System.Windows.Forms.TextBox
		Me.lblLocation = New System.Windows.Forms.TextBox
		Me.lblThesisDissertationType = New System.Windows.Forms.TextBox
		Me.lblUnpublishedType = New System.Windows.Forms.TextBox
		Me.lblUSCCANCitation = New System.Windows.Forms.TextBox
		Me.lblReportOrDocumentNumber = New System.Windows.Forms.TextBox
		Me.lblLegislativeHouse = New System.Windows.Forms.TextBox
		Me.lblNumberOfCongress = New System.Windows.Forms.TextBox
		Me.lblSessionOfCongress = New System.Windows.Forms.TextBox
		Me.lblStateLegislativeSession = New System.Windows.Forms.TextBox
		Me.lblSuDocNumber = New System.Windows.Forms.TextBox
		Me.lblLegislativeType = New System.Windows.Forms.TextBox
		Me.lblSeriesVolume = New System.Windows.Forms.TextBox
		Me.lblTitleOfSeriesIfNotIssuedByAuthor = New System.Windows.Forms.TextBox
		Me.lblLargerWorkTitle = New System.Windows.Forms.TextBox
		Me.cmbPagination = New System.Windows.Forms.ComboBox
		Me.lblVolume = New System.Windows.Forms.TextBox
		Me.lblPublicationMonthOrSeason = New System.Windows.Forms.TextBox
		Me.lblPage = New System.Windows.Forms.TextBox
		Me.chkSource = New System.Windows.Forms.CheckBox
		Me.lblPublicationDay = New System.Windows.Forms.TextBox
		Me.chkYear = New System.Windows.Forms.CheckBox
		Me.lblKeywords = New System.Windows.Forms.TextBox
		Me.lblJournalTitle = New System.Windows.Forms.TextBox
		Me.lblSourceType = New System.Windows.Forms.TextBox
		Me.lblArticleDesignation = New System.Windows.Forms.TextBox
		Me.lblInputInitials = New System.Windows.Forms.TextBox
		Me.lblDateUpdated = New System.Windows.Forms.TextBox
		Me.lblPublicationYear = New System.Windows.Forms.TextBox
		Me.txtStatus = New System.Windows.Forms.TextBox
		Me.lstNewKeywords = New System.Windows.Forms.ListBox
		Me.cmdGetNewKeywords = New System.Windows.Forms.Button
		Me.cmdNewAuthor = New System.Windows.Forms.Button
		Me.cmdNewJournal = New System.Windows.Forms.Button
		Me.txtMiscID = New System.Windows.Forms.TextBox
		Me.txtUnpublishedID = New System.Windows.Forms.TextBox
		Me.txtLegislativeID = New System.Windows.Forms.TextBox
		Me.txtTreatiseID = New System.Windows.Forms.TextBox
		Me.txtChapterID = New System.Windows.Forms.TextBox
		Me.txtArticleID = New System.Windows.Forms.TextBox
		Me.cmbRecordNumber = New System.Windows.Forms.ComboBox
		Me.cmdNextRecord = New System.Windows.Forms.Button
		Me.cmdPreviousRecord = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.lstKeywords = New System.Windows.Forms.ListBox
		Me.lstCurrentKeywords = New System.Windows.Forms.ListBox
		Me.cmbAETChoice = New System.Windows.Forms.ComboBox
		Me.txtSuDocNumber = New System.Windows.Forms.TextBox
		Me.txtLargerWorkID = New System.Windows.Forms.TextBox
		Me.cmbLargerWorkTitle = New System.Windows.Forms.ComboBox
		Me.txtReportOrDocumentNumber = New System.Windows.Forms.TextBox
		Me.txtUSCCANCitation = New System.Windows.Forms.TextBox
		Me.txtStateLegislativeSession = New System.Windows.Forms.TextBox
		Me.txtSessionOfCongress = New System.Windows.Forms.TextBox
		Me.txtNumberOfCongress = New System.Windows.Forms.TextBox
		Me.txtLegislativeHouse = New System.Windows.Forms.TextBox
		Me.cmbLegislativeType = New System.Windows.Forms.ComboBox
		Me.cmbMiscType = New System.Windows.Forms.ComboBox
		Me.txtLocation = New System.Windows.Forms.TextBox
		Me.cmbUnpublishedType = New System.Windows.Forms.ComboBox
		Me.cmbThesisDissertationType = New System.Windows.Forms.ComboBox
		Me.chkAllChaptersBySameAuthor = New System.Windows.Forms.CheckBox
		Me.txtTitleOfSeriesIfNotIssuedByAuthor = New System.Windows.Forms.TextBox
		Me.txtSeriesVolume = New System.Windows.Forms.TextBox
		Me.txtOriginalPublicationDate = New System.Windows.Forms.TextBox
		Me.txtPublisher = New System.Windows.Forms.TextBox
		Me.txtEditionAndPrinting = New System.Windows.Forms.TextBox
		Me.txtOrganizationIssuingNewsletter = New System.Windows.Forms.TextBox
		Me.txtNotes = New System.Windows.Forms.TextBox
		Me.txtPage = New System.Windows.Forms.TextBox
		Me.cmbPublicationMonthOrSeason = New System.Windows.Forms.ComboBox
		Me.txtVolume = New System.Windows.Forms.TextBox
		Me.txtPublicationDay = New System.Windows.Forms.TextBox
		Me.txtJournalID = New System.Windows.Forms.TextBox
		Me.cmbJournalTitle = New System.Windows.Forms.ComboBox
		Me.cmbArticleDesignation = New System.Windows.Forms.ComboBox
		Me.txtTitle = New System.Windows.Forms.TextBox
		Me.txtYear = New System.Windows.Forms.TextBox
		Me.txtInputInitials = New System.Windows.Forms.TextBox
		Me.txtDateUpdated = New System.Windows.Forms.TextBox
		Me.txtDateAdded = New System.Windows.Forms.TextBox
		Me.cmbSourceType = New System.Windows.Forms.ComboBox
		Me.lstAuthors = New System.Windows.Forms.ListBox
		Me.lstTranslators = New System.Windows.Forms.ListBox
		Me.lstCurrentAuthors = New System.Windows.Forms.ListBox
		Me.lstCurrentTranslators = New System.Windows.Forms.ListBox
		Me.lstCurrentEditors = New System.Windows.Forms.ListBox
		Me.lstEditors = New System.Windows.Forms.ListBox
		Me.lblT = New System.Windows.Forms.TextBox
		Me.lblE = New System.Windows.Forms.TextBox
		Me.lblA = New System.Windows.Forms.TextBox
		Me.lblAETChoice = New System.Windows.Forms.TextBox
		Me.chkKeepSelected = New System.Windows.Forms.CheckBox
		Me.lblDoubleClickToAdd = New System.Windows.Forms.TextBox
		Me.lblArrow = New System.Windows.Forms.TextBox
		Me.lblTitle = New System.Windows.Forms.TextBox
		Me.lblYear = New System.Windows.Forms.TextBox
		Me.frmEntryInfo = New System.Windows.Forms.GroupBox
		Me.frmRecordInfo = New System.Windows.Forms.GroupBox
		Me.frmCitationInfo = New System.Windows.Forms.GroupBox
		Me.frmAuthorInfo = New System.Windows.Forms.GroupBox
		Me.frmKeywordInfo = New System.Windows.Forms.GroupBox
		Me.frmNotes = New System.Windows.Forms.GroupBox
		Me.lblSeparateBottom = New System.Windows.Forms.Label
		Me.lblMiscID = New System.Windows.Forms.Label
		Me.lblTreatiseID = New System.Windows.Forms.Label
		Me.lblUnpublishedID = New System.Windows.Forms.Label
		Me.lblChapterID = New System.Windows.Forms.Label
		Me.lblArticleID = New System.Windows.Forms.Label
		Me.lblLegisID = New System.Windows.Forms.Label
		Me.tglNewRecords = New AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
		Me.tglUpdateRecords = New AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
		Me.tglImportRecords = New AxMicrosoft.Vbe.Interop.Forms.AxToggleButton
		Me.lblLargerWorkID = New System.Windows.Forms.Label
		Me.lblNotes = New System.Windows.Forms.Label
		Me.lblSeparateTop = New System.Windows.Forms.Label
		Me.mneNewAuthor = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuAdd = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuFile = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuNewJournal = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuNewKeyword = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.tglNewRecords, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tglUpdateRecords, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tglImportRecords, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mneNewAuthor, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuAdd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuFile, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuNewJournal, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuNewKeyword, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Data Input and Editing"
		Me.ClientSize = New System.Drawing.Size(996, 875)
		Me.Location = New System.Drawing.Point(10, 48)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.Name = "frmMain"
		Me._mnuFile_1.Name = "_mnuFile_1"
		Me._mnuFile_1.Text = "File"
		Me._mnuFile_1.Checked = False
		Me._mnuFile_1.Enabled = True
		Me._mnuFile_1.Visible = True
		Me._mnuAdd_2.Name = "_mnuAdd_2"
		Me._mnuAdd_2.Text = "Add"
		Me._mnuAdd_2.Checked = False
		Me._mnuAdd_2.Enabled = True
		Me._mnuAdd_2.Visible = True
		Me._mneNewAuthor_3.Name = "_mneNewAuthor_3"
		Me._mneNewAuthor_3.Text = "New Author"
		Me._mneNewAuthor_3.Checked = False
		Me._mneNewAuthor_3.Enabled = True
		Me._mneNewAuthor_3.Visible = True
		Me._mnuNewJournal_4.Name = "_mnuNewJournal_4"
		Me._mnuNewJournal_4.Text = "New Journal"
		Me._mnuNewJournal_4.Checked = False
		Me._mnuNewJournal_4.Enabled = True
		Me._mnuNewJournal_4.Visible = True
		Me._mnuNewKeyword_5.Name = "_mnuNewKeyword_5"
		Me._mnuNewKeyword_5.Text = "New Keyword"
		Me._mnuNewKeyword_5.Checked = False
		Me._mnuNewKeyword_5.Enabled = True
		Me._mnuNewKeyword_5.Visible = True
		Me.chkRepublished.Text = "Republished?"
		Me.chkRepublished.Size = New System.Drawing.Size(89, 17)
		Me.chkRepublished.Location = New System.Drawing.Point(416, 112)
		Me.chkRepublished.TabIndex = 131
		Me.chkRepublished.TabStop = False
		Me.chkRepublished.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRepublished.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRepublished.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRepublished.BackColor = System.Drawing.SystemColors.Control
		Me.chkRepublished.CausesValidation = True
		Me.chkRepublished.Enabled = True
		Me.chkRepublished.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRepublished.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRepublished.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRepublished.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRepublished.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRepublished.Visible = True
		Me.chkRepublished.Name = "chkRepublished"
		Me.txtJournaTitleShortForm.AutoSize = False
		Me.txtJournaTitleShortForm.Enabled = False
		Me.txtJournaTitleShortForm.Size = New System.Drawing.Size(81, 19)
		Me.txtJournaTitleShortForm.Location = New System.Drawing.Point(864, 752)
		Me.txtJournaTitleShortForm.TabIndex = 130
		Me.txtJournaTitleShortForm.Visible = False
		Me.txtJournaTitleShortForm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtJournaTitleShortForm.AcceptsReturn = True
		Me.txtJournaTitleShortForm.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtJournaTitleShortForm.BackColor = System.Drawing.SystemColors.Window
		Me.txtJournaTitleShortForm.CausesValidation = True
		Me.txtJournaTitleShortForm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtJournaTitleShortForm.HideSelection = True
		Me.txtJournaTitleShortForm.ReadOnly = False
		Me.txtJournaTitleShortForm.Maxlength = 0
		Me.txtJournaTitleShortForm.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtJournaTitleShortForm.MultiLine = False
		Me.txtJournaTitleShortForm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtJournaTitleShortForm.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtJournaTitleShortForm.TabStop = True
		Me.txtJournaTitleShortForm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtJournaTitleShortForm.Name = "txtJournaTitleShortForm"
		Me.cmdPreview.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPreview.Text = "Preview Citation Form"
		Me.cmdPreview.Size = New System.Drawing.Size(185, 17)
		Me.cmdPreview.Location = New System.Drawing.Point(368, 704)
		Me.cmdPreview.TabIndex = 129
		Me.cmdPreview.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPreview.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPreview.CausesValidation = True
		Me.cmdPreview.Enabled = True
		Me.cmdPreview.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPreview.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPreview.TabStop = True
		Me.cmdPreview.Name = "cmdPreview"
		Me.lblRecordNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.lblRecordNumber.Text = "Record Number"
		Me.lblRecordNumber.Size = New System.Drawing.Size(89, 17)
		Me.lblRecordNumber.Location = New System.Drawing.Point(32, 88)
		Me.lblRecordNumber.TabIndex = 128
		Me.lblRecordNumber.TabStop = False
		Me.lblRecordNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRecordNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblRecordNumber.CausesValidation = True
		Me.lblRecordNumber.Enabled = True
		Me.lblRecordNumber.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblRecordNumber.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRecordNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRecordNumber.Name = "lblRecordNumber"
		Me.chkLibraryCollection.Text = "In Library Collection?"
		Me.chkLibraryCollection.Size = New System.Drawing.Size(89, 33)
		Me.chkLibraryCollection.Location = New System.Drawing.Point(416, 80)
		Me.chkLibraryCollection.TabIndex = 127
		Me.chkLibraryCollection.TabStop = False
		Me.chkLibraryCollection.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLibraryCollection.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLibraryCollection.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLibraryCollection.BackColor = System.Drawing.SystemColors.Control
		Me.chkLibraryCollection.CausesValidation = True
		Me.chkLibraryCollection.Enabled = True
		Me.chkLibraryCollection.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLibraryCollection.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLibraryCollection.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLibraryCollection.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLibraryCollection.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLibraryCollection.Visible = True
		Me.chkLibraryCollection.Name = "chkLibraryCollection"
		Me.lblArrow2.AutoSize = False
		Me.lblArrow2.BackColor = System.Drawing.SystemColors.Control
		Me.lblArrow2.Enabled = False
		Me.lblArrow2.Size = New System.Drawing.Size(65, 13)
		Me.lblArrow2.Location = New System.Drawing.Point(312, 480)
		Me.lblArrow2.TabIndex = 126
		Me.lblArrow2.Text = "<<<--------->>>"
		Me.lblArrow2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblArrow2.AcceptsReturn = True
		Me.lblArrow2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblArrow2.CausesValidation = True
		Me.lblArrow2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblArrow2.HideSelection = True
		Me.lblArrow2.ReadOnly = False
		Me.lblArrow2.Maxlength = 0
		Me.lblArrow2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblArrow2.MultiLine = False
		Me.lblArrow2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblArrow2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblArrow2.TabStop = True
		Me.lblArrow2.Visible = True
		Me.lblArrow2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblArrow2.Name = "lblArrow2"
		Me.lblDblClicktoAdd2.AutoSize = False
		Me.lblDblClicktoAdd2.BackColor = System.Drawing.SystemColors.Control
		Me.lblDblClicktoAdd2.Enabled = False
		Me.lblDblClicktoAdd2.Size = New System.Drawing.Size(65, 45)
		Me.lblDblClicktoAdd2.Location = New System.Drawing.Point(312, 384)
		Me.lblDblClicktoAdd2.MultiLine = True
		Me.lblDblClicktoAdd2.TabIndex = 125
		Me.lblDblClicktoAdd2.Text = "Double-Click to Add or Remove"
		Me.lblDblClicktoAdd2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDblClicktoAdd2.AcceptsReturn = True
		Me.lblDblClicktoAdd2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblDblClicktoAdd2.CausesValidation = True
		Me.lblDblClicktoAdd2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblDblClicktoAdd2.HideSelection = True
		Me.lblDblClicktoAdd2.ReadOnly = False
		Me.lblDblClicktoAdd2.Maxlength = 0
		Me.lblDblClicktoAdd2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblDblClicktoAdd2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDblClicktoAdd2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblDblClicktoAdd2.TabStop = True
		Me.lblDblClicktoAdd2.Visible = True
		Me.lblDblClicktoAdd2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDblClicktoAdd2.Name = "lblDblClicktoAdd2"
		Me.lblStatus.AutoSize = False
		Me.lblStatus.BackColor = System.Drawing.SystemColors.Control
		Me.lblStatus.Enabled = False
		Me.lblStatus.Size = New System.Drawing.Size(73, 13)
		Me.lblStatus.Location = New System.Drawing.Point(8, 675)
		Me.lblStatus.TabIndex = 123
		Me.lblStatus.Text = "Record Status:"
		Me.lblStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStatus.AcceptsReturn = True
		Me.lblStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblStatus.CausesValidation = True
		Me.lblStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblStatus.HideSelection = True
		Me.lblStatus.ReadOnly = False
		Me.lblStatus.Maxlength = 0
		Me.lblStatus.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblStatus.MultiLine = False
		Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStatus.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblStatus.TabStop = True
		Me.lblStatus.Visible = True
		Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblStatus.Name = "lblStatus"
		Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDelete.Text = "X-Delete Record-X"
		Me.cmdDelete.Size = New System.Drawing.Size(97, 33)
		Me.cmdDelete.Location = New System.Drawing.Point(784, 680)
		Me.cmdDelete.TabIndex = 122
		Me.cmdDelete.TabStop = False
		Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDelete.CausesValidation = True
		Me.cmdDelete.Enabled = True
		Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDelete.Name = "cmdDelete"
		Me.cmdEditJournal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEditJournal.Text = "Edit This Journal"
		Me.cmdEditJournal.Enabled = False
		Me.cmdEditJournal.Size = New System.Drawing.Size(97, 21)
		Me.cmdEditJournal.Location = New System.Drawing.Point(448, 168)
		Me.cmdEditJournal.TabIndex = 121
		Me.cmdEditJournal.TabStop = False
		Me.cmdEditJournal.Visible = False
		Me.cmdEditJournal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEditJournal.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEditJournal.CausesValidation = True
		Me.cmdEditJournal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEditJournal.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEditJournal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEditJournal.Name = "cmdEditJournal"
		Me.txtCallNumber.AutoSize = False
		Me.txtCallNumber.Enabled = False
		Me.txtCallNumber.Size = New System.Drawing.Size(121, 19)
		Me.txtCallNumber.Location = New System.Drawing.Point(752, 824)
		Me.txtCallNumber.TabIndex = 93
		Me.txtCallNumber.Visible = False
		Me.txtCallNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCallNumber.AcceptsReturn = True
		Me.txtCallNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCallNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtCallNumber.CausesValidation = True
		Me.txtCallNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCallNumber.HideSelection = True
		Me.txtCallNumber.ReadOnly = False
		Me.txtCallNumber.Maxlength = 0
		Me.txtCallNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCallNumber.MultiLine = False
		Me.txtCallNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCallNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCallNumber.TabStop = True
		Me.txtCallNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCallNumber.Name = "txtCallNumber"
		Me.cmdNewLargerWork.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNewLargerWork.Text = "New Larger Work"
		Me.cmdNewLargerWork.Size = New System.Drawing.Size(129, 21)
		Me.cmdNewLargerWork.Location = New System.Drawing.Point(16, 800)
		Me.cmdNewLargerWork.TabIndex = 92
		Me.cmdNewLargerWork.TabStop = False
		Me.cmdNewLargerWork.Visible = False
		Me.cmdNewLargerWork.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNewLargerWork.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNewLargerWork.CausesValidation = True
		Me.cmdNewLargerWork.Enabled = True
		Me.cmdNewLargerWork.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNewLargerWork.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNewLargerWork.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNewLargerWork.Name = "cmdNewLargerWork"
		Me.lblOriginalPublicationDate.AutoSize = False
		Me.lblOriginalPublicationDate.BackColor = System.Drawing.SystemColors.Control
		Me.lblOriginalPublicationDate.Enabled = False
		Me.lblOriginalPublicationDate.Size = New System.Drawing.Size(121, 13)
		Me.lblOriginalPublicationDate.Location = New System.Drawing.Point(512, 744)
		Me.lblOriginalPublicationDate.TabIndex = 91
		Me.lblOriginalPublicationDate.Text = "Original Publication Date"
		Me.lblOriginalPublicationDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOriginalPublicationDate.AcceptsReturn = True
		Me.lblOriginalPublicationDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblOriginalPublicationDate.CausesValidation = True
		Me.lblOriginalPublicationDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblOriginalPublicationDate.HideSelection = True
		Me.lblOriginalPublicationDate.ReadOnly = False
		Me.lblOriginalPublicationDate.Maxlength = 0
		Me.lblOriginalPublicationDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblOriginalPublicationDate.MultiLine = False
		Me.lblOriginalPublicationDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOriginalPublicationDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblOriginalPublicationDate.TabStop = True
		Me.lblOriginalPublicationDate.Visible = True
		Me.lblOriginalPublicationDate.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOriginalPublicationDate.Name = "lblOriginalPublicationDate"
		Me.lblPublisher.AutoSize = False
		Me.lblPublisher.BackColor = System.Drawing.SystemColors.Control
		Me.lblPublisher.Enabled = False
		Me.lblPublisher.Size = New System.Drawing.Size(89, 13)
		Me.lblPublisher.Location = New System.Drawing.Point(664, 744)
		Me.lblPublisher.TabIndex = 90
		Me.lblPublisher.Text = "Publisher"
		Me.lblPublisher.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPublisher.AcceptsReturn = True
		Me.lblPublisher.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblPublisher.CausesValidation = True
		Me.lblPublisher.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblPublisher.HideSelection = True
		Me.lblPublisher.ReadOnly = False
		Me.lblPublisher.Maxlength = 0
		Me.lblPublisher.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblPublisher.MultiLine = False
		Me.lblPublisher.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPublisher.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblPublisher.TabStop = True
		Me.lblPublisher.Visible = True
		Me.lblPublisher.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPublisher.Name = "lblPublisher"
		Me.lblCallNumber.AutoSize = False
		Me.lblCallNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblCallNumber.Enabled = False
		Me.lblCallNumber.Size = New System.Drawing.Size(81, 13)
		Me.lblCallNumber.Location = New System.Drawing.Point(864, 792)
		Me.lblCallNumber.TabIndex = 89
		Me.lblCallNumber.Text = "Call Number"
		Me.lblCallNumber.Visible = False
		Me.lblCallNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCallNumber.AcceptsReturn = True
		Me.lblCallNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblCallNumber.CausesValidation = True
		Me.lblCallNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblCallNumber.HideSelection = True
		Me.lblCallNumber.ReadOnly = False
		Me.lblCallNumber.Maxlength = 0
		Me.lblCallNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblCallNumber.MultiLine = False
		Me.lblCallNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCallNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblCallNumber.TabStop = True
		Me.lblCallNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCallNumber.Name = "lblCallNumber"
		Me.lblEditionAndPrinting.AutoSize = False
		Me.lblEditionAndPrinting.BackColor = System.Drawing.SystemColors.Control
		Me.lblEditionAndPrinting.Enabled = False
		Me.lblEditionAndPrinting.Size = New System.Drawing.Size(97, 13)
		Me.lblEditionAndPrinting.Location = New System.Drawing.Point(368, 744)
		Me.lblEditionAndPrinting.TabIndex = 88
		Me.lblEditionAndPrinting.Text = "Edition And Printing"
		Me.lblEditionAndPrinting.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEditionAndPrinting.AcceptsReturn = True
		Me.lblEditionAndPrinting.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblEditionAndPrinting.CausesValidation = True
		Me.lblEditionAndPrinting.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblEditionAndPrinting.HideSelection = True
		Me.lblEditionAndPrinting.ReadOnly = False
		Me.lblEditionAndPrinting.Maxlength = 0
		Me.lblEditionAndPrinting.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblEditionAndPrinting.MultiLine = False
		Me.lblEditionAndPrinting.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEditionAndPrinting.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblEditionAndPrinting.TabStop = True
		Me.lblEditionAndPrinting.Visible = True
		Me.lblEditionAndPrinting.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblEditionAndPrinting.Name = "lblEditionAndPrinting"
		Me.lblMiscType.AutoSize = False
		Me.lblMiscType.BackColor = System.Drawing.SystemColors.Control
		Me.lblMiscType.Enabled = False
		Me.lblMiscType.Size = New System.Drawing.Size(105, 13)
		Me.lblMiscType.Location = New System.Drawing.Point(744, 768)
		Me.lblMiscType.TabIndex = 87
		Me.lblMiscType.Text = "Miscellaneous Type"
		Me.lblMiscType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMiscType.AcceptsReturn = True
		Me.lblMiscType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblMiscType.CausesValidation = True
		Me.lblMiscType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblMiscType.HideSelection = True
		Me.lblMiscType.ReadOnly = False
		Me.lblMiscType.Maxlength = 0
		Me.lblMiscType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblMiscType.MultiLine = False
		Me.lblMiscType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMiscType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblMiscType.TabStop = True
		Me.lblMiscType.Visible = True
		Me.lblMiscType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMiscType.Name = "lblMiscType"
		Me.lblLocation.AutoSize = False
		Me.lblLocation.BackColor = System.Drawing.SystemColors.Control
		Me.lblLocation.Enabled = False
		Me.lblLocation.Size = New System.Drawing.Size(81, 13)
		Me.lblLocation.Location = New System.Drawing.Point(752, 816)
		Me.lblLocation.TabIndex = 86
		Me.lblLocation.Text = "Location"
		Me.lblLocation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLocation.AcceptsReturn = True
		Me.lblLocation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblLocation.CausesValidation = True
		Me.lblLocation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblLocation.HideSelection = True
		Me.lblLocation.ReadOnly = False
		Me.lblLocation.Maxlength = 0
		Me.lblLocation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblLocation.MultiLine = False
		Me.lblLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLocation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblLocation.TabStop = True
		Me.lblLocation.Visible = True
		Me.lblLocation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLocation.Name = "lblLocation"
		Me.lblThesisDissertationType.AutoSize = False
		Me.lblThesisDissertationType.BackColor = System.Drawing.SystemColors.Control
		Me.lblThesisDissertationType.Enabled = False
		Me.lblThesisDissertationType.Size = New System.Drawing.Size(121, 13)
		Me.lblThesisDissertationType.Location = New System.Drawing.Point(568, 768)
		Me.lblThesisDissertationType.TabIndex = 85
		Me.lblThesisDissertationType.Text = "Thesis/Dissertation Type"
		Me.lblThesisDissertationType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblThesisDissertationType.AcceptsReturn = True
		Me.lblThesisDissertationType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblThesisDissertationType.CausesValidation = True
		Me.lblThesisDissertationType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblThesisDissertationType.HideSelection = True
		Me.lblThesisDissertationType.ReadOnly = False
		Me.lblThesisDissertationType.Maxlength = 0
		Me.lblThesisDissertationType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblThesisDissertationType.MultiLine = False
		Me.lblThesisDissertationType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblThesisDissertationType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblThesisDissertationType.TabStop = True
		Me.lblThesisDissertationType.Visible = True
		Me.lblThesisDissertationType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblThesisDissertationType.Name = "lblThesisDissertationType"
		Me.lblUnpublishedType.AutoSize = False
		Me.lblUnpublishedType.BackColor = System.Drawing.SystemColors.Control
		Me.lblUnpublishedType.Enabled = False
		Me.lblUnpublishedType.Size = New System.Drawing.Size(105, 13)
		Me.lblUnpublishedType.Location = New System.Drawing.Point(368, 768)
		Me.lblUnpublishedType.TabIndex = 84
		Me.lblUnpublishedType.Text = "Unpublished Type"
		Me.lblUnpublishedType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUnpublishedType.AcceptsReturn = True
		Me.lblUnpublishedType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblUnpublishedType.CausesValidation = True
		Me.lblUnpublishedType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblUnpublishedType.HideSelection = True
		Me.lblUnpublishedType.ReadOnly = False
		Me.lblUnpublishedType.Maxlength = 0
		Me.lblUnpublishedType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblUnpublishedType.MultiLine = False
		Me.lblUnpublishedType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUnpublishedType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblUnpublishedType.TabStop = True
		Me.lblUnpublishedType.Visible = True
		Me.lblUnpublishedType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUnpublishedType.Name = "lblUnpublishedType"
		Me.lblUSCCANCitation.AutoSize = False
		Me.lblUSCCANCitation.BackColor = System.Drawing.SystemColors.Control
		Me.lblUSCCANCitation.Enabled = False
		Me.lblUSCCANCitation.Size = New System.Drawing.Size(97, 13)
		Me.lblUSCCANCitation.Location = New System.Drawing.Point(784, 792)
		Me.lblUSCCANCitation.TabIndex = 83
		Me.lblUSCCANCitation.Text = "USCCAN Citation"
		Me.lblUSCCANCitation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUSCCANCitation.AcceptsReturn = True
		Me.lblUSCCANCitation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblUSCCANCitation.CausesValidation = True
		Me.lblUSCCANCitation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblUSCCANCitation.HideSelection = True
		Me.lblUSCCANCitation.ReadOnly = False
		Me.lblUSCCANCitation.Maxlength = 0
		Me.lblUSCCANCitation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblUSCCANCitation.MultiLine = False
		Me.lblUSCCANCitation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUSCCANCitation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblUSCCANCitation.TabStop = True
		Me.lblUSCCANCitation.Visible = True
		Me.lblUSCCANCitation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUSCCANCitation.Name = "lblUSCCANCitation"
		Me.lblReportOrDocumentNumber.AutoSize = False
		Me.lblReportOrDocumentNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblReportOrDocumentNumber.Enabled = False
		Me.lblReportOrDocumentNumber.Size = New System.Drawing.Size(145, 13)
		Me.lblReportOrDocumentNumber.Location = New System.Drawing.Point(432, 792)
		Me.lblReportOrDocumentNumber.TabIndex = 82
		Me.lblReportOrDocumentNumber.Text = "Report or Document Number"
		Me.lblReportOrDocumentNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblReportOrDocumentNumber.AcceptsReturn = True
		Me.lblReportOrDocumentNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblReportOrDocumentNumber.CausesValidation = True
		Me.lblReportOrDocumentNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblReportOrDocumentNumber.HideSelection = True
		Me.lblReportOrDocumentNumber.ReadOnly = False
		Me.lblReportOrDocumentNumber.Maxlength = 0
		Me.lblReportOrDocumentNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblReportOrDocumentNumber.MultiLine = False
		Me.lblReportOrDocumentNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblReportOrDocumentNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblReportOrDocumentNumber.TabStop = True
		Me.lblReportOrDocumentNumber.Visible = True
		Me.lblReportOrDocumentNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblReportOrDocumentNumber.Name = "lblReportOrDocumentNumber"
		Me.lblLegislativeHouse.AutoSize = False
		Me.lblLegislativeHouse.BackColor = System.Drawing.SystemColors.Control
		Me.lblLegislativeHouse.Enabled = False
		Me.lblLegislativeHouse.Size = New System.Drawing.Size(89, 13)
		Me.lblLegislativeHouse.Location = New System.Drawing.Point(560, 736)
		Me.lblLegislativeHouse.TabIndex = 81
		Me.lblLegislativeHouse.Text = "Legislative House"
		Me.lblLegislativeHouse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLegislativeHouse.AcceptsReturn = True
		Me.lblLegislativeHouse.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblLegislativeHouse.CausesValidation = True
		Me.lblLegislativeHouse.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblLegislativeHouse.HideSelection = True
		Me.lblLegislativeHouse.ReadOnly = False
		Me.lblLegislativeHouse.Maxlength = 0
		Me.lblLegislativeHouse.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblLegislativeHouse.MultiLine = False
		Me.lblLegislativeHouse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLegislativeHouse.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblLegislativeHouse.TabStop = True
		Me.lblLegislativeHouse.Visible = True
		Me.lblLegislativeHouse.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLegislativeHouse.Name = "lblLegislativeHouse"
		Me.lblNumberOfCongress.AutoSize = False
		Me.lblNumberOfCongress.BackColor = System.Drawing.SystemColors.Control
		Me.lblNumberOfCongress.Enabled = False
		Me.lblNumberOfCongress.Size = New System.Drawing.Size(105, 13)
		Me.lblNumberOfCongress.Location = New System.Drawing.Point(128, 792)
		Me.lblNumberOfCongress.TabIndex = 80
		Me.lblNumberOfCongress.Text = "Number of Congress"
		Me.lblNumberOfCongress.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblNumberOfCongress.AcceptsReturn = True
		Me.lblNumberOfCongress.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblNumberOfCongress.CausesValidation = True
		Me.lblNumberOfCongress.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblNumberOfCongress.HideSelection = True
		Me.lblNumberOfCongress.ReadOnly = False
		Me.lblNumberOfCongress.Maxlength = 0
		Me.lblNumberOfCongress.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblNumberOfCongress.MultiLine = False
		Me.lblNumberOfCongress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNumberOfCongress.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblNumberOfCongress.TabStop = True
		Me.lblNumberOfCongress.Visible = True
		Me.lblNumberOfCongress.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblNumberOfCongress.Name = "lblNumberOfCongress"
		Me.lblSessionOfCongress.AutoSize = False
		Me.lblSessionOfCongress.BackColor = System.Drawing.SystemColors.Control
		Me.lblSessionOfCongress.Enabled = False
		Me.lblSessionOfCongress.Size = New System.Drawing.Size(105, 13)
		Me.lblSessionOfCongress.Location = New System.Drawing.Point(272, 792)
		Me.lblSessionOfCongress.TabIndex = 79
		Me.lblSessionOfCongress.Text = "Session of Congress"
		Me.lblSessionOfCongress.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSessionOfCongress.AcceptsReturn = True
		Me.lblSessionOfCongress.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblSessionOfCongress.CausesValidation = True
		Me.lblSessionOfCongress.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblSessionOfCongress.HideSelection = True
		Me.lblSessionOfCongress.ReadOnly = False
		Me.lblSessionOfCongress.Maxlength = 0
		Me.lblSessionOfCongress.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblSessionOfCongress.MultiLine = False
		Me.lblSessionOfCongress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSessionOfCongress.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblSessionOfCongress.TabStop = True
		Me.lblSessionOfCongress.Visible = True
		Me.lblSessionOfCongress.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSessionOfCongress.Name = "lblSessionOfCongress"
		Me.lblStateLegislativeSession.AutoSize = False
		Me.lblStateLegislativeSession.BackColor = System.Drawing.SystemColors.Control
		Me.lblStateLegislativeSession.Enabled = False
		Me.lblStateLegislativeSession.Size = New System.Drawing.Size(121, 13)
		Me.lblStateLegislativeSession.Location = New System.Drawing.Point(768, 736)
		Me.lblStateLegislativeSession.TabIndex = 78
		Me.lblStateLegislativeSession.Text = "State Legislative Session"
		Me.lblStateLegislativeSession.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStateLegislativeSession.AcceptsReturn = True
		Me.lblStateLegislativeSession.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblStateLegislativeSession.CausesValidation = True
		Me.lblStateLegislativeSession.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblStateLegislativeSession.HideSelection = True
		Me.lblStateLegislativeSession.ReadOnly = False
		Me.lblStateLegislativeSession.Maxlength = 0
		Me.lblStateLegislativeSession.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblStateLegislativeSession.MultiLine = False
		Me.lblStateLegislativeSession.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStateLegislativeSession.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblStateLegislativeSession.TabStop = True
		Me.lblStateLegislativeSession.Visible = True
		Me.lblStateLegislativeSession.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblStateLegislativeSession.Name = "lblStateLegislativeSession"
		Me.lblSuDocNumber.AutoSize = False
		Me.lblSuDocNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblSuDocNumber.Enabled = False
		Me.lblSuDocNumber.Size = New System.Drawing.Size(81, 13)
		Me.lblSuDocNumber.Location = New System.Drawing.Point(624, 792)
		Me.lblSuDocNumber.TabIndex = 77
		Me.lblSuDocNumber.Text = "SuDoc Number"
		Me.lblSuDocNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSuDocNumber.AcceptsReturn = True
		Me.lblSuDocNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblSuDocNumber.CausesValidation = True
		Me.lblSuDocNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblSuDocNumber.HideSelection = True
		Me.lblSuDocNumber.ReadOnly = False
		Me.lblSuDocNumber.Maxlength = 0
		Me.lblSuDocNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblSuDocNumber.MultiLine = False
		Me.lblSuDocNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSuDocNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblSuDocNumber.TabStop = True
		Me.lblSuDocNumber.Visible = True
		Me.lblSuDocNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSuDocNumber.Name = "lblSuDocNumber"
		Me.lblLegislativeType.AutoSize = False
		Me.lblLegislativeType.BackColor = System.Drawing.SystemColors.Control
		Me.lblLegislativeType.Enabled = False
		Me.lblLegislativeType.Size = New System.Drawing.Size(81, 13)
		Me.lblLegislativeType.Location = New System.Drawing.Point(128, 736)
		Me.lblLegislativeType.TabIndex = 76
		Me.lblLegislativeType.Text = "Legislative Type"
		Me.lblLegislativeType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLegislativeType.AcceptsReturn = True
		Me.lblLegislativeType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblLegislativeType.CausesValidation = True
		Me.lblLegislativeType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblLegislativeType.HideSelection = True
		Me.lblLegislativeType.ReadOnly = False
		Me.lblLegislativeType.Maxlength = 0
		Me.lblLegislativeType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblLegislativeType.MultiLine = False
		Me.lblLegislativeType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLegislativeType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblLegislativeType.TabStop = True
		Me.lblLegislativeType.Visible = True
		Me.lblLegislativeType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLegislativeType.Name = "lblLegislativeType"
		Me.lblSeriesVolume.AutoSize = False
		Me.lblSeriesVolume.BackColor = System.Drawing.SystemColors.Control
		Me.lblSeriesVolume.Enabled = False
		Me.lblSeriesVolume.Size = New System.Drawing.Size(81, 13)
		Me.lblSeriesVolume.Location = New System.Drawing.Point(616, 816)
		Me.lblSeriesVolume.TabIndex = 75
		Me.lblSeriesVolume.Text = "Series Volume"
		Me.lblSeriesVolume.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSeriesVolume.AcceptsReturn = True
		Me.lblSeriesVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblSeriesVolume.CausesValidation = True
		Me.lblSeriesVolume.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblSeriesVolume.HideSelection = True
		Me.lblSeriesVolume.ReadOnly = False
		Me.lblSeriesVolume.Maxlength = 0
		Me.lblSeriesVolume.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblSeriesVolume.MultiLine = False
		Me.lblSeriesVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSeriesVolume.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblSeriesVolume.TabStop = True
		Me.lblSeriesVolume.Visible = True
		Me.lblSeriesVolume.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSeriesVolume.Name = "lblSeriesVolume"
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.AutoSize = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Enabled = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Size = New System.Drawing.Size(185, 13)
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Location = New System.Drawing.Point(720, 816)
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.TabIndex = 74
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Text = "Title Of Series If Not Issued By Author"
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.AcceptsReturn = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.CausesValidation = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.HideSelection = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.ReadOnly = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Maxlength = 0
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.MultiLine = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.TabStop = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Visible = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Name = "lblTitleOfSeriesIfNotIssuedByAuthor"
		Me.lblLargerWorkTitle.AutoSize = False
		Me.lblLargerWorkTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblLargerWorkTitle.Enabled = False
		Me.lblLargerWorkTitle.Size = New System.Drawing.Size(89, 13)
		Me.lblLargerWorkTitle.Location = New System.Drawing.Point(616, 768)
		Me.lblLargerWorkTitle.TabIndex = 73
		Me.lblLargerWorkTitle.Text = "Larger Work Title"
		Me.lblLargerWorkTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLargerWorkTitle.AcceptsReturn = True
		Me.lblLargerWorkTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblLargerWorkTitle.CausesValidation = True
		Me.lblLargerWorkTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblLargerWorkTitle.HideSelection = True
		Me.lblLargerWorkTitle.ReadOnly = False
		Me.lblLargerWorkTitle.Maxlength = 0
		Me.lblLargerWorkTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblLargerWorkTitle.MultiLine = False
		Me.lblLargerWorkTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLargerWorkTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblLargerWorkTitle.TabStop = True
		Me.lblLargerWorkTitle.Visible = True
		Me.lblLargerWorkTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLargerWorkTitle.Name = "lblLargerWorkTitle"
		Me.cmbPagination.Enabled = False
		Me.cmbPagination.Size = New System.Drawing.Size(249, 21)
		Me.cmbPagination.Location = New System.Drawing.Point(368, 680)
		Me.cmbPagination.TabIndex = 72
		Me.cmbPagination.Visible = False
		Me.cmbPagination.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbPagination.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPagination.CausesValidation = True
		Me.cmbPagination.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPagination.IntegralHeight = True
		Me.cmbPagination.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPagination.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPagination.Sorted = False
		Me.cmbPagination.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbPagination.TabStop = True
		Me.cmbPagination.Name = "cmbPagination"
		Me.lblVolume.AutoSize = False
		Me.lblVolume.BackColor = System.Drawing.SystemColors.Control
		Me.lblVolume.Enabled = False
		Me.lblVolume.Size = New System.Drawing.Size(49, 13)
		Me.lblVolume.Location = New System.Drawing.Point(696, 264)
		Me.lblVolume.TabIndex = 108
		Me.lblVolume.Text = "Volume"
		Me.lblVolume.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVolume.AcceptsReturn = True
		Me.lblVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblVolume.CausesValidation = True
		Me.lblVolume.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblVolume.HideSelection = True
		Me.lblVolume.ReadOnly = False
		Me.lblVolume.Maxlength = 0
		Me.lblVolume.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblVolume.MultiLine = False
		Me.lblVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVolume.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblVolume.TabStop = True
		Me.lblVolume.Visible = True
		Me.lblVolume.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVolume.Name = "lblVolume"
		Me.lblPublicationMonthOrSeason.AutoSize = False
		Me.lblPublicationMonthOrSeason.BackColor = System.Drawing.SystemColors.Control
		Me.lblPublicationMonthOrSeason.Enabled = False
		Me.lblPublicationMonthOrSeason.Size = New System.Drawing.Size(137, 13)
		Me.lblPublicationMonthOrSeason.Location = New System.Drawing.Point(216, 264)
		Me.lblPublicationMonthOrSeason.TabIndex = 105
		Me.lblPublicationMonthOrSeason.Text = "Publication Month or Season"
		Me.lblPublicationMonthOrSeason.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPublicationMonthOrSeason.AcceptsReturn = True
		Me.lblPublicationMonthOrSeason.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblPublicationMonthOrSeason.CausesValidation = True
		Me.lblPublicationMonthOrSeason.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblPublicationMonthOrSeason.HideSelection = True
		Me.lblPublicationMonthOrSeason.ReadOnly = False
		Me.lblPublicationMonthOrSeason.Maxlength = 0
		Me.lblPublicationMonthOrSeason.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblPublicationMonthOrSeason.MultiLine = False
		Me.lblPublicationMonthOrSeason.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPublicationMonthOrSeason.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblPublicationMonthOrSeason.TabStop = True
		Me.lblPublicationMonthOrSeason.Visible = True
		Me.lblPublicationMonthOrSeason.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPublicationMonthOrSeason.Name = "lblPublicationMonthOrSeason"
		Me.lblPage.AutoSize = False
		Me.lblPage.BackColor = System.Drawing.SystemColors.Control
		Me.lblPage.Enabled = False
		Me.lblPage.Size = New System.Drawing.Size(81, 13)
		Me.lblPage.Location = New System.Drawing.Point(552, 264)
		Me.lblPage.TabIndex = 107
		Me.lblPage.Text = "Page Number"
		Me.lblPage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPage.AcceptsReturn = True
		Me.lblPage.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblPage.CausesValidation = True
		Me.lblPage.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblPage.HideSelection = True
		Me.lblPage.ReadOnly = False
		Me.lblPage.Maxlength = 0
		Me.lblPage.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblPage.MultiLine = False
		Me.lblPage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPage.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblPage.TabStop = True
		Me.lblPage.Visible = True
		Me.lblPage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPage.Name = "lblPage"
		Me.chkSource.Text = "Check to keep same type"
		Me.chkSource.Size = New System.Drawing.Size(145, 13)
		Me.chkSource.Location = New System.Drawing.Point(216, 88)
		Me.chkSource.TabIndex = 98
		Me.chkSource.TabStop = False
		Me.chkSource.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSource.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSource.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSource.BackColor = System.Drawing.SystemColors.Control
		Me.chkSource.CausesValidation = True
		Me.chkSource.Enabled = True
		Me.chkSource.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSource.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSource.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSource.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSource.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSource.Visible = True
		Me.chkSource.Name = "chkSource"
		Me.lblPublicationDay.AutoSize = False
		Me.lblPublicationDay.BackColor = System.Drawing.SystemColors.Control
		Me.lblPublicationDay.Enabled = False
		Me.lblPublicationDay.Size = New System.Drawing.Size(81, 13)
		Me.lblPublicationDay.Location = New System.Drawing.Point(416, 264)
		Me.lblPublicationDay.TabIndex = 106
		Me.lblPublicationDay.Text = "Publication Day"
		Me.lblPublicationDay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPublicationDay.AcceptsReturn = True
		Me.lblPublicationDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblPublicationDay.CausesValidation = True
		Me.lblPublicationDay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblPublicationDay.HideSelection = True
		Me.lblPublicationDay.ReadOnly = False
		Me.lblPublicationDay.Maxlength = 0
		Me.lblPublicationDay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblPublicationDay.MultiLine = False
		Me.lblPublicationDay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPublicationDay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblPublicationDay.TabStop = True
		Me.lblPublicationDay.Visible = True
		Me.lblPublicationDay.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPublicationDay.Name = "lblPublicationDay"
		Me.chkYear.Text = "Check to keep same year"
		Me.chkYear.Enabled = False
		Me.chkYear.Size = New System.Drawing.Size(145, 17)
		Me.chkYear.Location = New System.Drawing.Point(648, 184)
		Me.chkYear.TabIndex = 71
		Me.chkYear.TabStop = False
		Me.chkYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkYear.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkYear.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkYear.BackColor = System.Drawing.SystemColors.Control
		Me.chkYear.CausesValidation = True
		Me.chkYear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkYear.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkYear.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkYear.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkYear.Visible = True
		Me.chkYear.Name = "chkYear"
		Me.lblKeywords.AutoSize = False
		Me.lblKeywords.BackColor = System.Drawing.SystemColors.Control
		Me.lblKeywords.Enabled = False
		Me.lblKeywords.Size = New System.Drawing.Size(89, 13)
		Me.lblKeywords.Location = New System.Drawing.Point(32, 464)
		Me.lblKeywords.TabIndex = 70
		Me.lblKeywords.Text = "Select Keywords"
		Me.lblKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblKeywords.AcceptsReturn = True
		Me.lblKeywords.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblKeywords.CausesValidation = True
		Me.lblKeywords.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblKeywords.HideSelection = True
		Me.lblKeywords.ReadOnly = False
		Me.lblKeywords.Maxlength = 0
		Me.lblKeywords.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblKeywords.MultiLine = False
		Me.lblKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblKeywords.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblKeywords.TabStop = True
		Me.lblKeywords.Visible = True
		Me.lblKeywords.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblKeywords.Name = "lblKeywords"
		Me.lblJournalTitle.AutoSize = False
		Me.lblJournalTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblJournalTitle.Enabled = False
		Me.lblJournalTitle.Size = New System.Drawing.Size(65, 13)
		Me.lblJournalTitle.Location = New System.Drawing.Point(32, 168)
		Me.lblJournalTitle.TabIndex = 99
		Me.lblJournalTitle.Text = "Journal Title"
		Me.lblJournalTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblJournalTitle.AcceptsReturn = True
		Me.lblJournalTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblJournalTitle.CausesValidation = True
		Me.lblJournalTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblJournalTitle.HideSelection = True
		Me.lblJournalTitle.ReadOnly = False
		Me.lblJournalTitle.Maxlength = 0
		Me.lblJournalTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblJournalTitle.MultiLine = False
		Me.lblJournalTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblJournalTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblJournalTitle.TabStop = True
		Me.lblJournalTitle.Visible = True
		Me.lblJournalTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblJournalTitle.Name = "lblJournalTitle"
		Me.lblSourceType.AutoSize = False
		Me.lblSourceType.BackColor = System.Drawing.SystemColors.Control
		Me.lblSourceType.Enabled = False
		Me.lblSourceType.Size = New System.Drawing.Size(89, 13)
		Me.lblSourceType.Location = New System.Drawing.Point(144, 88)
		Me.lblSourceType.TabIndex = 94
		Me.lblSourceType.Text = "Source Type"
		Me.lblSourceType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSourceType.AcceptsReturn = True
		Me.lblSourceType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblSourceType.CausesValidation = True
		Me.lblSourceType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblSourceType.HideSelection = True
		Me.lblSourceType.ReadOnly = False
		Me.lblSourceType.Maxlength = 0
		Me.lblSourceType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblSourceType.MultiLine = False
		Me.lblSourceType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSourceType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblSourceType.TabStop = True
		Me.lblSourceType.Visible = True
		Me.lblSourceType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSourceType.Name = "lblSourceType"
		Me.lblArticleDesignation.AutoSize = False
		Me.lblArticleDesignation.BackColor = System.Drawing.SystemColors.Control
		Me.lblArticleDesignation.Enabled = False
		Me.lblArticleDesignation.Size = New System.Drawing.Size(97, 13)
		Me.lblArticleDesignation.Location = New System.Drawing.Point(32, 264)
		Me.lblArticleDesignation.TabIndex = 104
		Me.lblArticleDesignation.Text = "Article Designation"
		Me.lblArticleDesignation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblArticleDesignation.AcceptsReturn = True
		Me.lblArticleDesignation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblArticleDesignation.CausesValidation = True
		Me.lblArticleDesignation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblArticleDesignation.HideSelection = True
		Me.lblArticleDesignation.ReadOnly = False
		Me.lblArticleDesignation.Maxlength = 0
		Me.lblArticleDesignation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblArticleDesignation.MultiLine = False
		Me.lblArticleDesignation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblArticleDesignation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblArticleDesignation.TabStop = True
		Me.lblArticleDesignation.Visible = True
		Me.lblArticleDesignation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblArticleDesignation.Name = "lblArticleDesignation"
		Me.lblInputInitials.AutoSize = False
		Me.lblInputInitials.BackColor = System.Drawing.SystemColors.Control
		Me.lblInputInitials.Enabled = False
		Me.lblInputInitials.Size = New System.Drawing.Size(81, 13)
		Me.lblInputInitials.Location = New System.Drawing.Point(752, 88)
		Me.lblInputInitials.TabIndex = 97
		Me.lblInputInitials.Text = "Input Initials"
		Me.lblInputInitials.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInputInitials.AcceptsReturn = True
		Me.lblInputInitials.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblInputInitials.CausesValidation = True
		Me.lblInputInitials.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblInputInitials.HideSelection = True
		Me.lblInputInitials.ReadOnly = False
		Me.lblInputInitials.Maxlength = 0
		Me.lblInputInitials.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblInputInitials.MultiLine = False
		Me.lblInputInitials.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInputInitials.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblInputInitials.TabStop = True
		Me.lblInputInitials.Visible = True
		Me.lblInputInitials.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInputInitials.Name = "lblInputInitials"
		Me.lblDateUpdated.AutoSize = False
		Me.lblDateUpdated.BackColor = System.Drawing.SystemColors.Control
		Me.lblDateUpdated.Enabled = False
		Me.lblDateUpdated.Size = New System.Drawing.Size(81, 13)
		Me.lblDateUpdated.Location = New System.Drawing.Point(640, 88)
		Me.lblDateUpdated.TabIndex = 96
		Me.lblDateUpdated.Text = "Date Updated"
		Me.lblDateUpdated.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDateUpdated.AcceptsReturn = True
		Me.lblDateUpdated.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblDateUpdated.CausesValidation = True
		Me.lblDateUpdated.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblDateUpdated.HideSelection = True
		Me.lblDateUpdated.ReadOnly = False
		Me.lblDateUpdated.Maxlength = 0
		Me.lblDateUpdated.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblDateUpdated.MultiLine = False
		Me.lblDateUpdated.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDateUpdated.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblDateUpdated.TabStop = True
		Me.lblDateUpdated.Visible = True
		Me.lblDateUpdated.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDateUpdated.Name = "lblDateUpdated"
		Me.lblPublicationYear.AutoSize = False
		Me.lblPublicationYear.BackColor = System.Drawing.SystemColors.Control
		Me.lblPublicationYear.Enabled = False
		Me.lblPublicationYear.Size = New System.Drawing.Size(81, 13)
		Me.lblPublicationYear.Location = New System.Drawing.Point(536, 88)
		Me.lblPublicationYear.TabIndex = 95
		Me.lblPublicationYear.Text = "Date Added"
		Me.lblPublicationYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPublicationYear.AcceptsReturn = True
		Me.lblPublicationYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblPublicationYear.CausesValidation = True
		Me.lblPublicationYear.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblPublicationYear.HideSelection = True
		Me.lblPublicationYear.ReadOnly = False
		Me.lblPublicationYear.Maxlength = 0
		Me.lblPublicationYear.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblPublicationYear.MultiLine = False
		Me.lblPublicationYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPublicationYear.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblPublicationYear.TabStop = True
		Me.lblPublicationYear.Visible = True
		Me.lblPublicationYear.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPublicationYear.Name = "lblPublicationYear"
		Me.txtStatus.AutoSize = False
		Me.txtStatus.BackColor = System.Drawing.SystemColors.GrayText
		Me.txtStatus.Enabled = False
		Me.txtStatus.Size = New System.Drawing.Size(73, 19)
		Me.txtStatus.Location = New System.Drawing.Point(88, 672)
		Me.txtStatus.TabIndex = 69
		Me.txtStatus.Text = "Unchanged"
		Me.txtStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtStatus.AcceptsReturn = True
		Me.txtStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStatus.CausesValidation = True
		Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStatus.HideSelection = True
		Me.txtStatus.ReadOnly = False
		Me.txtStatus.Maxlength = 0
		Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStatus.MultiLine = False
		Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStatus.TabStop = True
		Me.txtStatus.Visible = True
		Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStatus.Name = "txtStatus"
		Me.lstNewKeywords.Size = New System.Drawing.Size(169, 59)
		Me.lstNewKeywords.Location = New System.Drawing.Point(672, 480)
		Me.lstNewKeywords.Sorted = True
		Me.lstNewKeywords.TabIndex = 116
		Me.lstNewKeywords.TabStop = False
		Me.lstNewKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstNewKeywords.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstNewKeywords.BackColor = System.Drawing.SystemColors.Window
		Me.lstNewKeywords.CausesValidation = True
		Me.lstNewKeywords.Enabled = True
		Me.lstNewKeywords.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstNewKeywords.IntegralHeight = True
		Me.lstNewKeywords.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstNewKeywords.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstNewKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstNewKeywords.Visible = True
		Me.lstNewKeywords.MultiColumn = False
		Me.lstNewKeywords.Name = "lstNewKeywords"
		Me.cmdGetNewKeywords.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGetNewKeywords.Text = "Suggest New Keywords"
		Me.cmdGetNewKeywords.Size = New System.Drawing.Size(169, 17)
		Me.cmdGetNewKeywords.Location = New System.Drawing.Point(672, 464)
		Me.cmdGetNewKeywords.TabIndex = 117
		Me.cmdGetNewKeywords.TabStop = False
		Me.cmdGetNewKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGetNewKeywords.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGetNewKeywords.CausesValidation = True
		Me.cmdGetNewKeywords.Enabled = True
		Me.cmdGetNewKeywords.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGetNewKeywords.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGetNewKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGetNewKeywords.Name = "cmdGetNewKeywords"
		Me.cmdNewAuthor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNewAuthor.Text = "New Author"
		Me.cmdNewAuthor.Size = New System.Drawing.Size(97, 17)
		Me.cmdNewAuthor.Location = New System.Drawing.Point(216, 352)
		Me.cmdNewAuthor.TabIndex = 110
		Me.cmdNewAuthor.TabStop = False
		Me.cmdNewAuthor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNewAuthor.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNewAuthor.CausesValidation = True
		Me.cmdNewAuthor.Enabled = True
		Me.cmdNewAuthor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNewAuthor.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNewAuthor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNewAuthor.Name = "cmdNewAuthor"
		Me.cmdNewJournal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNewJournal.Text = "New Journal"
		Me.cmdNewJournal.Enabled = False
		Me.cmdNewJournal.Size = New System.Drawing.Size(97, 21)
		Me.cmdNewJournal.Location = New System.Drawing.Point(448, 192)
		Me.cmdNewJournal.TabIndex = 102
		Me.cmdNewJournal.TabStop = False
		Me.cmdNewJournal.Visible = False
		Me.cmdNewJournal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNewJournal.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNewJournal.CausesValidation = True
		Me.cmdNewJournal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNewJournal.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNewJournal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNewJournal.Name = "cmdNewJournal"
		Me.txtMiscID.AutoSize = False
		Me.txtMiscID.Enabled = False
		Me.txtMiscID.Size = New System.Drawing.Size(73, 19)
		Me.txtMiscID.Location = New System.Drawing.Point(96, 744)
		Me.txtMiscID.TabIndex = 58
		Me.txtMiscID.Visible = False
		Me.txtMiscID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMiscID.AcceptsReturn = True
		Me.txtMiscID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMiscID.BackColor = System.Drawing.SystemColors.Window
		Me.txtMiscID.CausesValidation = True
		Me.txtMiscID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMiscID.HideSelection = True
		Me.txtMiscID.ReadOnly = False
		Me.txtMiscID.Maxlength = 0
		Me.txtMiscID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMiscID.MultiLine = False
		Me.txtMiscID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMiscID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMiscID.TabStop = True
		Me.txtMiscID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMiscID.Name = "txtMiscID"
		Me.txtUnpublishedID.AutoSize = False
		Me.txtUnpublishedID.Enabled = False
		Me.txtUnpublishedID.Size = New System.Drawing.Size(73, 19)
		Me.txtUnpublishedID.Location = New System.Drawing.Point(96, 712)
		Me.txtUnpublishedID.TabIndex = 57
		Me.txtUnpublishedID.Visible = False
		Me.txtUnpublishedID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUnpublishedID.AcceptsReturn = True
		Me.txtUnpublishedID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUnpublishedID.BackColor = System.Drawing.SystemColors.Window
		Me.txtUnpublishedID.CausesValidation = True
		Me.txtUnpublishedID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUnpublishedID.HideSelection = True
		Me.txtUnpublishedID.ReadOnly = False
		Me.txtUnpublishedID.Maxlength = 0
		Me.txtUnpublishedID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUnpublishedID.MultiLine = False
		Me.txtUnpublishedID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUnpublishedID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUnpublishedID.TabStop = True
		Me.txtUnpublishedID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUnpublishedID.Name = "txtUnpublishedID"
		Me.txtLegislativeID.AutoSize = False
		Me.txtLegislativeID.Enabled = False
		Me.txtLegislativeID.Size = New System.Drawing.Size(73, 19)
		Me.txtLegislativeID.Location = New System.Drawing.Point(240, 704)
		Me.txtLegislativeID.TabIndex = 56
		Me.txtLegislativeID.Visible = False
		Me.txtLegislativeID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLegislativeID.AcceptsReturn = True
		Me.txtLegislativeID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLegislativeID.BackColor = System.Drawing.SystemColors.Window
		Me.txtLegislativeID.CausesValidation = True
		Me.txtLegislativeID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLegislativeID.HideSelection = True
		Me.txtLegislativeID.ReadOnly = False
		Me.txtLegislativeID.Maxlength = 0
		Me.txtLegislativeID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLegislativeID.MultiLine = False
		Me.txtLegislativeID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLegislativeID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLegislativeID.TabStop = True
		Me.txtLegislativeID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLegislativeID.Name = "txtLegislativeID"
		Me.txtTreatiseID.AutoSize = False
		Me.txtTreatiseID.Enabled = False
		Me.txtTreatiseID.Size = New System.Drawing.Size(73, 19)
		Me.txtTreatiseID.Location = New System.Drawing.Point(96, 728)
		Me.txtTreatiseID.TabIndex = 55
		Me.txtTreatiseID.Visible = False
		Me.txtTreatiseID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTreatiseID.AcceptsReturn = True
		Me.txtTreatiseID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTreatiseID.BackColor = System.Drawing.SystemColors.Window
		Me.txtTreatiseID.CausesValidation = True
		Me.txtTreatiseID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTreatiseID.HideSelection = True
		Me.txtTreatiseID.ReadOnly = False
		Me.txtTreatiseID.Maxlength = 0
		Me.txtTreatiseID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTreatiseID.MultiLine = False
		Me.txtTreatiseID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTreatiseID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTreatiseID.TabStop = True
		Me.txtTreatiseID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTreatiseID.Name = "txtTreatiseID"
		Me.txtChapterID.AutoSize = False
		Me.txtChapterID.Enabled = False
		Me.txtChapterID.Size = New System.Drawing.Size(73, 19)
		Me.txtChapterID.Location = New System.Drawing.Point(96, 696)
		Me.txtChapterID.TabIndex = 54
		Me.txtChapterID.Visible = False
		Me.txtChapterID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChapterID.AcceptsReturn = True
		Me.txtChapterID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtChapterID.BackColor = System.Drawing.SystemColors.Window
		Me.txtChapterID.CausesValidation = True
		Me.txtChapterID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtChapterID.HideSelection = True
		Me.txtChapterID.ReadOnly = False
		Me.txtChapterID.Maxlength = 0
		Me.txtChapterID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChapterID.MultiLine = False
		Me.txtChapterID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChapterID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChapterID.TabStop = True
		Me.txtChapterID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChapterID.Name = "txtChapterID"
		Me.txtArticleID.AutoSize = False
		Me.txtArticleID.Enabled = False
		Me.txtArticleID.Size = New System.Drawing.Size(73, 19)
		Me.txtArticleID.Location = New System.Drawing.Point(240, 720)
		Me.txtArticleID.TabIndex = 53
		Me.txtArticleID.Visible = False
		Me.txtArticleID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtArticleID.AcceptsReturn = True
		Me.txtArticleID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtArticleID.BackColor = System.Drawing.SystemColors.Window
		Me.txtArticleID.CausesValidation = True
		Me.txtArticleID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtArticleID.HideSelection = True
		Me.txtArticleID.ReadOnly = False
		Me.txtArticleID.Maxlength = 0
		Me.txtArticleID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtArticleID.MultiLine = False
		Me.txtArticleID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtArticleID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtArticleID.TabStop = True
		Me.txtArticleID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtArticleID.Name = "txtArticleID"
		Me.cmbRecordNumber.Size = New System.Drawing.Size(89, 21)
		Me.cmbRecordNumber.Location = New System.Drawing.Point(32, 104)
		Me.cmbRecordNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbRecordNumber.TabIndex = 120
		Me.cmbRecordNumber.TabStop = False
		Me.cmbRecordNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbRecordNumber.BackColor = System.Drawing.SystemColors.Window
		Me.cmbRecordNumber.CausesValidation = True
		Me.cmbRecordNumber.Enabled = True
		Me.cmbRecordNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbRecordNumber.IntegralHeight = True
		Me.cmbRecordNumber.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbRecordNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbRecordNumber.Sorted = False
		Me.cmbRecordNumber.Visible = True
		Me.cmbRecordNumber.Name = "cmbRecordNumber"
		Me.cmdNextRecord.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNextRecord.Text = "-->"
		Me.cmdNextRecord.Size = New System.Drawing.Size(81, 33)
		Me.cmdNextRecord.Location = New System.Drawing.Point(520, 672)
		Me.cmdNextRecord.TabIndex = 51
		Me.cmdNextRecord.TabStop = False
		Me.cmdNextRecord.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNextRecord.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNextRecord.CausesValidation = True
		Me.cmdNextRecord.Enabled = True
		Me.cmdNextRecord.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNextRecord.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNextRecord.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNextRecord.Name = "cmdNextRecord"
		Me.cmdPreviousRecord.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPreviousRecord.Text = "<--"
		Me.cmdPreviousRecord.Size = New System.Drawing.Size(81, 33)
		Me.cmdPreviousRecord.Location = New System.Drawing.Point(328, 672)
		Me.cmdPreviousRecord.TabIndex = 50
		Me.cmdPreviousRecord.TabStop = False
		Me.cmdPreviousRecord.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPreviousRecord.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPreviousRecord.CausesValidation = True
		Me.cmdPreviousRecord.Enabled = True
		Me.cmdPreviousRecord.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPreviousRecord.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPreviousRecord.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPreviousRecord.Name = "cmdPreviousRecord"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Size = New System.Drawing.Size(81, 33)
		Me.cmdSave.Location = New System.Drawing.Point(424, 672)
		Me.cmdSave.TabIndex = 20
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.lstKeywords.Size = New System.Drawing.Size(281, 59)
		Me.lstKeywords.Location = New System.Drawing.Point(32, 480)
		Me.lstKeywords.Sorted = True
		Me.lstKeywords.TabIndex = 16
		Me.lstKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstKeywords.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstKeywords.BackColor = System.Drawing.SystemColors.Window
		Me.lstKeywords.CausesValidation = True
		Me.lstKeywords.Enabled = True
		Me.lstKeywords.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstKeywords.IntegralHeight = True
		Me.lstKeywords.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstKeywords.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstKeywords.TabStop = True
		Me.lstKeywords.Visible = True
		Me.lstKeywords.MultiColumn = False
		Me.lstKeywords.Name = "lstKeywords"
		Me.lstCurrentKeywords.Size = New System.Drawing.Size(281, 59)
		Me.lstCurrentKeywords.Location = New System.Drawing.Point(376, 480)
		Me.lstCurrentKeywords.TabIndex = 118
		Me.lstCurrentKeywords.TabStop = False
		Me.lstCurrentKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCurrentKeywords.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCurrentKeywords.BackColor = System.Drawing.SystemColors.Window
		Me.lstCurrentKeywords.CausesValidation = True
		Me.lstCurrentKeywords.Enabled = True
		Me.lstCurrentKeywords.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCurrentKeywords.IntegralHeight = True
		Me.lstCurrentKeywords.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCurrentKeywords.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCurrentKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCurrentKeywords.Sorted = False
		Me.lstCurrentKeywords.Visible = True
		Me.lstCurrentKeywords.MultiColumn = False
		Me.lstCurrentKeywords.Name = "lstCurrentKeywords"
		Me.cmbAETChoice.Size = New System.Drawing.Size(113, 21)
		Me.cmbAETChoice.Location = New System.Drawing.Point(72, 344)
		Me.cmbAETChoice.TabIndex = 109
		Me.cmbAETChoice.TabStop = False
		Me.cmbAETChoice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbAETChoice.BackColor = System.Drawing.SystemColors.Window
		Me.cmbAETChoice.CausesValidation = True
		Me.cmbAETChoice.Enabled = True
		Me.cmbAETChoice.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbAETChoice.IntegralHeight = True
		Me.cmbAETChoice.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbAETChoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbAETChoice.Sorted = False
		Me.cmbAETChoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbAETChoice.Visible = True
		Me.cmbAETChoice.Name = "cmbAETChoice"
		Me.txtSuDocNumber.AutoSize = False
		Me.txtSuDocNumber.Size = New System.Drawing.Size(113, 19)
		Me.txtSuDocNumber.Location = New System.Drawing.Point(864, 224)
		Me.txtSuDocNumber.TabIndex = 43
		Me.txtSuDocNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSuDocNumber.AcceptsReturn = True
		Me.txtSuDocNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSuDocNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtSuDocNumber.CausesValidation = True
		Me.txtSuDocNumber.Enabled = True
		Me.txtSuDocNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSuDocNumber.HideSelection = True
		Me.txtSuDocNumber.ReadOnly = False
		Me.txtSuDocNumber.Maxlength = 0
		Me.txtSuDocNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSuDocNumber.MultiLine = False
		Me.txtSuDocNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSuDocNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSuDocNumber.TabStop = True
		Me.txtSuDocNumber.Visible = True
		Me.txtSuDocNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSuDocNumber.Name = "txtSuDocNumber"
		Me.txtLargerWorkID.AutoSize = False
		Me.txtLargerWorkID.Size = New System.Drawing.Size(97, 19)
		Me.txtLargerWorkID.Location = New System.Drawing.Point(128, 824)
		Me.txtLargerWorkID.TabIndex = 41
		Me.txtLargerWorkID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLargerWorkID.AcceptsReturn = True
		Me.txtLargerWorkID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLargerWorkID.BackColor = System.Drawing.SystemColors.Window
		Me.txtLargerWorkID.CausesValidation = True
		Me.txtLargerWorkID.Enabled = True
		Me.txtLargerWorkID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLargerWorkID.HideSelection = True
		Me.txtLargerWorkID.ReadOnly = False
		Me.txtLargerWorkID.Maxlength = 0
		Me.txtLargerWorkID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLargerWorkID.MultiLine = False
		Me.txtLargerWorkID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLargerWorkID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLargerWorkID.TabStop = True
		Me.txtLargerWorkID.Visible = True
		Me.txtLargerWorkID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLargerWorkID.Name = "txtLargerWorkID"
		Me.cmbLargerWorkTitle.Size = New System.Drawing.Size(353, 21)
		Me.cmbLargerWorkTitle.Location = New System.Drawing.Point(608, 688)
		Me.cmbLargerWorkTitle.Sorted = True
		Me.cmbLargerWorkTitle.TabIndex = 39
		Me.cmbLargerWorkTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbLargerWorkTitle.BackColor = System.Drawing.SystemColors.Window
		Me.cmbLargerWorkTitle.CausesValidation = True
		Me.cmbLargerWorkTitle.Enabled = True
		Me.cmbLargerWorkTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbLargerWorkTitle.IntegralHeight = True
		Me.cmbLargerWorkTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbLargerWorkTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbLargerWorkTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbLargerWorkTitle.TabStop = True
		Me.cmbLargerWorkTitle.Visible = True
		Me.cmbLargerWorkTitle.Name = "cmbLargerWorkTitle"
		Me.txtReportOrDocumentNumber.AutoSize = False
		Me.txtReportOrDocumentNumber.Size = New System.Drawing.Size(145, 19)
		Me.txtReportOrDocumentNumber.Location = New System.Drawing.Point(864, 504)
		Me.txtReportOrDocumentNumber.TabIndex = 38
		Me.txtReportOrDocumentNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtReportOrDocumentNumber.AcceptsReturn = True
		Me.txtReportOrDocumentNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtReportOrDocumentNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtReportOrDocumentNumber.CausesValidation = True
		Me.txtReportOrDocumentNumber.Enabled = True
		Me.txtReportOrDocumentNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtReportOrDocumentNumber.HideSelection = True
		Me.txtReportOrDocumentNumber.ReadOnly = False
		Me.txtReportOrDocumentNumber.Maxlength = 0
		Me.txtReportOrDocumentNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtReportOrDocumentNumber.MultiLine = False
		Me.txtReportOrDocumentNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtReportOrDocumentNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtReportOrDocumentNumber.TabStop = True
		Me.txtReportOrDocumentNumber.Visible = True
		Me.txtReportOrDocumentNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtReportOrDocumentNumber.Name = "txtReportOrDocumentNumber"
		Me.txtUSCCANCitation.AutoSize = False
		Me.txtUSCCANCitation.Size = New System.Drawing.Size(105, 19)
		Me.txtUSCCANCitation.Location = New System.Drawing.Point(864, 480)
		Me.txtUSCCANCitation.TabIndex = 37
		Me.txtUSCCANCitation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUSCCANCitation.AcceptsReturn = True
		Me.txtUSCCANCitation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUSCCANCitation.BackColor = System.Drawing.SystemColors.Window
		Me.txtUSCCANCitation.CausesValidation = True
		Me.txtUSCCANCitation.Enabled = True
		Me.txtUSCCANCitation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUSCCANCitation.HideSelection = True
		Me.txtUSCCANCitation.ReadOnly = False
		Me.txtUSCCANCitation.Maxlength = 0
		Me.txtUSCCANCitation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUSCCANCitation.MultiLine = False
		Me.txtUSCCANCitation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUSCCANCitation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUSCCANCitation.TabStop = True
		Me.txtUSCCANCitation.Visible = True
		Me.txtUSCCANCitation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUSCCANCitation.Name = "txtUSCCANCitation"
		Me.txtStateLegislativeSession.AutoSize = False
		Me.txtStateLegislativeSession.Size = New System.Drawing.Size(97, 19)
		Me.txtStateLegislativeSession.Location = New System.Drawing.Point(864, 320)
		Me.txtStateLegislativeSession.TabIndex = 36
		Me.txtStateLegislativeSession.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtStateLegislativeSession.AcceptsReturn = True
		Me.txtStateLegislativeSession.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStateLegislativeSession.BackColor = System.Drawing.SystemColors.Window
		Me.txtStateLegislativeSession.CausesValidation = True
		Me.txtStateLegislativeSession.Enabled = True
		Me.txtStateLegislativeSession.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStateLegislativeSession.HideSelection = True
		Me.txtStateLegislativeSession.ReadOnly = False
		Me.txtStateLegislativeSession.Maxlength = 0
		Me.txtStateLegislativeSession.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStateLegislativeSession.MultiLine = False
		Me.txtStateLegislativeSession.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStateLegislativeSession.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStateLegislativeSession.TabStop = True
		Me.txtStateLegislativeSession.Visible = True
		Me.txtStateLegislativeSession.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStateLegislativeSession.Name = "txtStateLegislativeSession"
		Me.txtSessionOfCongress.AutoSize = False
		Me.txtSessionOfCongress.Size = New System.Drawing.Size(113, 19)
		Me.txtSessionOfCongress.Location = New System.Drawing.Point(864, 304)
		Me.txtSessionOfCongress.TabIndex = 35
		Me.txtSessionOfCongress.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSessionOfCongress.AcceptsReturn = True
		Me.txtSessionOfCongress.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSessionOfCongress.BackColor = System.Drawing.SystemColors.Window
		Me.txtSessionOfCongress.CausesValidation = True
		Me.txtSessionOfCongress.Enabled = True
		Me.txtSessionOfCongress.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSessionOfCongress.HideSelection = True
		Me.txtSessionOfCongress.ReadOnly = False
		Me.txtSessionOfCongress.Maxlength = 0
		Me.txtSessionOfCongress.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSessionOfCongress.MultiLine = False
		Me.txtSessionOfCongress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSessionOfCongress.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSessionOfCongress.TabStop = True
		Me.txtSessionOfCongress.Visible = True
		Me.txtSessionOfCongress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSessionOfCongress.Name = "txtSessionOfCongress"
		Me.txtNumberOfCongress.AutoSize = False
		Me.txtNumberOfCongress.Size = New System.Drawing.Size(105, 19)
		Me.txtNumberOfCongress.Location = New System.Drawing.Point(864, 464)
		Me.txtNumberOfCongress.TabIndex = 34
		Me.txtNumberOfCongress.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNumberOfCongress.AcceptsReturn = True
		Me.txtNumberOfCongress.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNumberOfCongress.BackColor = System.Drawing.SystemColors.Window
		Me.txtNumberOfCongress.CausesValidation = True
		Me.txtNumberOfCongress.Enabled = True
		Me.txtNumberOfCongress.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNumberOfCongress.HideSelection = True
		Me.txtNumberOfCongress.ReadOnly = False
		Me.txtNumberOfCongress.Maxlength = 0
		Me.txtNumberOfCongress.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNumberOfCongress.MultiLine = False
		Me.txtNumberOfCongress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNumberOfCongress.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNumberOfCongress.TabStop = True
		Me.txtNumberOfCongress.Visible = True
		Me.txtNumberOfCongress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNumberOfCongress.Name = "txtNumberOfCongress"
		Me.txtLegislativeHouse.AutoSize = False
		Me.txtLegislativeHouse.Size = New System.Drawing.Size(121, 19)
		Me.txtLegislativeHouse.Location = New System.Drawing.Point(864, 408)
		Me.txtLegislativeHouse.TabIndex = 33
		Me.txtLegislativeHouse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLegislativeHouse.AcceptsReturn = True
		Me.txtLegislativeHouse.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLegislativeHouse.BackColor = System.Drawing.SystemColors.Window
		Me.txtLegislativeHouse.CausesValidation = True
		Me.txtLegislativeHouse.Enabled = True
		Me.txtLegislativeHouse.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLegislativeHouse.HideSelection = True
		Me.txtLegislativeHouse.ReadOnly = False
		Me.txtLegislativeHouse.Maxlength = 0
		Me.txtLegislativeHouse.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLegislativeHouse.MultiLine = False
		Me.txtLegislativeHouse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLegislativeHouse.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLegislativeHouse.TabStop = True
		Me.txtLegislativeHouse.Visible = True
		Me.txtLegislativeHouse.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLegislativeHouse.Name = "txtLegislativeHouse"
		Me.cmbLegislativeType.Size = New System.Drawing.Size(137, 21)
		Me.cmbLegislativeType.Location = New System.Drawing.Point(864, 368)
		Me.cmbLegislativeType.TabIndex = 32
		Me.cmbLegislativeType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbLegislativeType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbLegislativeType.CausesValidation = True
		Me.cmbLegislativeType.Enabled = True
		Me.cmbLegislativeType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbLegislativeType.IntegralHeight = True
		Me.cmbLegislativeType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbLegislativeType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbLegislativeType.Sorted = False
		Me.cmbLegislativeType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbLegislativeType.TabStop = True
		Me.cmbLegislativeType.Visible = True
		Me.cmbLegislativeType.Name = "cmbLegislativeType"
		Me.cmbMiscType.Size = New System.Drawing.Size(145, 21)
		Me.cmbMiscType.Location = New System.Drawing.Point(784, 712)
		Me.cmbMiscType.TabIndex = 31
		Me.cmbMiscType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbMiscType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbMiscType.CausesValidation = True
		Me.cmbMiscType.Enabled = True
		Me.cmbMiscType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbMiscType.IntegralHeight = True
		Me.cmbMiscType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbMiscType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbMiscType.Sorted = False
		Me.cmbMiscType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbMiscType.TabStop = True
		Me.cmbMiscType.Visible = True
		Me.cmbMiscType.Name = "cmbMiscType"
		Me.txtLocation.AutoSize = False
		Me.txtLocation.Size = New System.Drawing.Size(161, 19)
		Me.txtLocation.Location = New System.Drawing.Point(352, 720)
		Me.txtLocation.TabIndex = 30
		Me.txtLocation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLocation.AcceptsReturn = True
		Me.txtLocation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLocation.BackColor = System.Drawing.SystemColors.Window
		Me.txtLocation.CausesValidation = True
		Me.txtLocation.Enabled = True
		Me.txtLocation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLocation.HideSelection = True
		Me.txtLocation.ReadOnly = False
		Me.txtLocation.Maxlength = 0
		Me.txtLocation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLocation.MultiLine = False
		Me.txtLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLocation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLocation.TabStop = True
		Me.txtLocation.Visible = True
		Me.txtLocation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLocation.Name = "txtLocation"
		Me.cmbUnpublishedType.Size = New System.Drawing.Size(113, 21)
		Me.cmbUnpublishedType.Location = New System.Drawing.Point(584, 712)
		Me.cmbUnpublishedType.TabIndex = 27
		Me.cmbUnpublishedType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbUnpublishedType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbUnpublishedType.CausesValidation = True
		Me.cmbUnpublishedType.Enabled = True
		Me.cmbUnpublishedType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbUnpublishedType.IntegralHeight = True
		Me.cmbUnpublishedType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbUnpublishedType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbUnpublishedType.Sorted = False
		Me.cmbUnpublishedType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbUnpublishedType.TabStop = True
		Me.cmbUnpublishedType.Visible = True
		Me.cmbUnpublishedType.Name = "cmbUnpublishedType"
		Me.cmbThesisDissertationType.Size = New System.Drawing.Size(129, 21)
		Me.cmbThesisDissertationType.Location = New System.Drawing.Point(640, 656)
		Me.cmbThesisDissertationType.TabIndex = 26
		Me.cmbThesisDissertationType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbThesisDissertationType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbThesisDissertationType.CausesValidation = True
		Me.cmbThesisDissertationType.Enabled = True
		Me.cmbThesisDissertationType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbThesisDissertationType.IntegralHeight = True
		Me.cmbThesisDissertationType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbThesisDissertationType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbThesisDissertationType.Sorted = False
		Me.cmbThesisDissertationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbThesisDissertationType.TabStop = True
		Me.cmbThesisDissertationType.Visible = True
		Me.cmbThesisDissertationType.Name = "cmbThesisDissertationType"
		Me.chkAllChaptersBySameAuthor.Text = "All Chapters By Same Author?"
		Me.chkAllChaptersBySameAuthor.Size = New System.Drawing.Size(177, 17)
		Me.chkAllChaptersBySameAuthor.Location = New System.Drawing.Point(304, 792)
		Me.chkAllChaptersBySameAuthor.TabIndex = 25
		Me.chkAllChaptersBySameAuthor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAllChaptersBySameAuthor.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAllChaptersBySameAuthor.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAllChaptersBySameAuthor.BackColor = System.Drawing.SystemColors.Control
		Me.chkAllChaptersBySameAuthor.CausesValidation = True
		Me.chkAllChaptersBySameAuthor.Enabled = True
		Me.chkAllChaptersBySameAuthor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAllChaptersBySameAuthor.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAllChaptersBySameAuthor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAllChaptersBySameAuthor.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAllChaptersBySameAuthor.TabStop = True
		Me.chkAllChaptersBySameAuthor.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAllChaptersBySameAuthor.Visible = True
		Me.chkAllChaptersBySameAuthor.Name = "chkAllChaptersBySameAuthor"
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.AutoSize = False
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Size = New System.Drawing.Size(273, 19)
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Location = New System.Drawing.Point(472, 784)
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.TabIndex = 24
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.AcceptsReturn = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.BackColor = System.Drawing.SystemColors.Window
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.CausesValidation = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Enabled = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.HideSelection = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.ReadOnly = False
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Maxlength = 0
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.MultiLine = False
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.TabStop = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Visible = True
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Name = "txtTitleOfSeriesIfNotIssuedByAuthor"
		Me.txtSeriesVolume.AutoSize = False
		Me.txtSeriesVolume.Size = New System.Drawing.Size(73, 19)
		Me.txtSeriesVolume.Location = New System.Drawing.Point(496, 688)
		Me.txtSeriesVolume.TabIndex = 23
		Me.txtSeriesVolume.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSeriesVolume.AcceptsReturn = True
		Me.txtSeriesVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSeriesVolume.BackColor = System.Drawing.SystemColors.Window
		Me.txtSeriesVolume.CausesValidation = True
		Me.txtSeriesVolume.Enabled = True
		Me.txtSeriesVolume.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSeriesVolume.HideSelection = True
		Me.txtSeriesVolume.ReadOnly = False
		Me.txtSeriesVolume.Maxlength = 0
		Me.txtSeriesVolume.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSeriesVolume.MultiLine = False
		Me.txtSeriesVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSeriesVolume.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSeriesVolume.TabStop = True
		Me.txtSeriesVolume.Visible = True
		Me.txtSeriesVolume.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSeriesVolume.Name = "txtSeriesVolume"
		Me.txtOriginalPublicationDate.AutoSize = False
		Me.txtOriginalPublicationDate.Size = New System.Drawing.Size(81, 19)
		Me.txtOriginalPublicationDate.Location = New System.Drawing.Point(880, 552)
		Me.txtOriginalPublicationDate.TabIndex = 22
		Me.txtOriginalPublicationDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOriginalPublicationDate.AcceptsReturn = True
		Me.txtOriginalPublicationDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOriginalPublicationDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtOriginalPublicationDate.CausesValidation = True
		Me.txtOriginalPublicationDate.Enabled = True
		Me.txtOriginalPublicationDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOriginalPublicationDate.HideSelection = True
		Me.txtOriginalPublicationDate.ReadOnly = False
		Me.txtOriginalPublicationDate.Maxlength = 0
		Me.txtOriginalPublicationDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOriginalPublicationDate.MultiLine = False
		Me.txtOriginalPublicationDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOriginalPublicationDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOriginalPublicationDate.TabStop = True
		Me.txtOriginalPublicationDate.Visible = True
		Me.txtOriginalPublicationDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOriginalPublicationDate.Name = "txtOriginalPublicationDate"
		Me.txtPublisher.AutoSize = False
		Me.txtPublisher.Size = New System.Drawing.Size(321, 19)
		Me.txtPublisher.Location = New System.Drawing.Point(640, 848)
		Me.txtPublisher.TabIndex = 18
		Me.txtPublisher.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPublisher.AcceptsReturn = True
		Me.txtPublisher.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPublisher.BackColor = System.Drawing.SystemColors.Window
		Me.txtPublisher.CausesValidation = True
		Me.txtPublisher.Enabled = True
		Me.txtPublisher.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPublisher.HideSelection = True
		Me.txtPublisher.ReadOnly = False
		Me.txtPublisher.Maxlength = 0
		Me.txtPublisher.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPublisher.MultiLine = False
		Me.txtPublisher.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPublisher.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPublisher.TabStop = True
		Me.txtPublisher.Visible = True
		Me.txtPublisher.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPublisher.Name = "txtPublisher"
		Me.txtEditionAndPrinting.AutoSize = False
		Me.txtEditionAndPrinting.Size = New System.Drawing.Size(81, 19)
		Me.txtEditionAndPrinting.Location = New System.Drawing.Point(176, 768)
		Me.txtEditionAndPrinting.TabIndex = 14
		Me.txtEditionAndPrinting.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEditionAndPrinting.AcceptsReturn = True
		Me.txtEditionAndPrinting.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEditionAndPrinting.BackColor = System.Drawing.SystemColors.Window
		Me.txtEditionAndPrinting.CausesValidation = True
		Me.txtEditionAndPrinting.Enabled = True
		Me.txtEditionAndPrinting.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEditionAndPrinting.HideSelection = True
		Me.txtEditionAndPrinting.ReadOnly = False
		Me.txtEditionAndPrinting.Maxlength = 0
		Me.txtEditionAndPrinting.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEditionAndPrinting.MultiLine = False
		Me.txtEditionAndPrinting.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEditionAndPrinting.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEditionAndPrinting.TabStop = True
		Me.txtEditionAndPrinting.Visible = True
		Me.txtEditionAndPrinting.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEditionAndPrinting.Name = "txtEditionAndPrinting"
		Me.txtOrganizationIssuingNewsletter.AutoSize = False
		Me.txtOrganizationIssuingNewsletter.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtOrganizationIssuingNewsletter.Enabled = False
		Me.txtOrganizationIssuingNewsletter.Size = New System.Drawing.Size(265, 19)
		Me.txtOrganizationIssuingNewsletter.Location = New System.Drawing.Point(344, 848)
		Me.txtOrganizationIssuingNewsletter.TabIndex = 12
		Me.txtOrganizationIssuingNewsletter.Visible = False
		Me.txtOrganizationIssuingNewsletter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOrganizationIssuingNewsletter.AcceptsReturn = True
		Me.txtOrganizationIssuingNewsletter.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOrganizationIssuingNewsletter.CausesValidation = True
		Me.txtOrganizationIssuingNewsletter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrganizationIssuingNewsletter.HideSelection = True
		Me.txtOrganizationIssuingNewsletter.ReadOnly = False
		Me.txtOrganizationIssuingNewsletter.Maxlength = 0
		Me.txtOrganizationIssuingNewsletter.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrganizationIssuingNewsletter.MultiLine = False
		Me.txtOrganizationIssuingNewsletter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrganizationIssuingNewsletter.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrganizationIssuingNewsletter.TabStop = True
		Me.txtOrganizationIssuingNewsletter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrganizationIssuingNewsletter.Name = "txtOrganizationIssuingNewsletter"
		Me.txtNotes.AutoSize = False
		Me.txtNotes.Size = New System.Drawing.Size(817, 49)
		Me.txtNotes.Location = New System.Drawing.Point(32, 568)
		Me.txtNotes.MultiLine = True
		Me.txtNotes.TabIndex = 19
		Me.txtNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNotes.AcceptsReturn = True
		Me.txtNotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNotes.BackColor = System.Drawing.SystemColors.Window
		Me.txtNotes.CausesValidation = True
		Me.txtNotes.Enabled = True
		Me.txtNotes.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNotes.HideSelection = True
		Me.txtNotes.ReadOnly = False
		Me.txtNotes.Maxlength = 0
		Me.txtNotes.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNotes.TabStop = True
		Me.txtNotes.Visible = True
		Me.txtNotes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNotes.Name = "txtNotes"
		Me.txtPage.AutoSize = False
		Me.txtPage.Size = New System.Drawing.Size(81, 19)
		Me.txtPage.Location = New System.Drawing.Point(552, 280)
		Me.txtPage.TabIndex = 6
		Me.txtPage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPage.AcceptsReturn = True
		Me.txtPage.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPage.BackColor = System.Drawing.SystemColors.Window
		Me.txtPage.CausesValidation = True
		Me.txtPage.Enabled = True
		Me.txtPage.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPage.HideSelection = True
		Me.txtPage.ReadOnly = False
		Me.txtPage.Maxlength = 0
		Me.txtPage.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPage.MultiLine = False
		Me.txtPage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPage.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPage.TabStop = True
		Me.txtPage.Visible = True
		Me.txtPage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPage.Name = "txtPage"
		Me.cmbPublicationMonthOrSeason.Size = New System.Drawing.Size(137, 21)
		Me.cmbPublicationMonthOrSeason.Location = New System.Drawing.Point(216, 280)
		Me.cmbPublicationMonthOrSeason.TabIndex = 4
		Me.cmbPublicationMonthOrSeason.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbPublicationMonthOrSeason.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPublicationMonthOrSeason.CausesValidation = True
		Me.cmbPublicationMonthOrSeason.Enabled = True
		Me.cmbPublicationMonthOrSeason.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPublicationMonthOrSeason.IntegralHeight = True
		Me.cmbPublicationMonthOrSeason.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPublicationMonthOrSeason.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPublicationMonthOrSeason.Sorted = False
		Me.cmbPublicationMonthOrSeason.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbPublicationMonthOrSeason.TabStop = True
		Me.cmbPublicationMonthOrSeason.Visible = True
		Me.cmbPublicationMonthOrSeason.Name = "cmbPublicationMonthOrSeason"
		Me.txtVolume.AutoSize = False
		Me.txtVolume.Size = New System.Drawing.Size(81, 19)
		Me.txtVolume.Location = New System.Drawing.Point(696, 280)
		Me.txtVolume.TabIndex = 7
		Me.txtVolume.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtVolume.AcceptsReturn = True
		Me.txtVolume.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtVolume.BackColor = System.Drawing.SystemColors.Window
		Me.txtVolume.CausesValidation = True
		Me.txtVolume.Enabled = True
		Me.txtVolume.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtVolume.HideSelection = True
		Me.txtVolume.ReadOnly = False
		Me.txtVolume.Maxlength = 0
		Me.txtVolume.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtVolume.MultiLine = False
		Me.txtVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtVolume.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtVolume.TabStop = True
		Me.txtVolume.Visible = True
		Me.txtVolume.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtVolume.Name = "txtVolume"
		Me.txtPublicationDay.AutoSize = False
		Me.txtPublicationDay.Size = New System.Drawing.Size(73, 19)
		Me.txtPublicationDay.Location = New System.Drawing.Point(416, 280)
		Me.txtPublicationDay.TabIndex = 5
		Me.txtPublicationDay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPublicationDay.AcceptsReturn = True
		Me.txtPublicationDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPublicationDay.BackColor = System.Drawing.SystemColors.Window
		Me.txtPublicationDay.CausesValidation = True
		Me.txtPublicationDay.Enabled = True
		Me.txtPublicationDay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPublicationDay.HideSelection = True
		Me.txtPublicationDay.ReadOnly = False
		Me.txtPublicationDay.Maxlength = 0
		Me.txtPublicationDay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPublicationDay.MultiLine = False
		Me.txtPublicationDay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPublicationDay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPublicationDay.TabStop = True
		Me.txtPublicationDay.Visible = True
		Me.txtPublicationDay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPublicationDay.Name = "txtPublicationDay"
		Me.txtJournalID.AutoSize = False
		Me.txtJournalID.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtJournalID.Enabled = False
		Me.txtJournalID.Size = New System.Drawing.Size(81, 21)
		Me.txtJournalID.Location = New System.Drawing.Point(224, 808)
		Me.txtJournalID.TabIndex = 10
		Me.txtJournalID.Visible = False
		Me.txtJournalID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtJournalID.AcceptsReturn = True
		Me.txtJournalID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtJournalID.CausesValidation = True
		Me.txtJournalID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtJournalID.HideSelection = True
		Me.txtJournalID.ReadOnly = False
		Me.txtJournalID.Maxlength = 0
		Me.txtJournalID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtJournalID.MultiLine = False
		Me.txtJournalID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtJournalID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtJournalID.TabStop = True
		Me.txtJournalID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtJournalID.Name = "txtJournalID"
		Me.cmbJournalTitle.Size = New System.Drawing.Size(409, 21)
		Me.cmbJournalTitle.Location = New System.Drawing.Point(32, 184)
		Me.cmbJournalTitle.Sorted = True
		Me.cmbJournalTitle.TabIndex = 0
		Me.cmbJournalTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbJournalTitle.BackColor = System.Drawing.SystemColors.Window
		Me.cmbJournalTitle.CausesValidation = True
		Me.cmbJournalTitle.Enabled = True
		Me.cmbJournalTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbJournalTitle.IntegralHeight = True
		Me.cmbJournalTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbJournalTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbJournalTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbJournalTitle.TabStop = True
		Me.cmbJournalTitle.Visible = True
		Me.cmbJournalTitle.Name = "cmbJournalTitle"
		Me.cmbArticleDesignation.Size = New System.Drawing.Size(145, 21)
		Me.cmbArticleDesignation.Location = New System.Drawing.Point(32, 280)
		Me.cmbArticleDesignation.TabIndex = 3
		Me.cmbArticleDesignation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbArticleDesignation.BackColor = System.Drawing.SystemColors.Window
		Me.cmbArticleDesignation.CausesValidation = True
		Me.cmbArticleDesignation.Enabled = True
		Me.cmbArticleDesignation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbArticleDesignation.IntegralHeight = True
		Me.cmbArticleDesignation.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbArticleDesignation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbArticleDesignation.Sorted = False
		Me.cmbArticleDesignation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbArticleDesignation.TabStop = True
		Me.cmbArticleDesignation.Visible = True
		Me.cmbArticleDesignation.Name = "cmbArticleDesignation"
		Me.txtTitle.AutoSize = False
		Me.txtTitle.Size = New System.Drawing.Size(785, 21)
		Me.txtTitle.Location = New System.Drawing.Point(32, 232)
		Me.txtTitle.TabIndex = 2
		Me.txtTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTitle.AcceptsReturn = True
		Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTitle.BackColor = System.Drawing.SystemColors.Window
		Me.txtTitle.CausesValidation = True
		Me.txtTitle.Enabled = True
		Me.txtTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTitle.HideSelection = True
		Me.txtTitle.ReadOnly = False
		Me.txtTitle.Maxlength = 0
		Me.txtTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTitle.MultiLine = False
		Me.txtTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTitle.TabStop = True
		Me.txtTitle.Visible = True
		Me.txtTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTitle.Name = "txtTitle"
		Me.txtYear.AutoSize = False
		Me.txtYear.Size = New System.Drawing.Size(73, 21)
		Me.txtYear.Location = New System.Drawing.Point(568, 184)
		Me.txtYear.TabIndex = 1
		Me.txtYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtYear.AcceptsReturn = True
		Me.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtYear.BackColor = System.Drawing.SystemColors.Window
		Me.txtYear.CausesValidation = True
		Me.txtYear.Enabled = True
		Me.txtYear.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtYear.HideSelection = True
		Me.txtYear.ReadOnly = False
		Me.txtYear.Maxlength = 0
		Me.txtYear.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtYear.MultiLine = False
		Me.txtYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtYear.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtYear.TabStop = True
		Me.txtYear.Visible = True
		Me.txtYear.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtYear.Name = "txtYear"
		Me.txtInputInitials.AutoSize = False
		Me.txtInputInitials.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtInputInitials.Enabled = False
		Me.txtInputInitials.Size = New System.Drawing.Size(89, 21)
		Me.txtInputInitials.Location = New System.Drawing.Point(752, 104)
		Me.txtInputInitials.TabIndex = 46
		Me.txtInputInitials.TabStop = False
		Me.txtInputInitials.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInputInitials.AcceptsReturn = True
		Me.txtInputInitials.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInputInitials.CausesValidation = True
		Me.txtInputInitials.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInputInitials.HideSelection = True
		Me.txtInputInitials.ReadOnly = False
		Me.txtInputInitials.Maxlength = 0
		Me.txtInputInitials.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInputInitials.MultiLine = False
		Me.txtInputInitials.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInputInitials.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInputInitials.Visible = True
		Me.txtInputInitials.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInputInitials.Name = "txtInputInitials"
		Me.txtDateUpdated.AutoSize = False
		Me.txtDateUpdated.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtDateUpdated.Enabled = False
		Me.txtDateUpdated.ForeColor = System.Drawing.SystemColors.ControlText
		Me.txtDateUpdated.Size = New System.Drawing.Size(89, 21)
		Me.txtDateUpdated.Location = New System.Drawing.Point(640, 104)
		Me.txtDateUpdated.TabIndex = 47
		Me.txtDateUpdated.TabStop = False
		Me.txtDateUpdated.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDateUpdated.AcceptsReturn = True
		Me.txtDateUpdated.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDateUpdated.CausesValidation = True
		Me.txtDateUpdated.HideSelection = True
		Me.txtDateUpdated.ReadOnly = False
		Me.txtDateUpdated.Maxlength = 0
		Me.txtDateUpdated.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDateUpdated.MultiLine = False
		Me.txtDateUpdated.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDateUpdated.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDateUpdated.Visible = True
		Me.txtDateUpdated.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDateUpdated.Name = "txtDateUpdated"
		Me.txtDateAdded.AutoSize = False
		Me.txtDateAdded.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtDateAdded.Enabled = False
		Me.txtDateAdded.Size = New System.Drawing.Size(89, 21)
		Me.txtDateAdded.Location = New System.Drawing.Point(536, 104)
		Me.txtDateAdded.TabIndex = 48
		Me.txtDateAdded.TabStop = False
		Me.txtDateAdded.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDateAdded.AcceptsReturn = True
		Me.txtDateAdded.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDateAdded.CausesValidation = True
		Me.txtDateAdded.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDateAdded.HideSelection = True
		Me.txtDateAdded.ReadOnly = False
		Me.txtDateAdded.Maxlength = 0
		Me.txtDateAdded.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDateAdded.MultiLine = False
		Me.txtDateAdded.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDateAdded.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDateAdded.Visible = True
		Me.txtDateAdded.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDateAdded.Name = "txtDateAdded"
		Me.cmbSourceType.Size = New System.Drawing.Size(257, 21)
		Me.cmbSourceType.Location = New System.Drawing.Point(144, 104)
		Me.cmbSourceType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbSourceType.TabIndex = 124
		Me.cmbSourceType.TabStop = False
		Me.cmbSourceType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbSourceType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbSourceType.CausesValidation = True
		Me.cmbSourceType.Enabled = True
		Me.cmbSourceType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbSourceType.IntegralHeight = True
		Me.cmbSourceType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbSourceType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbSourceType.Sorted = False
		Me.cmbSourceType.Visible = True
		Me.cmbSourceType.Name = "cmbSourceType"
		Me.lstAuthors.Size = New System.Drawing.Size(281, 59)
		Me.lstAuthors.Location = New System.Drawing.Point(32, 368)
		Me.lstAuthors.Sorted = True
		Me.lstAuthors.TabIndex = 15
		Me.lstAuthors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstAuthors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstAuthors.BackColor = System.Drawing.SystemColors.Window
		Me.lstAuthors.CausesValidation = True
		Me.lstAuthors.Enabled = True
		Me.lstAuthors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstAuthors.IntegralHeight = True
		Me.lstAuthors.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstAuthors.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstAuthors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstAuthors.TabStop = True
		Me.lstAuthors.Visible = True
		Me.lstAuthors.MultiColumn = False
		Me.lstAuthors.Name = "lstAuthors"
		Me.lstTranslators.Enabled = False
		Me.lstTranslators.Size = New System.Drawing.Size(281, 59)
		Me.lstTranslators.Location = New System.Drawing.Point(32, 368)
		Me.lstTranslators.Sorted = True
		Me.lstTranslators.TabIndex = 21
		Me.lstTranslators.Visible = False
		Me.lstTranslators.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstTranslators.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstTranslators.BackColor = System.Drawing.SystemColors.Window
		Me.lstTranslators.CausesValidation = True
		Me.lstTranslators.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstTranslators.IntegralHeight = True
		Me.lstTranslators.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstTranslators.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstTranslators.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstTranslators.TabStop = True
		Me.lstTranslators.MultiColumn = False
		Me.lstTranslators.Name = "lstTranslators"
		Me.lstCurrentAuthors.Size = New System.Drawing.Size(281, 59)
		Me.lstCurrentAuthors.Location = New System.Drawing.Point(376, 368)
		Me.lstCurrentAuthors.TabIndex = 119
		Me.lstCurrentAuthors.TabStop = False
		Me.lstCurrentAuthors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCurrentAuthors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCurrentAuthors.BackColor = System.Drawing.SystemColors.Window
		Me.lstCurrentAuthors.CausesValidation = True
		Me.lstCurrentAuthors.Enabled = True
		Me.lstCurrentAuthors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCurrentAuthors.IntegralHeight = True
		Me.lstCurrentAuthors.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCurrentAuthors.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCurrentAuthors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCurrentAuthors.Sorted = False
		Me.lstCurrentAuthors.Visible = True
		Me.lstCurrentAuthors.MultiColumn = False
		Me.lstCurrentAuthors.Name = "lstCurrentAuthors"
		Me.lstCurrentTranslators.Enabled = False
		Me.lstCurrentTranslators.Size = New System.Drawing.Size(281, 59)
		Me.lstCurrentTranslators.Location = New System.Drawing.Point(376, 368)
		Me.lstCurrentTranslators.Sorted = True
		Me.lstCurrentTranslators.TabIndex = 17
		Me.lstCurrentTranslators.Visible = False
		Me.lstCurrentTranslators.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCurrentTranslators.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCurrentTranslators.BackColor = System.Drawing.SystemColors.Window
		Me.lstCurrentTranslators.CausesValidation = True
		Me.lstCurrentTranslators.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCurrentTranslators.IntegralHeight = True
		Me.lstCurrentTranslators.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCurrentTranslators.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCurrentTranslators.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCurrentTranslators.TabStop = True
		Me.lstCurrentTranslators.MultiColumn = False
		Me.lstCurrentTranslators.Name = "lstCurrentTranslators"
		Me.lstCurrentEditors.Enabled = False
		Me.lstCurrentEditors.Size = New System.Drawing.Size(281, 59)
		Me.lstCurrentEditors.Location = New System.Drawing.Point(376, 368)
		Me.lstCurrentEditors.TabIndex = 52
		Me.lstCurrentEditors.Visible = False
		Me.lstCurrentEditors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCurrentEditors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCurrentEditors.BackColor = System.Drawing.SystemColors.Window
		Me.lstCurrentEditors.CausesValidation = True
		Me.lstCurrentEditors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCurrentEditors.IntegralHeight = True
		Me.lstCurrentEditors.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCurrentEditors.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCurrentEditors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCurrentEditors.Sorted = False
		Me.lstCurrentEditors.TabStop = True
		Me.lstCurrentEditors.MultiColumn = False
		Me.lstCurrentEditors.Name = "lstCurrentEditors"
		Me.lstEditors.Enabled = False
		Me.lstEditors.Size = New System.Drawing.Size(281, 59)
		Me.lstEditors.Location = New System.Drawing.Point(32, 368)
		Me.lstEditors.Sorted = True
		Me.lstEditors.TabIndex = 13
		Me.lstEditors.Visible = False
		Me.lstEditors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstEditors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstEditors.BackColor = System.Drawing.SystemColors.Window
		Me.lstEditors.CausesValidation = True
		Me.lstEditors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstEditors.IntegralHeight = True
		Me.lstEditors.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstEditors.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstEditors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstEditors.TabStop = True
		Me.lstEditors.MultiColumn = False
		Me.lstEditors.Name = "lstEditors"
		Me.lblT.AutoSize = False
		Me.lblT.BackColor = System.Drawing.SystemColors.Control
		Me.lblT.Enabled = False
		Me.lblT.ForeColor = System.Drawing.SystemColors.InactiveCaption
		Me.lblT.Size = New System.Drawing.Size(73, 13)
		Me.lblT.Location = New System.Drawing.Point(712, 416)
		Me.lblT.TabIndex = 8
		Me.lblT.Text = "No Translator"
		Me.lblT.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblT.AcceptsReturn = True
		Me.lblT.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblT.CausesValidation = True
		Me.lblT.HideSelection = True
		Me.lblT.ReadOnly = False
		Me.lblT.Maxlength = 0
		Me.lblT.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblT.MultiLine = False
		Me.lblT.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblT.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblT.TabStop = True
		Me.lblT.Visible = True
		Me.lblT.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblT.Name = "lblT"
		Me.lblE.AutoSize = False
		Me.lblE.BackColor = System.Drawing.SystemColors.Control
		Me.lblE.Enabled = False
		Me.lblE.ForeColor = System.Drawing.SystemColors.InactiveCaption
		Me.lblE.Size = New System.Drawing.Size(73, 13)
		Me.lblE.Location = New System.Drawing.Point(712, 384)
		Me.lblE.TabIndex = 9
		Me.lblE.Text = "No Editor"
		Me.lblE.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblE.AcceptsReturn = True
		Me.lblE.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblE.CausesValidation = True
		Me.lblE.HideSelection = True
		Me.lblE.ReadOnly = False
		Me.lblE.Maxlength = 0
		Me.lblE.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblE.MultiLine = False
		Me.lblE.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblE.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblE.TabStop = True
		Me.lblE.Visible = True
		Me.lblE.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblE.Name = "lblE"
		Me.lblA.AutoSize = False
		Me.lblA.BackColor = System.Drawing.SystemColors.Control
		Me.lblA.Enabled = False
		Me.lblA.ForeColor = System.Drawing.SystemColors.InactiveCaption
		Me.lblA.Size = New System.Drawing.Size(73, 13)
		Me.lblA.Location = New System.Drawing.Point(712, 352)
		Me.lblA.TabIndex = 44
		Me.lblA.Text = "No Author"
		Me.lblA.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblA.AcceptsReturn = True
		Me.lblA.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblA.CausesValidation = True
		Me.lblA.HideSelection = True
		Me.lblA.ReadOnly = False
		Me.lblA.Maxlength = 0
		Me.lblA.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblA.MultiLine = False
		Me.lblA.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblA.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblA.TabStop = True
		Me.lblA.Visible = True
		Me.lblA.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblA.Name = "lblA"
		Me.lblAETChoice.AutoSize = False
		Me.lblAETChoice.BackColor = System.Drawing.SystemColors.Control
		Me.lblAETChoice.Enabled = False
		Me.lblAETChoice.Size = New System.Drawing.Size(49, 15)
		Me.lblAETChoice.Location = New System.Drawing.Point(32, 350)
		Me.lblAETChoice.TabIndex = 40
		Me.lblAETChoice.Text = "Select"
		Me.lblAETChoice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAETChoice.AcceptsReturn = True
		Me.lblAETChoice.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblAETChoice.CausesValidation = True
		Me.lblAETChoice.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblAETChoice.HideSelection = True
		Me.lblAETChoice.ReadOnly = False
		Me.lblAETChoice.Maxlength = 0
		Me.lblAETChoice.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblAETChoice.MultiLine = False
		Me.lblAETChoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAETChoice.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblAETChoice.TabStop = True
		Me.lblAETChoice.Visible = True
		Me.lblAETChoice.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAETChoice.Name = "lblAETChoice"
		Me.chkKeepSelected.Text = "Check to keep same jourrnal selected for multiple entries"
		Me.chkKeepSelected.Enabled = False
		Me.chkKeepSelected.Size = New System.Drawing.Size(393, 13)
		Me.chkKeepSelected.Location = New System.Drawing.Point(96, 168)
		Me.chkKeepSelected.TabIndex = 100
		Me.chkKeepSelected.TabStop = False
		Me.chkKeepSelected.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkKeepSelected.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkKeepSelected.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkKeepSelected.BackColor = System.Drawing.SystemColors.Control
		Me.chkKeepSelected.CausesValidation = True
		Me.chkKeepSelected.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkKeepSelected.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkKeepSelected.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkKeepSelected.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkKeepSelected.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkKeepSelected.Visible = True
		Me.chkKeepSelected.Name = "chkKeepSelected"
		Me.lblDoubleClickToAdd.AutoSize = False
		Me.lblDoubleClickToAdd.BackColor = System.Drawing.SystemColors.Control
		Me.lblDoubleClickToAdd.Enabled = False
		Me.lblDoubleClickToAdd.Size = New System.Drawing.Size(65, 45)
		Me.lblDoubleClickToAdd.Location = New System.Drawing.Point(312, 496)
		Me.lblDoubleClickToAdd.MultiLine = True
		Me.lblDoubleClickToAdd.TabIndex = 49
		Me.lblDoubleClickToAdd.Text = "Double-Click to Add or Remove"
		Me.lblDoubleClickToAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDoubleClickToAdd.AcceptsReturn = True
		Me.lblDoubleClickToAdd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblDoubleClickToAdd.CausesValidation = True
		Me.lblDoubleClickToAdd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblDoubleClickToAdd.HideSelection = True
		Me.lblDoubleClickToAdd.ReadOnly = False
		Me.lblDoubleClickToAdd.Maxlength = 0
		Me.lblDoubleClickToAdd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblDoubleClickToAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDoubleClickToAdd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblDoubleClickToAdd.TabStop = True
		Me.lblDoubleClickToAdd.Visible = True
		Me.lblDoubleClickToAdd.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDoubleClickToAdd.Name = "lblDoubleClickToAdd"
		Me.lblArrow.AutoSize = False
		Me.lblArrow.BackColor = System.Drawing.SystemColors.Control
		Me.lblArrow.Enabled = False
		Me.lblArrow.Size = New System.Drawing.Size(65, 13)
		Me.lblArrow.Location = New System.Drawing.Point(312, 368)
		Me.lblArrow.TabIndex = 42
		Me.lblArrow.Text = "<<<--------->>>"
		Me.lblArrow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblArrow.AcceptsReturn = True
		Me.lblArrow.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblArrow.CausesValidation = True
		Me.lblArrow.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblArrow.HideSelection = True
		Me.lblArrow.ReadOnly = False
		Me.lblArrow.Maxlength = 0
		Me.lblArrow.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblArrow.MultiLine = False
		Me.lblArrow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblArrow.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblArrow.TabStop = True
		Me.lblArrow.Visible = True
		Me.lblArrow.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblArrow.Name = "lblArrow"
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitle.Enabled = False
		Me.lblTitle.Size = New System.Drawing.Size(89, 13)
		Me.lblTitle.Location = New System.Drawing.Point(32, 216)
		Me.lblTitle.TabIndex = 103
		Me.lblTitle.Text = "Article Title"
		Me.lblTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.AcceptsReturn = True
		Me.lblTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblTitle.CausesValidation = True
		Me.lblTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblTitle.HideSelection = True
		Me.lblTitle.ReadOnly = False
		Me.lblTitle.Maxlength = 0
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblTitle.MultiLine = False
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblTitle.TabStop = True
		Me.lblTitle.Visible = True
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me.lblYear.AutoSize = False
		Me.lblYear.BackColor = System.Drawing.SystemColors.Control
		Me.lblYear.Enabled = False
		Me.lblYear.Size = New System.Drawing.Size(81, 13)
		Me.lblYear.Location = New System.Drawing.Point(568, 168)
		Me.lblYear.TabIndex = 101
		Me.lblYear.Text = "Publication Year"
		Me.lblYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblYear.AcceptsReturn = True
		Me.lblYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.lblYear.CausesValidation = True
		Me.lblYear.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblYear.HideSelection = True
		Me.lblYear.ReadOnly = False
		Me.lblYear.Maxlength = 0
		Me.lblYear.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.lblYear.MultiLine = False
		Me.lblYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblYear.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblYear.TabStop = True
		Me.lblYear.Visible = True
		Me.lblYear.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblYear.Name = "lblYear"
		Me.frmEntryInfo.Text = "Entry Information"
		Me.frmEntryInfo.Size = New System.Drawing.Size(345, 65)
		Me.frmEntryInfo.Location = New System.Drawing.Point(520, 72)
		Me.frmEntryInfo.TabIndex = 112
		Me.frmEntryInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmEntryInfo.BackColor = System.Drawing.SystemColors.Control
		Me.frmEntryInfo.Enabled = True
		Me.frmEntryInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmEntryInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmEntryInfo.Visible = True
		Me.frmEntryInfo.Padding = New System.Windows.Forms.Padding(0)
		Me.frmEntryInfo.Name = "frmEntryInfo"
		Me.frmRecordInfo.Text = "Record Information"
		Me.frmRecordInfo.Size = New System.Drawing.Size(497, 65)
		Me.frmRecordInfo.Location = New System.Drawing.Point(16, 72)
		Me.frmRecordInfo.TabIndex = 111
		Me.frmRecordInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmRecordInfo.BackColor = System.Drawing.SystemColors.Control
		Me.frmRecordInfo.Enabled = True
		Me.frmRecordInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmRecordInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmRecordInfo.Visible = True
		Me.frmRecordInfo.Padding = New System.Windows.Forms.Padding(0)
		Me.frmRecordInfo.Name = "frmRecordInfo"
		Me.frmCitationInfo.Text = "Citation Information"
		Me.frmCitationInfo.Size = New System.Drawing.Size(849, 185)
		Me.frmCitationInfo.Location = New System.Drawing.Point(16, 144)
		Me.frmCitationInfo.TabIndex = 113
		Me.frmCitationInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmCitationInfo.BackColor = System.Drawing.SystemColors.Control
		Me.frmCitationInfo.Enabled = True
		Me.frmCitationInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmCitationInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmCitationInfo.Visible = True
		Me.frmCitationInfo.Padding = New System.Windows.Forms.Padding(0)
		Me.frmCitationInfo.Name = "frmCitationInfo"
		Me.frmAuthorInfo.Text = "Author Information"
		Me.frmAuthorInfo.Size = New System.Drawing.Size(849, 121)
		Me.frmAuthorInfo.Location = New System.Drawing.Point(16, 328)
		Me.frmAuthorInfo.TabIndex = 114
		Me.frmAuthorInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmAuthorInfo.BackColor = System.Drawing.SystemColors.Control
		Me.frmAuthorInfo.Enabled = True
		Me.frmAuthorInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmAuthorInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmAuthorInfo.Visible = True
		Me.frmAuthorInfo.Padding = New System.Windows.Forms.Padding(0)
		Me.frmAuthorInfo.Name = "frmAuthorInfo"
		Me.frmKeywordInfo.Text = "Keyword Information"
		Me.frmKeywordInfo.Size = New System.Drawing.Size(849, 105)
		Me.frmKeywordInfo.Location = New System.Drawing.Point(16, 448)
		Me.frmKeywordInfo.TabIndex = 115
		Me.frmKeywordInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmKeywordInfo.BackColor = System.Drawing.SystemColors.Control
		Me.frmKeywordInfo.Enabled = True
		Me.frmKeywordInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmKeywordInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmKeywordInfo.Visible = True
		Me.frmKeywordInfo.Padding = New System.Windows.Forms.Padding(0)
		Me.frmKeywordInfo.Name = "frmKeywordInfo"
		Me.frmNotes.Text = "Notes"
		Me.frmNotes.Size = New System.Drawing.Size(849, 73)
		Me.frmNotes.Location = New System.Drawing.Point(16, 552)
		Me.frmNotes.TabIndex = 45
		Me.frmNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmNotes.BackColor = System.Drawing.SystemColors.Control
		Me.frmNotes.Enabled = True
		Me.frmNotes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmNotes.Visible = True
		Me.frmNotes.Padding = New System.Windows.Forms.Padding(0)
		Me.frmNotes.Name = "frmNotes"
		Me.lblSeparateBottom.Size = New System.Drawing.Size(2000, 385)
		Me.lblSeparateBottom.Location = New System.Drawing.Point(-224, 648)
		Me.lblSeparateBottom.TabIndex = 66
		Me.lblSeparateBottom.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSeparateBottom.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSeparateBottom.BackColor = System.Drawing.Color.Transparent
		Me.lblSeparateBottom.Enabled = True
		Me.lblSeparateBottom.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSeparateBottom.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSeparateBottom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSeparateBottom.UseMnemonic = True
		Me.lblSeparateBottom.Visible = True
		Me.lblSeparateBottom.AutoSize = False
		Me.lblSeparateBottom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblSeparateBottom.Name = "lblSeparateBottom"
		Me.lblMiscID.Text = "Misc ID"
		Me.lblMiscID.Size = New System.Drawing.Size(57, 17)
		Me.lblMiscID.Location = New System.Drawing.Point(40, 744)
		Me.lblMiscID.TabIndex = 65
		Me.lblMiscID.Visible = False
		Me.lblMiscID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMiscID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMiscID.BackColor = System.Drawing.SystemColors.Control
		Me.lblMiscID.Enabled = True
		Me.lblMiscID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMiscID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMiscID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMiscID.UseMnemonic = True
		Me.lblMiscID.AutoSize = False
		Me.lblMiscID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMiscID.Name = "lblMiscID"
		Me.lblTreatiseID.Text = "Treatise ID"
		Me.lblTreatiseID.Size = New System.Drawing.Size(57, 17)
		Me.lblTreatiseID.Location = New System.Drawing.Point(40, 728)
		Me.lblTreatiseID.TabIndex = 64
		Me.lblTreatiseID.Visible = False
		Me.lblTreatiseID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTreatiseID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTreatiseID.BackColor = System.Drawing.SystemColors.Control
		Me.lblTreatiseID.Enabled = True
		Me.lblTreatiseID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTreatiseID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTreatiseID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTreatiseID.UseMnemonic = True
		Me.lblTreatiseID.AutoSize = False
		Me.lblTreatiseID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTreatiseID.Name = "lblTreatiseID"
		Me.lblUnpublishedID.Text = "Unpublished ID"
		Me.lblUnpublishedID.Size = New System.Drawing.Size(57, 17)
		Me.lblUnpublishedID.Location = New System.Drawing.Point(40, 712)
		Me.lblUnpublishedID.TabIndex = 63
		Me.lblUnpublishedID.Visible = False
		Me.lblUnpublishedID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUnpublishedID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUnpublishedID.BackColor = System.Drawing.SystemColors.Control
		Me.lblUnpublishedID.Enabled = True
		Me.lblUnpublishedID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblUnpublishedID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUnpublishedID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUnpublishedID.UseMnemonic = True
		Me.lblUnpublishedID.AutoSize = False
		Me.lblUnpublishedID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUnpublishedID.Name = "lblUnpublishedID"
		Me.lblChapterID.Text = "Chapter ID"
		Me.lblChapterID.Size = New System.Drawing.Size(57, 17)
		Me.lblChapterID.Location = New System.Drawing.Point(40, 696)
		Me.lblChapterID.TabIndex = 62
		Me.lblChapterID.Visible = False
		Me.lblChapterID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblChapterID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblChapterID.BackColor = System.Drawing.SystemColors.Control
		Me.lblChapterID.Enabled = True
		Me.lblChapterID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblChapterID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblChapterID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblChapterID.UseMnemonic = True
		Me.lblChapterID.AutoSize = False
		Me.lblChapterID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblChapterID.Name = "lblChapterID"
		Me.lblArticleID.Text = "Article ID"
		Me.lblArticleID.Size = New System.Drawing.Size(57, 17)
		Me.lblArticleID.Location = New System.Drawing.Point(184, 720)
		Me.lblArticleID.TabIndex = 61
		Me.lblArticleID.Visible = False
		Me.lblArticleID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblArticleID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblArticleID.BackColor = System.Drawing.SystemColors.Control
		Me.lblArticleID.Enabled = True
		Me.lblArticleID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblArticleID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblArticleID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblArticleID.UseMnemonic = True
		Me.lblArticleID.AutoSize = False
		Me.lblArticleID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblArticleID.Name = "lblArticleID"
		Me.lblLegisID.Text = "Legis ID"
		Me.lblLegisID.Size = New System.Drawing.Size(57, 17)
		Me.lblLegisID.Location = New System.Drawing.Point(184, 704)
		Me.lblLegisID.TabIndex = 60
		Me.lblLegisID.Visible = False
		Me.lblLegisID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLegisID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLegisID.BackColor = System.Drawing.SystemColors.Control
		Me.lblLegisID.Enabled = True
		Me.lblLegisID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLegisID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLegisID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLegisID.UseMnemonic = True
		Me.lblLegisID.AutoSize = False
		Me.lblLegisID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLegisID.Name = "lblLegisID"
		tglNewRecords.OcxState = CType(resources.GetObject("tglNewRecords.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tglNewRecords.Size = New System.Drawing.Size(97, 25)
		Me.tglNewRecords.Location = New System.Drawing.Point(136, 32)
		Me.tglNewRecords.TabIndex = 68
		Me.tglNewRecords.Name = "tglNewRecords"
		tglUpdateRecords.OcxState = CType(resources.GetObject("tglUpdateRecords.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tglUpdateRecords.Size = New System.Drawing.Size(97, 25)
		Me.tglUpdateRecords.Location = New System.Drawing.Point(416, 32)
		Me.tglUpdateRecords.TabIndex = 29
		Me.tglUpdateRecords.Name = "tglUpdateRecords"
		tglImportRecords.OcxState = CType(resources.GetObject("tglImportRecords.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tglImportRecords.Size = New System.Drawing.Size(97, 25)
		Me.tglImportRecords.Location = New System.Drawing.Point(688, 32)
		Me.tglImportRecords.TabIndex = 28
		Me.tglImportRecords.Name = "tglImportRecords"
		Me.lblLargerWorkID.Text = "Larger Work ID"
		Me.lblLargerWorkID.Size = New System.Drawing.Size(89, 17)
		Me.lblLargerWorkID.Location = New System.Drawing.Point(192, 688)
		Me.lblLargerWorkID.TabIndex = 59
		Me.lblLargerWorkID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLargerWorkID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLargerWorkID.BackColor = System.Drawing.SystemColors.Control
		Me.lblLargerWorkID.Enabled = True
		Me.lblLargerWorkID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLargerWorkID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLargerWorkID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLargerWorkID.UseMnemonic = True
		Me.lblLargerWorkID.Visible = True
		Me.lblLargerWorkID.AutoSize = False
		Me.lblLargerWorkID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLargerWorkID.Name = "lblLargerWorkID"
		Me.lblNotes.Text = "Notes"
		Me.lblNotes.Size = New System.Drawing.Size(81, 17)
		Me.lblNotes.Location = New System.Drawing.Point(576, 392)
		Me.lblNotes.TabIndex = 11
		Me.lblNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblNotes.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblNotes.BackColor = System.Drawing.SystemColors.Control
		Me.lblNotes.Enabled = True
		Me.lblNotes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblNotes.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNotes.UseMnemonic = True
		Me.lblNotes.Visible = True
		Me.lblNotes.AutoSize = False
		Me.lblNotes.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblNotes.Name = "lblNotes"
		Me.lblSeparateTop.Size = New System.Drawing.Size(1333, 41)
		Me.lblSeparateTop.Location = New System.Drawing.Point(-16, 24)
		Me.lblSeparateTop.TabIndex = 67
		Me.lblSeparateTop.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSeparateTop.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSeparateTop.BackColor = System.Drawing.Color.Transparent
		Me.lblSeparateTop.Enabled = True
		Me.lblSeparateTop.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSeparateTop.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSeparateTop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSeparateTop.UseMnemonic = True
		Me.lblSeparateTop.Visible = True
		Me.lblSeparateTop.AutoSize = False
		Me.lblSeparateTop.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblSeparateTop.Name = "lblSeparateTop"
		Me.Controls.Add(chkRepublished)
		Me.Controls.Add(txtJournaTitleShortForm)
		Me.Controls.Add(cmdPreview)
		Me.Controls.Add(lblRecordNumber)
		Me.Controls.Add(chkLibraryCollection)
		Me.Controls.Add(lblArrow2)
		Me.Controls.Add(lblDblClicktoAdd2)
		Me.Controls.Add(lblStatus)
		Me.Controls.Add(cmdDelete)
		Me.Controls.Add(cmdEditJournal)
		Me.Controls.Add(txtCallNumber)
		Me.Controls.Add(cmdNewLargerWork)
		Me.Controls.Add(lblOriginalPublicationDate)
		Me.Controls.Add(lblPublisher)
		Me.Controls.Add(lblCallNumber)
		Me.Controls.Add(lblEditionAndPrinting)
		Me.Controls.Add(lblMiscType)
		Me.Controls.Add(lblLocation)
		Me.Controls.Add(lblThesisDissertationType)
		Me.Controls.Add(lblUnpublishedType)
		Me.Controls.Add(lblUSCCANCitation)
		Me.Controls.Add(lblReportOrDocumentNumber)
		Me.Controls.Add(lblLegislativeHouse)
		Me.Controls.Add(lblNumberOfCongress)
		Me.Controls.Add(lblSessionOfCongress)
		Me.Controls.Add(lblStateLegislativeSession)
		Me.Controls.Add(lblSuDocNumber)
		Me.Controls.Add(lblLegislativeType)
		Me.Controls.Add(lblSeriesVolume)
		Me.Controls.Add(lblTitleOfSeriesIfNotIssuedByAuthor)
		Me.Controls.Add(lblLargerWorkTitle)
		Me.Controls.Add(cmbPagination)
		Me.Controls.Add(lblVolume)
		Me.Controls.Add(lblPublicationMonthOrSeason)
		Me.Controls.Add(lblPage)
		Me.Controls.Add(chkSource)
		Me.Controls.Add(lblPublicationDay)
		Me.Controls.Add(chkYear)
		Me.Controls.Add(lblKeywords)
		Me.Controls.Add(lblJournalTitle)
		Me.Controls.Add(lblSourceType)
		Me.Controls.Add(lblArticleDesignation)
		Me.Controls.Add(lblInputInitials)
		Me.Controls.Add(lblDateUpdated)
		Me.Controls.Add(lblPublicationYear)
		Me.Controls.Add(txtStatus)
		Me.Controls.Add(lstNewKeywords)
		Me.Controls.Add(cmdGetNewKeywords)
		Me.Controls.Add(cmdNewAuthor)
		Me.Controls.Add(cmdNewJournal)
		Me.Controls.Add(txtMiscID)
		Me.Controls.Add(txtUnpublishedID)
		Me.Controls.Add(txtLegislativeID)
		Me.Controls.Add(txtTreatiseID)
		Me.Controls.Add(txtChapterID)
		Me.Controls.Add(txtArticleID)
		Me.Controls.Add(cmbRecordNumber)
		Me.Controls.Add(cmdNextRecord)
		Me.Controls.Add(cmdPreviousRecord)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(lstKeywords)
		Me.Controls.Add(lstCurrentKeywords)
		Me.Controls.Add(cmbAETChoice)
		Me.Controls.Add(txtSuDocNumber)
		Me.Controls.Add(txtLargerWorkID)
		Me.Controls.Add(cmbLargerWorkTitle)
		Me.Controls.Add(txtReportOrDocumentNumber)
		Me.Controls.Add(txtUSCCANCitation)
		Me.Controls.Add(txtStateLegislativeSession)
		Me.Controls.Add(txtSessionOfCongress)
		Me.Controls.Add(txtNumberOfCongress)
		Me.Controls.Add(txtLegislativeHouse)
		Me.Controls.Add(cmbLegislativeType)
		Me.Controls.Add(cmbMiscType)
		Me.Controls.Add(txtLocation)
		Me.Controls.Add(cmbUnpublishedType)
		Me.Controls.Add(cmbThesisDissertationType)
		Me.Controls.Add(chkAllChaptersBySameAuthor)
		Me.Controls.Add(txtTitleOfSeriesIfNotIssuedByAuthor)
		Me.Controls.Add(txtSeriesVolume)
		Me.Controls.Add(txtOriginalPublicationDate)
		Me.Controls.Add(txtPublisher)
		Me.Controls.Add(txtEditionAndPrinting)
		Me.Controls.Add(txtOrganizationIssuingNewsletter)
		Me.Controls.Add(txtNotes)
		Me.Controls.Add(txtPage)
		Me.Controls.Add(cmbPublicationMonthOrSeason)
		Me.Controls.Add(txtVolume)
		Me.Controls.Add(txtPublicationDay)
		Me.Controls.Add(txtJournalID)
		Me.Controls.Add(cmbJournalTitle)
		Me.Controls.Add(cmbArticleDesignation)
		Me.Controls.Add(txtTitle)
		Me.Controls.Add(txtYear)
		Me.Controls.Add(txtInputInitials)
		Me.Controls.Add(txtDateUpdated)
		Me.Controls.Add(txtDateAdded)
		Me.Controls.Add(cmbSourceType)
		Me.Controls.Add(lstAuthors)
		Me.Controls.Add(lstTranslators)
		Me.Controls.Add(lstCurrentAuthors)
		Me.Controls.Add(lstCurrentTranslators)
		Me.Controls.Add(lstCurrentEditors)
		Me.Controls.Add(lstEditors)
		Me.Controls.Add(lblT)
		Me.Controls.Add(lblE)
		Me.Controls.Add(lblA)
		Me.Controls.Add(lblAETChoice)
		Me.Controls.Add(chkKeepSelected)
		Me.Controls.Add(lblDoubleClickToAdd)
		Me.Controls.Add(lblArrow)
		Me.Controls.Add(lblTitle)
		Me.Controls.Add(lblYear)
		Me.Controls.Add(frmEntryInfo)
		Me.Controls.Add(frmRecordInfo)
		Me.Controls.Add(frmCitationInfo)
		Me.Controls.Add(frmAuthorInfo)
		Me.Controls.Add(frmKeywordInfo)
		Me.Controls.Add(frmNotes)
		Me.Controls.Add(lblSeparateBottom)
		Me.Controls.Add(lblMiscID)
		Me.Controls.Add(lblTreatiseID)
		Me.Controls.Add(lblUnpublishedID)
		Me.Controls.Add(lblChapterID)
		Me.Controls.Add(lblArticleID)
		Me.Controls.Add(lblLegisID)
		Me.Controls.Add(tglNewRecords)
		Me.Controls.Add(tglUpdateRecords)
		Me.Controls.Add(tglImportRecords)
		Me.Controls.Add(lblLargerWorkID)
		Me.Controls.Add(lblNotes)
		Me.Controls.Add(lblSeparateTop)
		Me.mneNewAuthor.SetIndex(_mneNewAuthor_3, CType(3, Short))
		Me.mnuAdd.SetIndex(_mnuAdd_2, CType(2, Short))
		Me.mnuFile.SetIndex(_mnuFile_1, CType(1, Short))
		Me.mnuNewJournal.SetIndex(_mnuNewJournal_4, CType(4, Short))
		Me.mnuNewKeyword.SetIndex(_mnuNewKeyword_5, CType(5, Short))
		CType(Me.mnuNewKeyword, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuNewJournal, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuFile, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuAdd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mneNewAuthor, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tglImportRecords, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tglUpdateRecords, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tglNewRecords, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._mnuFile_1, Me._mnuAdd_2})
		_mnuAdd_2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me._mneNewAuthor_3, Me._mnuNewJournal_4, Me._mnuNewKeyword_5})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class