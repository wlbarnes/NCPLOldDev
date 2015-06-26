<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewLargerWork
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
	Public WithEvents chkAllChaptersBySameAuthor As System.Windows.Forms.CheckBox
	Public WithEvents txtTitleOfSeriesIfNotIssuedByAuthor As System.Windows.Forms.TextBox
	Public WithEvents txtSeriesVolume As System.Windows.Forms.TextBox
	Public WithEvents txtOriginalPublicationDate As System.Windows.Forms.TextBox
	Public WithEvents txtPublisher As System.Windows.Forms.TextBox
	Public WithEvents txtCallNumber As System.Windows.Forms.TextBox
	Public WithEvents lblOriginalPublicationDate As System.Windows.Forms.TextBox
	Public WithEvents lblSeriesVolume As System.Windows.Forms.TextBox
	Public WithEvents lblPublisher As System.Windows.Forms.TextBox
	Public WithEvents lblTitleOfSeriesIfNotIssuedByAuthor As System.Windows.Forms.TextBox
	Public WithEvents lblCallNumber As System.Windows.Forms.TextBox
	Public WithEvents txtLargerWorkID As System.Windows.Forms.TextBox
	Public WithEvents txtLargerWorkTitle As System.Windows.Forms.TextBox
	Public WithEvents txtEditionandPrinting As System.Windows.Forms.TextBox
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents lblLargerWorkName As System.Windows.Forms.Label
	Public WithEvents lblEditionandPrinting As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewLargerWork))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkAllChaptersBySameAuthor = New System.Windows.Forms.CheckBox
		Me.txtTitleOfSeriesIfNotIssuedByAuthor = New System.Windows.Forms.TextBox
		Me.txtSeriesVolume = New System.Windows.Forms.TextBox
		Me.txtOriginalPublicationDate = New System.Windows.Forms.TextBox
		Me.txtPublisher = New System.Windows.Forms.TextBox
		Me.txtCallNumber = New System.Windows.Forms.TextBox
		Me.lblOriginalPublicationDate = New System.Windows.Forms.TextBox
		Me.lblSeriesVolume = New System.Windows.Forms.TextBox
		Me.lblPublisher = New System.Windows.Forms.TextBox
		Me.lblTitleOfSeriesIfNotIssuedByAuthor = New System.Windows.Forms.TextBox
		Me.lblCallNumber = New System.Windows.Forms.TextBox
		Me.txtLargerWorkID = New System.Windows.Forms.TextBox
		Me.txtLargerWorkTitle = New System.Windows.Forms.TextBox
		Me.txtEditionandPrinting = New System.Windows.Forms.TextBox
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.lblLargerWorkName = New System.Windows.Forms.Label
		Me.lblEditionandPrinting = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(831, 352)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmNewLargerWork"
		Me.chkAllChaptersBySameAuthor.Text = "All Chapters By Same Author?"
		Me.chkAllChaptersBySameAuthor.Size = New System.Drawing.Size(185, 33)
		Me.chkAllChaptersBySameAuthor.Location = New System.Drawing.Point(192, 296)
		Me.chkAllChaptersBySameAuthor.TabIndex = 17
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
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Size = New System.Drawing.Size(497, 19)
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.Location = New System.Drawing.Point(192, 224)
		Me.txtTitleOfSeriesIfNotIssuedByAuthor.TabIndex = 16
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
		Me.txtSeriesVolume.Location = New System.Drawing.Point(192, 184)
		Me.txtSeriesVolume.TabIndex = 15
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
		Me.txtOriginalPublicationDate.Size = New System.Drawing.Size(121, 19)
		Me.txtOriginalPublicationDate.Location = New System.Drawing.Point(192, 144)
		Me.txtOriginalPublicationDate.TabIndex = 14
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
		Me.txtPublisher.Location = New System.Drawing.Point(192, 104)
		Me.txtPublisher.TabIndex = 13
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
		Me.txtCallNumber.AutoSize = False
		Me.txtCallNumber.Size = New System.Drawing.Size(121, 19)
		Me.txtCallNumber.Location = New System.Drawing.Point(192, 264)
		Me.txtCallNumber.TabIndex = 12
		Me.txtCallNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCallNumber.AcceptsReturn = True
		Me.txtCallNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCallNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtCallNumber.CausesValidation = True
		Me.txtCallNumber.Enabled = True
		Me.txtCallNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCallNumber.HideSelection = True
		Me.txtCallNumber.ReadOnly = False
		Me.txtCallNumber.Maxlength = 0
		Me.txtCallNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCallNumber.MultiLine = False
		Me.txtCallNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCallNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCallNumber.TabStop = True
		Me.txtCallNumber.Visible = True
		Me.txtCallNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCallNumber.Name = "txtCallNumber"
		Me.lblOriginalPublicationDate.AutoSize = False
		Me.lblOriginalPublicationDate.BackColor = System.Drawing.SystemColors.Control
		Me.lblOriginalPublicationDate.Enabled = False
		Me.lblOriginalPublicationDate.Size = New System.Drawing.Size(121, 13)
		Me.lblOriginalPublicationDate.Location = New System.Drawing.Point(16, 144)
		Me.lblOriginalPublicationDate.TabIndex = 11
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
		Me.lblSeriesVolume.AutoSize = False
		Me.lblSeriesVolume.BackColor = System.Drawing.SystemColors.Control
		Me.lblSeriesVolume.Enabled = False
		Me.lblSeriesVolume.Size = New System.Drawing.Size(81, 13)
		Me.lblSeriesVolume.Location = New System.Drawing.Point(16, 184)
		Me.lblSeriesVolume.TabIndex = 10
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
		Me.lblPublisher.AutoSize = False
		Me.lblPublisher.BackColor = System.Drawing.SystemColors.Control
		Me.lblPublisher.Enabled = False
		Me.lblPublisher.Size = New System.Drawing.Size(89, 13)
		Me.lblPublisher.Location = New System.Drawing.Point(16, 104)
		Me.lblPublisher.TabIndex = 9
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
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.AutoSize = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Enabled = False
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Size = New System.Drawing.Size(129, 37)
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Location = New System.Drawing.Point(16, 216)
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.MultiLine = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.TabIndex = 8
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
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.TabStop = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Visible = True
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitleOfSeriesIfNotIssuedByAuthor.Name = "lblTitleOfSeriesIfNotIssuedByAuthor"
		Me.lblCallNumber.AutoSize = False
		Me.lblCallNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblCallNumber.Enabled = False
		Me.lblCallNumber.Size = New System.Drawing.Size(81, 13)
		Me.lblCallNumber.Location = New System.Drawing.Point(16, 272)
		Me.lblCallNumber.TabIndex = 7
		Me.lblCallNumber.Text = "Call Number"
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
		Me.lblCallNumber.Visible = True
		Me.lblCallNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCallNumber.Name = "lblCallNumber"
		Me.txtLargerWorkID.AutoSize = False
		Me.txtLargerWorkID.Enabled = False
		Me.txtLargerWorkID.Size = New System.Drawing.Size(81, 33)
		Me.txtLargerWorkID.Location = New System.Drawing.Point(704, 160)
		Me.txtLargerWorkID.TabIndex = 6
		Me.txtLargerWorkID.Visible = False
		Me.txtLargerWorkID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLargerWorkID.AcceptsReturn = True
		Me.txtLargerWorkID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLargerWorkID.BackColor = System.Drawing.SystemColors.Window
		Me.txtLargerWorkID.CausesValidation = True
		Me.txtLargerWorkID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLargerWorkID.HideSelection = True
		Me.txtLargerWorkID.ReadOnly = False
		Me.txtLargerWorkID.Maxlength = 0
		Me.txtLargerWorkID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLargerWorkID.MultiLine = False
		Me.txtLargerWorkID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLargerWorkID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLargerWorkID.TabStop = True
		Me.txtLargerWorkID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLargerWorkID.Name = "txtLargerWorkID"
		Me.txtLargerWorkTitle.AutoSize = False
		Me.txtLargerWorkTitle.Size = New System.Drawing.Size(489, 19)
		Me.txtLargerWorkTitle.Location = New System.Drawing.Point(192, 32)
		Me.txtLargerWorkTitle.TabIndex = 3
		Me.txtLargerWorkTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLargerWorkTitle.AcceptsReturn = True
		Me.txtLargerWorkTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLargerWorkTitle.BackColor = System.Drawing.SystemColors.Window
		Me.txtLargerWorkTitle.CausesValidation = True
		Me.txtLargerWorkTitle.Enabled = True
		Me.txtLargerWorkTitle.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLargerWorkTitle.HideSelection = True
		Me.txtLargerWorkTitle.ReadOnly = False
		Me.txtLargerWorkTitle.Maxlength = 0
		Me.txtLargerWorkTitle.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLargerWorkTitle.MultiLine = False
		Me.txtLargerWorkTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLargerWorkTitle.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLargerWorkTitle.TabStop = True
		Me.txtLargerWorkTitle.Visible = True
		Me.txtLargerWorkTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLargerWorkTitle.Name = "txtLargerWorkTitle"
		Me.txtEditionandPrinting.AutoSize = False
		Me.txtEditionandPrinting.Size = New System.Drawing.Size(129, 19)
		Me.txtEditionandPrinting.Location = New System.Drawing.Point(192, 64)
		Me.txtEditionandPrinting.TabIndex = 2
		Me.txtEditionandPrinting.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEditionandPrinting.AcceptsReturn = True
		Me.txtEditionandPrinting.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEditionandPrinting.BackColor = System.Drawing.SystemColors.Window
		Me.txtEditionandPrinting.CausesValidation = True
		Me.txtEditionandPrinting.Enabled = True
		Me.txtEditionandPrinting.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEditionandPrinting.HideSelection = True
		Me.txtEditionandPrinting.ReadOnly = False
		Me.txtEditionandPrinting.Maxlength = 0
		Me.txtEditionandPrinting.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEditionandPrinting.MultiLine = False
		Me.txtEditionandPrinting.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEditionandPrinting.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEditionandPrinting.TabStop = True
		Me.txtEditionandPrinting.Visible = True
		Me.txtEditionandPrinting.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEditionandPrinting.Name = "txtEditionandPrinting"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Size = New System.Drawing.Size(81, 33)
		Me.cmdSave.Location = New System.Drawing.Point(704, 48)
		Me.cmdSave.TabIndex = 1
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(704, 112)
		Me.cmdCancel.TabIndex = 0
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.lblLargerWorkName.Text = "Larger Work Name"
		Me.lblLargerWorkName.Size = New System.Drawing.Size(121, 17)
		Me.lblLargerWorkName.Location = New System.Drawing.Point(16, 32)
		Me.lblLargerWorkName.TabIndex = 5
		Me.lblLargerWorkName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLargerWorkName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLargerWorkName.BackColor = System.Drawing.SystemColors.Control
		Me.lblLargerWorkName.Enabled = True
		Me.lblLargerWorkName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLargerWorkName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLargerWorkName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLargerWorkName.UseMnemonic = True
		Me.lblLargerWorkName.Visible = True
		Me.lblLargerWorkName.AutoSize = False
		Me.lblLargerWorkName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLargerWorkName.Name = "lblLargerWorkName"
		Me.lblEditionandPrinting.Text = "Edition and Printing"
		Me.lblEditionandPrinting.Size = New System.Drawing.Size(129, 17)
		Me.lblEditionandPrinting.Location = New System.Drawing.Point(16, 64)
		Me.lblEditionandPrinting.TabIndex = 4
		Me.lblEditionandPrinting.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEditionandPrinting.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblEditionandPrinting.BackColor = System.Drawing.SystemColors.Control
		Me.lblEditionandPrinting.Enabled = True
		Me.lblEditionandPrinting.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEditionandPrinting.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEditionandPrinting.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEditionandPrinting.UseMnemonic = True
		Me.lblEditionandPrinting.Visible = True
		Me.lblEditionandPrinting.AutoSize = False
		Me.lblEditionandPrinting.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblEditionandPrinting.Name = "lblEditionandPrinting"
		Me.Controls.Add(chkAllChaptersBySameAuthor)
		Me.Controls.Add(txtTitleOfSeriesIfNotIssuedByAuthor)
		Me.Controls.Add(txtSeriesVolume)
		Me.Controls.Add(txtOriginalPublicationDate)
		Me.Controls.Add(txtPublisher)
		Me.Controls.Add(txtCallNumber)
		Me.Controls.Add(lblOriginalPublicationDate)
		Me.Controls.Add(lblSeriesVolume)
		Me.Controls.Add(lblPublisher)
		Me.Controls.Add(lblTitleOfSeriesIfNotIssuedByAuthor)
		Me.Controls.Add(lblCallNumber)
		Me.Controls.Add(txtLargerWorkID)
		Me.Controls.Add(txtLargerWorkTitle)
		Me.Controls.Add(txtEditionandPrinting)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(lblLargerWorkName)
		Me.Controls.Add(lblEditionandPrinting)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class