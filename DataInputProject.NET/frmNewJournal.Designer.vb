<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewJournal
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
	Public WithEvents txtJournalID As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmbPagination As System.Windows.Forms.ComboBox
	Public WithEvents txtPlaceOfPublication As System.Windows.Forms.TextBox
	Public WithEvents txtCallNumber As System.Windows.Forms.TextBox
	Public WithEvents txtNewJournalShortForm As System.Windows.Forms.TextBox
	Public WithEvents txtNewJournal As System.Windows.Forms.TextBox
	Public WithEvents lblPagination As System.Windows.Forms.Label
	Public WithEvents lblPlaceOfPublication As System.Windows.Forms.Label
	Public WithEvents lblCallNumber As System.Windows.Forms.Label
	Public WithEvents lblnewJournalShortForm As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewJournal))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtJournalID = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmbPagination = New System.Windows.Forms.ComboBox
		Me.txtPlaceOfPublication = New System.Windows.Forms.TextBox
		Me.txtCallNumber = New System.Windows.Forms.TextBox
		Me.txtNewJournalShortForm = New System.Windows.Forms.TextBox
		Me.txtNewJournal = New System.Windows.Forms.TextBox
		Me.lblPagination = New System.Windows.Forms.Label
		Me.lblPlaceOfPublication = New System.Windows.Forms.Label
		Me.lblCallNumber = New System.Windows.Forms.Label
		Me.lblnewJournalShortForm = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "New Journal"
		Me.ClientSize = New System.Drawing.Size(766, 197)
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
		Me.Name = "frmNewJournal"
		Me.txtJournalID.AutoSize = False
		Me.txtJournalID.Enabled = False
		Me.txtJournalID.Size = New System.Drawing.Size(81, 33)
		Me.txtJournalID.Location = New System.Drawing.Point(632, 144)
		Me.txtJournalID.TabIndex = 12
		Me.txtJournalID.Visible = False
		Me.txtJournalID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtJournalID.AcceptsReturn = True
		Me.txtJournalID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtJournalID.BackColor = System.Drawing.SystemColors.Window
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(632, 88)
		Me.cmdCancel.TabIndex = 11
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Size = New System.Drawing.Size(81, 33)
		Me.cmdSave.Location = New System.Drawing.Point(632, 24)
		Me.cmdSave.TabIndex = 10
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmbPagination.Size = New System.Drawing.Size(193, 21)
		Me.cmbPagination.Location = New System.Drawing.Point(184, 72)
		Me.cmbPagination.TabIndex = 6
		Me.cmbPagination.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbPagination.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPagination.CausesValidation = True
		Me.cmbPagination.Enabled = True
		Me.cmbPagination.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPagination.IntegralHeight = True
		Me.cmbPagination.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPagination.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPagination.Sorted = False
		Me.cmbPagination.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbPagination.TabStop = True
		Me.cmbPagination.Visible = True
		Me.cmbPagination.Name = "cmbPagination"
		Me.txtPlaceOfPublication.AutoSize = False
		Me.txtPlaceOfPublication.Size = New System.Drawing.Size(193, 19)
		Me.txtPlaceOfPublication.Location = New System.Drawing.Point(184, 136)
		Me.txtPlaceOfPublication.TabIndex = 5
		Me.txtPlaceOfPublication.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPlaceOfPublication.AcceptsReturn = True
		Me.txtPlaceOfPublication.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPlaceOfPublication.BackColor = System.Drawing.SystemColors.Window
		Me.txtPlaceOfPublication.CausesValidation = True
		Me.txtPlaceOfPublication.Enabled = True
		Me.txtPlaceOfPublication.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPlaceOfPublication.HideSelection = True
		Me.txtPlaceOfPublication.ReadOnly = False
		Me.txtPlaceOfPublication.Maxlength = 0
		Me.txtPlaceOfPublication.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPlaceOfPublication.MultiLine = False
		Me.txtPlaceOfPublication.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPlaceOfPublication.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPlaceOfPublication.TabStop = True
		Me.txtPlaceOfPublication.Visible = True
		Me.txtPlaceOfPublication.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPlaceOfPublication.Name = "txtPlaceOfPublication"
		Me.txtCallNumber.AutoSize = False
		Me.txtCallNumber.Size = New System.Drawing.Size(193, 19)
		Me.txtCallNumber.Location = New System.Drawing.Point(184, 104)
		Me.txtCallNumber.TabIndex = 4
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
		Me.txtNewJournalShortForm.AutoSize = False
		Me.txtNewJournalShortForm.Size = New System.Drawing.Size(377, 19)
		Me.txtNewJournalShortForm.Location = New System.Drawing.Point(184, 40)
		Me.txtNewJournalShortForm.TabIndex = 1
		Me.txtNewJournalShortForm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNewJournalShortForm.AcceptsReturn = True
		Me.txtNewJournalShortForm.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNewJournalShortForm.BackColor = System.Drawing.SystemColors.Window
		Me.txtNewJournalShortForm.CausesValidation = True
		Me.txtNewJournalShortForm.Enabled = True
		Me.txtNewJournalShortForm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNewJournalShortForm.HideSelection = True
		Me.txtNewJournalShortForm.ReadOnly = False
		Me.txtNewJournalShortForm.Maxlength = 0
		Me.txtNewJournalShortForm.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNewJournalShortForm.MultiLine = False
		Me.txtNewJournalShortForm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNewJournalShortForm.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNewJournalShortForm.TabStop = True
		Me.txtNewJournalShortForm.Visible = True
		Me.txtNewJournalShortForm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNewJournalShortForm.Name = "txtNewJournalShortForm"
		Me.txtNewJournal.AutoSize = False
		Me.txtNewJournal.Size = New System.Drawing.Size(377, 19)
		Me.txtNewJournal.Location = New System.Drawing.Point(184, 8)
		Me.txtNewJournal.TabIndex = 0
		Me.txtNewJournal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNewJournal.AcceptsReturn = True
		Me.txtNewJournal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNewJournal.BackColor = System.Drawing.SystemColors.Window
		Me.txtNewJournal.CausesValidation = True
		Me.txtNewJournal.Enabled = True
		Me.txtNewJournal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNewJournal.HideSelection = True
		Me.txtNewJournal.ReadOnly = False
		Me.txtNewJournal.Maxlength = 0
		Me.txtNewJournal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNewJournal.MultiLine = False
		Me.txtNewJournal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNewJournal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNewJournal.TabStop = True
		Me.txtNewJournal.Visible = True
		Me.txtNewJournal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNewJournal.Name = "txtNewJournal"
		Me.lblPagination.Text = "Pagination"
		Me.lblPagination.Size = New System.Drawing.Size(137, 17)
		Me.lblPagination.Location = New System.Drawing.Point(8, 72)
		Me.lblPagination.TabIndex = 9
		Me.lblPagination.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPagination.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPagination.BackColor = System.Drawing.SystemColors.Control
		Me.lblPagination.Enabled = True
		Me.lblPagination.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPagination.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPagination.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPagination.UseMnemonic = True
		Me.lblPagination.Visible = True
		Me.lblPagination.AutoSize = False
		Me.lblPagination.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPagination.Name = "lblPagination"
		Me.lblPlaceOfPublication.Text = "Place of Publication"
		Me.lblPlaceOfPublication.Size = New System.Drawing.Size(137, 17)
		Me.lblPlaceOfPublication.Location = New System.Drawing.Point(8, 136)
		Me.lblPlaceOfPublication.TabIndex = 8
		Me.lblPlaceOfPublication.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPlaceOfPublication.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPlaceOfPublication.BackColor = System.Drawing.SystemColors.Control
		Me.lblPlaceOfPublication.Enabled = True
		Me.lblPlaceOfPublication.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPlaceOfPublication.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPlaceOfPublication.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPlaceOfPublication.UseMnemonic = True
		Me.lblPlaceOfPublication.Visible = True
		Me.lblPlaceOfPublication.AutoSize = False
		Me.lblPlaceOfPublication.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPlaceOfPublication.Name = "lblPlaceOfPublication"
		Me.lblCallNumber.Text = "Call Number"
		Me.lblCallNumber.Size = New System.Drawing.Size(137, 17)
		Me.lblCallNumber.Location = New System.Drawing.Point(8, 104)
		Me.lblCallNumber.TabIndex = 7
		Me.lblCallNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCallNumber.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCallNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblCallNumber.Enabled = True
		Me.lblCallNumber.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCallNumber.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCallNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCallNumber.UseMnemonic = True
		Me.lblCallNumber.Visible = True
		Me.lblCallNumber.AutoSize = False
		Me.lblCallNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCallNumber.Name = "lblCallNumber"
		Me.lblnewJournalShortForm.Text = "Journal Short Form"
		Me.lblnewJournalShortForm.Size = New System.Drawing.Size(129, 17)
		Me.lblnewJournalShortForm.Location = New System.Drawing.Point(8, 40)
		Me.lblnewJournalShortForm.TabIndex = 3
		Me.lblnewJournalShortForm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblnewJournalShortForm.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblnewJournalShortForm.BackColor = System.Drawing.SystemColors.Control
		Me.lblnewJournalShortForm.Enabled = True
		Me.lblnewJournalShortForm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblnewJournalShortForm.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblnewJournalShortForm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblnewJournalShortForm.UseMnemonic = True
		Me.lblnewJournalShortForm.Visible = True
		Me.lblnewJournalShortForm.AutoSize = False
		Me.lblnewJournalShortForm.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblnewJournalShortForm.Name = "lblnewJournalShortForm"
		Me.Label1.Text = "Journal Name"
		Me.Label1.Size = New System.Drawing.Size(121, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 2
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(txtJournalID)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmbPagination)
		Me.Controls.Add(txtPlaceOfPublication)
		Me.Controls.Add(txtCallNumber)
		Me.Controls.Add(txtNewJournalShortForm)
		Me.Controls.Add(txtNewJournal)
		Me.Controls.Add(lblPagination)
		Me.Controls.Add(lblPlaceOfPublication)
		Me.Controls.Add(lblCallNumber)
		Me.Controls.Add(lblnewJournalShortForm)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class