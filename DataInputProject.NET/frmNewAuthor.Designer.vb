<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewAuthor
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
	Public WithEvents txtAETID As System.Windows.Forms.TextBox
	Public WithEvents cmbType As System.Windows.Forms.ComboBox
	Public WithEvents txtSuffix As System.Windows.Forms.TextBox
	Public WithEvents txtLastName As System.Windows.Forms.TextBox
	Public WithEvents txtMiddleName As System.Windows.Forms.TextBox
	Public WithEvents txtFirstName As System.Windows.Forms.TextBox
	Public WithEvents txtInstitutionalEntity As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As AxMicrosoft.Vbe.Interop.Forms.AxCommandButton
	Public WithEvents cmdSave As AxMicrosoft.Vbe.Interop.Forms.AxCommandButton
	Public WithEvents lblType As System.Windows.Forms.Label
	Public WithEvents lblSuffix As System.Windows.Forms.Label
	Public WithEvents lblLastName As System.Windows.Forms.Label
	Public WithEvents lblMiddleName As System.Windows.Forms.Label
	Public WithEvents lblFirstName As System.Windows.Forms.Label
	Public WithEvents lblInstitutionalEntity As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewAuthor))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtAETID = New System.Windows.Forms.TextBox
		Me.cmbType = New System.Windows.Forms.ComboBox
		Me.txtSuffix = New System.Windows.Forms.TextBox
		Me.txtLastName = New System.Windows.Forms.TextBox
		Me.txtMiddleName = New System.Windows.Forms.TextBox
		Me.txtFirstName = New System.Windows.Forms.TextBox
		Me.txtInstitutionalEntity = New System.Windows.Forms.TextBox
		Me.cmdCancel = New AxMicrosoft.Vbe.Interop.Forms.AxCommandButton
		Me.cmdSave = New AxMicrosoft.Vbe.Interop.Forms.AxCommandButton
		Me.lblType = New System.Windows.Forms.Label
		Me.lblSuffix = New System.Windows.Forms.Label
		Me.lblLastName = New System.Windows.Forms.Label
		Me.lblMiddleName = New System.Windows.Forms.Label
		Me.lblFirstName = New System.Windows.Forms.Label
		Me.lblInstitutionalEntity = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdCancel, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdSave, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Add New Author"
		Me.ClientSize = New System.Drawing.Size(601, 233)
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
		Me.Name = "frmNewAuthor"
		Me.txtAETID.AutoSize = False
		Me.txtAETID.Enabled = False
		Me.txtAETID.Size = New System.Drawing.Size(81, 25)
		Me.txtAETID.Location = New System.Drawing.Point(448, 160)
		Me.txtAETID.TabIndex = 14
		Me.txtAETID.Visible = False
		Me.txtAETID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAETID.AcceptsReturn = True
		Me.txtAETID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAETID.BackColor = System.Drawing.SystemColors.Window
		Me.txtAETID.CausesValidation = True
		Me.txtAETID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAETID.HideSelection = True
		Me.txtAETID.ReadOnly = False
		Me.txtAETID.Maxlength = 0
		Me.txtAETID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAETID.MultiLine = False
		Me.txtAETID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAETID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAETID.TabStop = True
		Me.txtAETID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAETID.Name = "txtAETID"
		Me.cmbType.Size = New System.Drawing.Size(169, 21)
		Me.cmbType.Location = New System.Drawing.Point(176, 176)
		Me.cmbType.TabIndex = 4
		Me.cmbType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbType.BackColor = System.Drawing.SystemColors.Window
		Me.cmbType.CausesValidation = True
		Me.cmbType.Enabled = True
		Me.cmbType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbType.IntegralHeight = True
		Me.cmbType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbType.Sorted = False
		Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbType.TabStop = True
		Me.cmbType.Visible = True
		Me.cmbType.Name = "cmbType"
		Me.txtSuffix.AutoSize = False
		Me.txtSuffix.Size = New System.Drawing.Size(169, 19)
		Me.txtSuffix.Location = New System.Drawing.Point(176, 144)
		Me.txtSuffix.TabIndex = 3
		Me.txtSuffix.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSuffix.AcceptsReturn = True
		Me.txtSuffix.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSuffix.BackColor = System.Drawing.SystemColors.Window
		Me.txtSuffix.CausesValidation = True
		Me.txtSuffix.Enabled = True
		Me.txtSuffix.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSuffix.HideSelection = True
		Me.txtSuffix.ReadOnly = False
		Me.txtSuffix.Maxlength = 0
		Me.txtSuffix.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSuffix.MultiLine = False
		Me.txtSuffix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSuffix.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSuffix.TabStop = True
		Me.txtSuffix.Visible = True
		Me.txtSuffix.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSuffix.Name = "txtSuffix"
		Me.txtLastName.AutoSize = False
		Me.txtLastName.Size = New System.Drawing.Size(169, 19)
		Me.txtLastName.Location = New System.Drawing.Point(176, 112)
		Me.txtLastName.TabIndex = 2
		Me.txtLastName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLastName.AcceptsReturn = True
		Me.txtLastName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLastName.BackColor = System.Drawing.SystemColors.Window
		Me.txtLastName.CausesValidation = True
		Me.txtLastName.Enabled = True
		Me.txtLastName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLastName.HideSelection = True
		Me.txtLastName.ReadOnly = False
		Me.txtLastName.Maxlength = 0
		Me.txtLastName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLastName.MultiLine = False
		Me.txtLastName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLastName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLastName.TabStop = True
		Me.txtLastName.Visible = True
		Me.txtLastName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLastName.Name = "txtLastName"
		Me.txtMiddleName.AutoSize = False
		Me.txtMiddleName.Size = New System.Drawing.Size(169, 19)
		Me.txtMiddleName.Location = New System.Drawing.Point(176, 80)
		Me.txtMiddleName.TabIndex = 1
		Me.txtMiddleName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMiddleName.AcceptsReturn = True
		Me.txtMiddleName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMiddleName.BackColor = System.Drawing.SystemColors.Window
		Me.txtMiddleName.CausesValidation = True
		Me.txtMiddleName.Enabled = True
		Me.txtMiddleName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMiddleName.HideSelection = True
		Me.txtMiddleName.ReadOnly = False
		Me.txtMiddleName.Maxlength = 0
		Me.txtMiddleName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMiddleName.MultiLine = False
		Me.txtMiddleName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMiddleName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMiddleName.TabStop = True
		Me.txtMiddleName.Visible = True
		Me.txtMiddleName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMiddleName.Name = "txtMiddleName"
		Me.txtFirstName.AutoSize = False
		Me.txtFirstName.Size = New System.Drawing.Size(169, 19)
		Me.txtFirstName.Location = New System.Drawing.Point(176, 48)
		Me.txtFirstName.TabIndex = 0
		Me.txtFirstName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFirstName.AcceptsReturn = True
		Me.txtFirstName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFirstName.BackColor = System.Drawing.SystemColors.Window
		Me.txtFirstName.CausesValidation = True
		Me.txtFirstName.Enabled = True
		Me.txtFirstName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFirstName.HideSelection = True
		Me.txtFirstName.ReadOnly = False
		Me.txtFirstName.Maxlength = 0
		Me.txtFirstName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFirstName.MultiLine = False
		Me.txtFirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFirstName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFirstName.TabStop = True
		Me.txtFirstName.Visible = True
		Me.txtFirstName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFirstName.Name = "txtFirstName"
		Me.txtInstitutionalEntity.AutoSize = False
		Me.txtInstitutionalEntity.Size = New System.Drawing.Size(169, 19)
		Me.txtInstitutionalEntity.Location = New System.Drawing.Point(176, 16)
		Me.txtInstitutionalEntity.TabIndex = 12
		Me.txtInstitutionalEntity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInstitutionalEntity.AcceptsReturn = True
		Me.txtInstitutionalEntity.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInstitutionalEntity.BackColor = System.Drawing.SystemColors.Window
		Me.txtInstitutionalEntity.CausesValidation = True
		Me.txtInstitutionalEntity.Enabled = True
		Me.txtInstitutionalEntity.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInstitutionalEntity.HideSelection = True
		Me.txtInstitutionalEntity.ReadOnly = False
		Me.txtInstitutionalEntity.Maxlength = 0
		Me.txtInstitutionalEntity.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInstitutionalEntity.MultiLine = False
		Me.txtInstitutionalEntity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInstitutionalEntity.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInstitutionalEntity.TabStop = True
		Me.txtInstitutionalEntity.Visible = True
		Me.txtInstitutionalEntity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInstitutionalEntity.Name = "txtInstitutionalEntity"
		cmdCancel.OcxState = CType(resources.GetObject("cmdCancel.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdCancel.Size = New System.Drawing.Size(97, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(448, 112)
		Me.cmdCancel.TabIndex = 13
		Me.cmdCancel.Name = "cmdCancel"
		cmdSave.OcxState = CType(resources.GetObject("cmdSave.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdSave.Size = New System.Drawing.Size(97, 33)
		Me.cmdSave.Location = New System.Drawing.Point(448, 32)
		Me.cmdSave.TabIndex = 5
		Me.cmdSave.Name = "cmdSave"
		Me.lblType.Text = "Author, Ed., or Trans."
		Me.lblType.Size = New System.Drawing.Size(137, 17)
		Me.lblType.Location = New System.Drawing.Point(8, 176)
		Me.lblType.TabIndex = 11
		Me.lblType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblType.BackColor = System.Drawing.SystemColors.Control
		Me.lblType.Enabled = True
		Me.lblType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblType.UseMnemonic = True
		Me.lblType.Visible = True
		Me.lblType.AutoSize = False
		Me.lblType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblType.Name = "lblType"
		Me.lblSuffix.Text = "Suffix"
		Me.lblSuffix.Size = New System.Drawing.Size(129, 17)
		Me.lblSuffix.Location = New System.Drawing.Point(8, 144)
		Me.lblSuffix.TabIndex = 10
		Me.lblSuffix.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSuffix.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSuffix.BackColor = System.Drawing.SystemColors.Control
		Me.lblSuffix.Enabled = True
		Me.lblSuffix.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSuffix.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSuffix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSuffix.UseMnemonic = True
		Me.lblSuffix.Visible = True
		Me.lblSuffix.AutoSize = False
		Me.lblSuffix.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSuffix.Name = "lblSuffix"
		Me.lblLastName.Text = "Last Name"
		Me.lblLastName.Size = New System.Drawing.Size(129, 17)
		Me.lblLastName.Location = New System.Drawing.Point(8, 112)
		Me.lblLastName.TabIndex = 9
		Me.lblLastName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLastName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLastName.BackColor = System.Drawing.SystemColors.Control
		Me.lblLastName.Enabled = True
		Me.lblLastName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLastName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLastName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLastName.UseMnemonic = True
		Me.lblLastName.Visible = True
		Me.lblLastName.AutoSize = False
		Me.lblLastName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLastName.Name = "lblLastName"
		Me.lblMiddleName.Text = "Middle Name"
		Me.lblMiddleName.Size = New System.Drawing.Size(129, 17)
		Me.lblMiddleName.Location = New System.Drawing.Point(8, 80)
		Me.lblMiddleName.TabIndex = 8
		Me.lblMiddleName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMiddleName.BackColor = System.Drawing.SystemColors.Control
		Me.lblMiddleName.Enabled = True
		Me.lblMiddleName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMiddleName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMiddleName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMiddleName.UseMnemonic = True
		Me.lblMiddleName.Visible = True
		Me.lblMiddleName.AutoSize = False
		Me.lblMiddleName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMiddleName.Name = "lblMiddleName"
		Me.lblFirstName.Text = "First Name"
		Me.lblFirstName.Size = New System.Drawing.Size(129, 17)
		Me.lblFirstName.Location = New System.Drawing.Point(8, 48)
		Me.lblFirstName.TabIndex = 7
		Me.lblFirstName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFirstName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFirstName.BackColor = System.Drawing.SystemColors.Control
		Me.lblFirstName.Enabled = True
		Me.lblFirstName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFirstName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFirstName.UseMnemonic = True
		Me.lblFirstName.Visible = True
		Me.lblFirstName.AutoSize = False
		Me.lblFirstName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFirstName.Name = "lblFirstName"
		Me.lblInstitutionalEntity.Text = "Institutional Entity"
		Me.lblInstitutionalEntity.Size = New System.Drawing.Size(129, 17)
		Me.lblInstitutionalEntity.Location = New System.Drawing.Point(8, 16)
		Me.lblInstitutionalEntity.TabIndex = 6
		Me.lblInstitutionalEntity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInstitutionalEntity.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblInstitutionalEntity.BackColor = System.Drawing.SystemColors.Control
		Me.lblInstitutionalEntity.Enabled = True
		Me.lblInstitutionalEntity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInstitutionalEntity.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInstitutionalEntity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInstitutionalEntity.UseMnemonic = True
		Me.lblInstitutionalEntity.Visible = True
		Me.lblInstitutionalEntity.AutoSize = False
		Me.lblInstitutionalEntity.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInstitutionalEntity.Name = "lblInstitutionalEntity"
		Me.Controls.Add(txtAETID)
		Me.Controls.Add(cmbType)
		Me.Controls.Add(txtSuffix)
		Me.Controls.Add(txtLastName)
		Me.Controls.Add(txtMiddleName)
		Me.Controls.Add(txtFirstName)
		Me.Controls.Add(txtInstitutionalEntity)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(lblType)
		Me.Controls.Add(lblSuffix)
		Me.Controls.Add(lblLastName)
		Me.Controls.Add(lblMiddleName)
		Me.Controls.Add(lblFirstName)
		Me.Controls.Add(lblInstitutionalEntity)
		CType(Me.cmdSave, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdCancel, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class