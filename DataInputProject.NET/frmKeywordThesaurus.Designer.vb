<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmKeywordThesaurus
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
	Public WithEvents cmdModifyThesaurus As System.Windows.Forms.Button
	Public WithEvents cmdModifyKeyword As System.Windows.Forms.Button
	Public WithEvents cmdDeleteThesaurus As System.Windows.Forms.Button
	Public WithEvents cmdDeleteKeyword As System.Windows.Forms.Button
	Public WithEvents txtKeywordID As System.Windows.Forms.TextBox
	Public WithEvents cmdAddThesaurus As System.Windows.Forms.Button
	Public WithEvents cmdAddKeyword As System.Windows.Forms.Button
	Public WithEvents lstThesaurus As System.Windows.Forms.ListBox
	Public WithEvents lstKeywords As System.Windows.Forms.ListBox
	Public WithEvents lblKeywords As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmKeywordThesaurus))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdModifyThesaurus = New System.Windows.Forms.Button
		Me.cmdModifyKeyword = New System.Windows.Forms.Button
		Me.cmdDeleteThesaurus = New System.Windows.Forms.Button
		Me.cmdDeleteKeyword = New System.Windows.Forms.Button
		Me.txtKeywordID = New System.Windows.Forms.TextBox
		Me.cmdAddThesaurus = New System.Windows.Forms.Button
		Me.cmdAddKeyword = New System.Windows.Forms.Button
		Me.lstThesaurus = New System.Windows.Forms.ListBox
		Me.lstKeywords = New System.Windows.Forms.ListBox
		Me.lblKeywords = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Modify Keywords and Thesaurus Equivalents"
		Me.ClientSize = New System.Drawing.Size(673, 420)
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
		Me.Name = "frmKeywordThesaurus"
		Me.cmdModifyThesaurus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdModifyThesaurus.Text = "Modify Thesaurus Equivalent for Selected Keyword"
		Me.cmdModifyThesaurus.Size = New System.Drawing.Size(153, 33)
		Me.cmdModifyThesaurus.Location = New System.Drawing.Point(408, 336)
		Me.cmdModifyThesaurus.TabIndex = 9
		Me.cmdModifyThesaurus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdModifyThesaurus.BackColor = System.Drawing.SystemColors.Control
		Me.cmdModifyThesaurus.CausesValidation = True
		Me.cmdModifyThesaurus.Enabled = True
		Me.cmdModifyThesaurus.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdModifyThesaurus.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdModifyThesaurus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdModifyThesaurus.TabStop = True
		Me.cmdModifyThesaurus.Name = "cmdModifyThesaurus"
		Me.cmdModifyKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdModifyKeyword.Text = "Modify Keyword"
		Me.cmdModifyKeyword.Size = New System.Drawing.Size(105, 33)
		Me.cmdModifyKeyword.Location = New System.Drawing.Point(96, 336)
		Me.cmdModifyKeyword.TabIndex = 8
		Me.cmdModifyKeyword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdModifyKeyword.BackColor = System.Drawing.SystemColors.Control
		Me.cmdModifyKeyword.CausesValidation = True
		Me.cmdModifyKeyword.Enabled = True
		Me.cmdModifyKeyword.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdModifyKeyword.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdModifyKeyword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdModifyKeyword.TabStop = True
		Me.cmdModifyKeyword.Name = "cmdModifyKeyword"
		Me.cmdDeleteThesaurus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDeleteThesaurus.Text = "Delete Thesaurus Equivalent for Selected Keyword"
		Me.cmdDeleteThesaurus.Size = New System.Drawing.Size(153, 33)
		Me.cmdDeleteThesaurus.Location = New System.Drawing.Point(408, 296)
		Me.cmdDeleteThesaurus.TabIndex = 7
		Me.cmdDeleteThesaurus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDeleteThesaurus.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDeleteThesaurus.CausesValidation = True
		Me.cmdDeleteThesaurus.Enabled = True
		Me.cmdDeleteThesaurus.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDeleteThesaurus.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDeleteThesaurus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDeleteThesaurus.TabStop = True
		Me.cmdDeleteThesaurus.Name = "cmdDeleteThesaurus"
		Me.cmdDeleteKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDeleteKeyword.Text = "Delete Keyword"
		Me.cmdDeleteKeyword.Size = New System.Drawing.Size(105, 33)
		Me.cmdDeleteKeyword.Location = New System.Drawing.Point(96, 296)
		Me.cmdDeleteKeyword.TabIndex = 6
		Me.cmdDeleteKeyword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDeleteKeyword.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDeleteKeyword.CausesValidation = True
		Me.cmdDeleteKeyword.Enabled = True
		Me.cmdDeleteKeyword.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDeleteKeyword.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDeleteKeyword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDeleteKeyword.TabStop = True
		Me.cmdDeleteKeyword.Name = "cmdDeleteKeyword"
		Me.txtKeywordID.AutoSize = False
		Me.txtKeywordID.BackColor = System.Drawing.SystemColors.InactiveCaptionText
		Me.txtKeywordID.Enabled = False
		Me.txtKeywordID.Size = New System.Drawing.Size(81, 25)
		Me.txtKeywordID.Location = New System.Drawing.Point(88, 32)
		Me.txtKeywordID.TabIndex = 5
		Me.txtKeywordID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtKeywordID.AcceptsReturn = True
		Me.txtKeywordID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtKeywordID.CausesValidation = True
		Me.txtKeywordID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtKeywordID.HideSelection = True
		Me.txtKeywordID.ReadOnly = False
		Me.txtKeywordID.Maxlength = 0
		Me.txtKeywordID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtKeywordID.MultiLine = False
		Me.txtKeywordID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtKeywordID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtKeywordID.TabStop = True
		Me.txtKeywordID.Visible = True
		Me.txtKeywordID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtKeywordID.Name = "txtKeywordID"
		Me.cmdAddThesaurus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAddThesaurus.Text = "Add Thesaurus Equivalent for Selected Keyword"
		Me.cmdAddThesaurus.Size = New System.Drawing.Size(153, 33)
		Me.cmdAddThesaurus.Location = New System.Drawing.Point(408, 256)
		Me.cmdAddThesaurus.TabIndex = 4
		Me.cmdAddThesaurus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAddThesaurus.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAddThesaurus.CausesValidation = True
		Me.cmdAddThesaurus.Enabled = True
		Me.cmdAddThesaurus.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAddThesaurus.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAddThesaurus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAddThesaurus.TabStop = True
		Me.cmdAddThesaurus.Name = "cmdAddThesaurus"
		Me.cmdAddKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAddKeyword.Text = "Add Keyword"
		Me.cmdAddKeyword.Size = New System.Drawing.Size(105, 33)
		Me.cmdAddKeyword.Location = New System.Drawing.Point(96, 256)
		Me.cmdAddKeyword.TabIndex = 3
		Me.cmdAddKeyword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAddKeyword.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAddKeyword.CausesValidation = True
		Me.cmdAddKeyword.Enabled = True
		Me.cmdAddKeyword.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAddKeyword.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAddKeyword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAddKeyword.TabStop = True
		Me.cmdAddKeyword.Name = "cmdAddKeyword"
		Me.lstThesaurus.Size = New System.Drawing.Size(281, 189)
		Me.lstThesaurus.Location = New System.Drawing.Point(336, 64)
		Me.lstThesaurus.Sorted = True
		Me.lstThesaurus.TabIndex = 2
		Me.lstThesaurus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstThesaurus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstThesaurus.BackColor = System.Drawing.SystemColors.Window
		Me.lstThesaurus.CausesValidation = True
		Me.lstThesaurus.Enabled = True
		Me.lstThesaurus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstThesaurus.IntegralHeight = True
		Me.lstThesaurus.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstThesaurus.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstThesaurus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstThesaurus.TabStop = True
		Me.lstThesaurus.Visible = True
		Me.lstThesaurus.MultiColumn = False
		Me.lstThesaurus.Name = "lstThesaurus"
		Me.lstKeywords.Size = New System.Drawing.Size(281, 189)
		Me.lstKeywords.Location = New System.Drawing.Point(8, 64)
		Me.lstKeywords.Sorted = True
		Me.lstKeywords.TabIndex = 0
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
		Me.lblKeywords.Text = "Keywords"
		Me.lblKeywords.Size = New System.Drawing.Size(57, 17)
		Me.lblKeywords.Location = New System.Drawing.Point(8, 40)
		Me.lblKeywords.TabIndex = 1
		Me.lblKeywords.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblKeywords.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblKeywords.BackColor = System.Drawing.SystemColors.Control
		Me.lblKeywords.Enabled = True
		Me.lblKeywords.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblKeywords.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblKeywords.UseMnemonic = True
		Me.lblKeywords.Visible = True
		Me.lblKeywords.AutoSize = False
		Me.lblKeywords.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblKeywords.Name = "lblKeywords"
		Me.Controls.Add(cmdModifyThesaurus)
		Me.Controls.Add(cmdModifyKeyword)
		Me.Controls.Add(cmdDeleteThesaurus)
		Me.Controls.Add(cmdDeleteKeyword)
		Me.Controls.Add(txtKeywordID)
		Me.Controls.Add(cmdAddThesaurus)
		Me.Controls.Add(cmdAddKeyword)
		Me.Controls.Add(lstThesaurus)
		Me.Controls.Add(lstKeywords)
		Me.Controls.Add(lblKeywords)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class