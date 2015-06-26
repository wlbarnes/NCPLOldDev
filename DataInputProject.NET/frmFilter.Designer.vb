<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFilter
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
	Public WithEvents cmdRecordID As System.Windows.Forms.Button
	Public WithEvents cmdConvert As System.Windows.Forms.Button
	Public WithEvents cmdClear As System.Windows.Forms.Button
	Public WithEvents cmdSQL As System.Windows.Forms.Button
	Public WithEvents cmdStandard As System.Windows.Forms.Button
	Public WithEvents txtQuery As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFilter))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdRecordID = New System.Windows.Forms.Button
		Me.cmdConvert = New System.Windows.Forms.Button
		Me.cmdClear = New System.Windows.Forms.Button
		Me.cmdSQL = New System.Windows.Forms.Button
		Me.cmdStandard = New System.Windows.Forms.Button
		Me.txtQuery = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Filter Records"
		Me.ClientSize = New System.Drawing.Size(678, 263)
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
		Me.Name = "frmFilter"
		Me.cmdRecordID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRecordID.Text = "RecordID"
		Me.cmdRecordID.Size = New System.Drawing.Size(97, 33)
		Me.cmdRecordID.Location = New System.Drawing.Point(560, 24)
		Me.cmdRecordID.TabIndex = 5
		Me.cmdRecordID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRecordID.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRecordID.CausesValidation = True
		Me.cmdRecordID.Enabled = True
		Me.cmdRecordID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRecordID.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRecordID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRecordID.TabStop = True
		Me.cmdRecordID.Name = "cmdRecordID"
		Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdConvert.Text = "Convert Standard to SQL Query"
		Me.cmdConvert.Size = New System.Drawing.Size(97, 33)
		Me.cmdConvert.Location = New System.Drawing.Point(560, 120)
		Me.cmdConvert.TabIndex = 4
		Me.cmdConvert.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdConvert.BackColor = System.Drawing.SystemColors.Control
		Me.cmdConvert.CausesValidation = True
		Me.cmdConvert.Enabled = True
		Me.cmdConvert.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdConvert.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdConvert.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdConvert.TabStop = True
		Me.cmdConvert.Name = "cmdConvert"
		Me.cmdClear.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClear.Text = "Clear Form"
		Me.cmdClear.Size = New System.Drawing.Size(97, 33)
		Me.cmdClear.Location = New System.Drawing.Point(560, 216)
		Me.cmdClear.TabIndex = 3
		Me.cmdClear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClear.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClear.CausesValidation = True
		Me.cmdClear.Enabled = True
		Me.cmdClear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClear.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClear.TabStop = True
		Me.cmdClear.Name = "cmdClear"
		Me.cmdSQL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSQL.Text = "Execute SQL Query"
		Me.cmdSQL.Size = New System.Drawing.Size(97, 33)
		Me.cmdSQL.Location = New System.Drawing.Point(560, 168)
		Me.cmdSQL.TabIndex = 2
		Me.cmdSQL.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSQL.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSQL.CausesValidation = True
		Me.cmdSQL.Enabled = True
		Me.cmdSQL.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSQL.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSQL.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSQL.TabStop = True
		Me.cmdSQL.Name = "cmdSQL"
		Me.cmdStandard.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdStandard.Text = "Execute Standard Query"
		Me.cmdStandard.Size = New System.Drawing.Size(97, 33)
		Me.cmdStandard.Location = New System.Drawing.Point(560, 72)
		Me.cmdStandard.TabIndex = 1
		Me.cmdStandard.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdStandard.BackColor = System.Drawing.SystemColors.Control
		Me.cmdStandard.CausesValidation = True
		Me.cmdStandard.Enabled = True
		Me.cmdStandard.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdStandard.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdStandard.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdStandard.TabStop = True
		Me.cmdStandard.Name = "cmdStandard"
		Me.txtQuery.AutoSize = False
		Me.txtQuery.Size = New System.Drawing.Size(497, 209)
		Me.txtQuery.Location = New System.Drawing.Point(24, 24)
		Me.txtQuery.MultiLine = True
		Me.txtQuery.TabIndex = 0
		Me.txtQuery.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtQuery.AcceptsReturn = True
		Me.txtQuery.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtQuery.BackColor = System.Drawing.SystemColors.Window
		Me.txtQuery.CausesValidation = True
		Me.txtQuery.Enabled = True
		Me.txtQuery.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtQuery.HideSelection = True
		Me.txtQuery.ReadOnly = False
		Me.txtQuery.Maxlength = 0
		Me.txtQuery.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtQuery.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtQuery.TabStop = True
		Me.txtQuery.Visible = True
		Me.txtQuery.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtQuery.Name = "txtQuery"
		Me.Controls.Add(cmdRecordID)
		Me.Controls.Add(cmdConvert)
		Me.Controls.Add(cmdClear)
		Me.Controls.Add(cmdSQL)
		Me.Controls.Add(cmdStandard)
		Me.Controls.Add(txtQuery)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class