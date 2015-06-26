<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmJump
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
	Public WithEvents cmdJump As System.Windows.Forms.Button
	Public WithEvents txtRecNum As System.Windows.Forms.TextBox
	Public WithEvents lblJump As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmJump))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdJump = New System.Windows.Forms.Button
		Me.txtRecNum = New System.Windows.Forms.TextBox
		Me.lblJump = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(345, 86)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
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
		Me.Name = "frmJump"
		Me.cmdJump.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdJump.Text = "Jump"
		Me.cmdJump.Size = New System.Drawing.Size(81, 33)
		Me.cmdJump.Location = New System.Drawing.Point(248, 24)
		Me.cmdJump.TabIndex = 2
		Me.cmdJump.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdJump.BackColor = System.Drawing.SystemColors.Control
		Me.cmdJump.CausesValidation = True
		Me.cmdJump.Enabled = True
		Me.cmdJump.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdJump.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdJump.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdJump.TabStop = True
		Me.cmdJump.Name = "cmdJump"
		Me.txtRecNum.AutoSize = False
		Me.txtRecNum.Size = New System.Drawing.Size(89, 19)
		Me.txtRecNum.Location = New System.Drawing.Point(152, 32)
		Me.txtRecNum.TabIndex = 1
		Me.txtRecNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRecNum.AcceptsReturn = True
		Me.txtRecNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRecNum.BackColor = System.Drawing.SystemColors.Window
		Me.txtRecNum.CausesValidation = True
		Me.txtRecNum.Enabled = True
		Me.txtRecNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRecNum.HideSelection = True
		Me.txtRecNum.ReadOnly = False
		Me.txtRecNum.Maxlength = 0
		Me.txtRecNum.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRecNum.MultiLine = False
		Me.txtRecNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRecNum.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRecNum.TabStop = True
		Me.txtRecNum.Visible = True
		Me.txtRecNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRecNum.Name = "txtRecNum"
		Me.lblJump.Text = "Jump to Record Number:"
		Me.lblJump.Size = New System.Drawing.Size(121, 17)
		Me.lblJump.Location = New System.Drawing.Point(24, 32)
		Me.lblJump.TabIndex = 0
		Me.lblJump.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblJump.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblJump.BackColor = System.Drawing.SystemColors.Control
		Me.lblJump.Enabled = True
		Me.lblJump.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblJump.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblJump.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblJump.UseMnemonic = True
		Me.lblJump.Visible = True
		Me.lblJump.AutoSize = False
		Me.lblJump.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblJump.Name = "lblJump"
		Me.Controls.Add(cmdJump)
		Me.Controls.Add(txtRecNum)
		Me.Controls.Add(lblJump)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class