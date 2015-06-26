Option Strict Off
Option Explicit On
Friend Class frmJump
	Inherits System.Windows.Forms.Form
	Private Sub cmdJump_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdJump.Click
		On Error GoTo Err_Renamed
		frmMain.cmbRecordNumber.Text = Me.txtRecNum.Text
		Me.Close()
Err_Renamed: 
		Select Case Err.Number
			Case 0
			Case Else
				MsgBox("Record does not exist")
				Me.Close()
		End Select
		Exit Sub
	End Sub
End Class