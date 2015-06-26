Option Strict Off
Option Explicit On
Friend Class frmWordPreview
	Inherits System.Windows.Forms.Form
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		
		Me.Close()
	End Sub
	
	Private Sub frmWordPreview_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim Bopen As Boolean
		Dim iCount As Short
		Dim i As Short
		
		
		'iCount = frmWordPreview.OLEWord.object.Application.Documents.Count
		'If iCount > 2 Then Bopen = True
		
		'For i = 1 To iCount
		'    If frmWordPreview.OLEWord.object.Application.Documents(iCount) = "Document in Unnamed" Then frmWordPreview.OLEWord.object.Application.Documents(iCount).Close
		'Next
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmWordPreview.OLEWord.object.Application. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Me.OLEWord.object.Application.Documents(1).Close()
		'If Not Bopen Then
		'UPGRADE_WARNING: Couldn't resolve default property of object frmWordPreview.OLEWord.object.Application. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Me.OLEWord.object.Application.Application.Quit()
		
	End Sub
End Class