Option Strict Off
Option Explicit On
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Public rstJournals As ADODB.Recordset
	Public rstAuthors As ADODB.Recordset
	Public rstEditors As ADODB.Recordset
	Public rstTranslators As ADODB.Recordset
	Dim rstArticles As ADODB.Recordset
	Dim rstChapters As ADODB.Recordset
	Dim rstMisc As ADODB.Recordset
	Dim rstLegislativeMaterial As ADODB.Recordset
	Public rstRecords As ADODB.Recordset
	Dim rstRecordsAET As ADODB.Recordset
	Dim rstRecordsKeywords As ADODB.Recordset
	Dim rstTreatises As ADODB.Recordset
	Dim rstUnpublishedWork As ADODB.Recordset
	
	Public rstLargerWorks As ADODB.Recordset
	Public rstKeywords As ADODB.Recordset
	Public cnReadDatabase As ADODB.Connection
	Public cnWriteDatabase As ADODB.Connection
	Dim iSaveListIndex As Short
	Public cAuthors As Collection
	Public cEditors As Collection
	Public cTranslators As Collection
	Public cKeywords As Collection
	
	Private Sub Position_Article_Form()
		With Me.lblVolume
            'UPGRADE_ISSUE: TextBox property lblVolume'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(8280)
            '.TabIndex = 0
            .Text = "Volume"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(735)
            .Visible = True
        End With
        With Me.lblPage
            'UPGRADE_ISSUE: TextBox property lblPage'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10440)
            '.TabIndex = 2
            .Text = "Page Number"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblPublicationMonthOrSeason
            'UPGRADE_ISSUE: TextBox property lblPublicationMonthOrSeason'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3240)
            '.TabIndex = 1
            .Text = "Publication Month or Season"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(2055)
            .Visible = True
        End With

        With Me.lblPublicationDay
            'UPGRADE_ISSUE: TextBox property lblPublicationDay'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(6240)
            '.TabIndex = 3
            .Text = "Publication Day"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblJournalTitle
            'UPGRADE_ISSUE: TextBox property lblJournalTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 5
            .Text = "Journal Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(975)
            .Visible = True
        End With
        With Me.lblArticleDesignation
            'UPGRADE_ISSUE: TextBox property lblArticleDesignation'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 6
            .Text = "Article Designation"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 17
            .Text = "Article Title"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(8520)
            '.TabIndex = 18
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With

        With Me.chkYear
            .Text = "Check to keep same year"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(9720)
            .TabIndex = 999
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With

        With Me.cmbJournalTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            '.Sorted = -1             'True
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(6135)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(8520)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbArticleDesignation
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbPublicationMonthOrSeason
            '.height = 315
            .Left = VB6.TwipsToPixelsX(3240)
            If Me.cmbPagination.Text = "Nonconsecutive" Then .TabIndex = 5 Else .TabIndex = 9
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2055)
            .Visible = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            If Me.cmbPagination.Text = "Consecutive" Then .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
        End With
        With Me.txtPublicationDay
            '.height = 285
            .Left = VB6.TwipsToPixelsX(6240)
            If Me.cmbPagination.Text = "Nonconsecutive" Then .TabIndex = 6 Else .TabIndex = 10
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            If Me.cmbPagination.Text = "Consecutive" Then .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
        End With
        With Me.txtVolume
            '.height = 285
            .Left = VB6.TwipsToPixelsX(8280)
            If Me.cmbPagination.Text = "Consecutive" Then .TabIndex = 5 Else .TabIndex = 9
            '.TabIndex = 7
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            If Me.cmbPagination.Text = "Nonconsecutive" Then .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)

        End With
        With Me.txtPage
            '.height = 285
            .Left = VB6.TwipsToPixelsX(10440)
            .TabIndex = 8
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            .Enabled = True
        End With

        With Me.cmdNewJournal
            .Text = "New Journal"
            '.Enabled = 0             'False
            '.height = 315
            .Left = VB6.TwipsToPixelsX(6720)
            .TabIndex = 999
            .Top = VB6.TwipsToPixelsY(2520)
            .Visible = 1 'False
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmdEditJournal
            .Text = "Edit This Journal"
            '.Enabled = 0             'False
            '.height = 315
            .Left = VB6.TwipsToPixelsX(6720)
            .TabIndex = 999
            .Top = VB6.TwipsToPixelsY(2160)
            .Visible = 1 'False
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
            .Enabled = True
        End With
        With Me.chkKeepSelected
            .Text = "Check to keep same jourrnal selected for multiple entries"
            '.height = 195
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 999
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(5895)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With
    End Sub

    Private Sub Position_Treatise_Form()
        With Me.lblEditionAndPrinting
            'UPGRADE_ISSUE: TextBox property lblEditionandPrinting'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3000)
            '.TabIndex = 17
            .Text = "Edition And Printing"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
        End With
        With Me.lblCallNumber
            'UPGRADE_ISSUE: TextBox property lblCallNumber'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10440)
            '.TabIndex = 16
            .Text = "Call Number"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblTitleOfSeriesIfNotIssuedByAuthor
            'UPGRADE_ISSUE: TextBox property lblTitleOfSeriesIfNotIssuedByAuthor'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(1800)
            '.TabIndex = 15
            .Text = "Title Of Series If Not Issued By Author"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(2775)
            .Visible = True
        End With
        With Me.lblPublisher
            'UPGRADE_ISSUE: TextBox property lblPublisher'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(7440)
            '.TabIndex = 14
            .Text = "Publisher"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.lblSeriesVolume
            'UPGRADE_ISSUE: TextBox property lblSeriesVolume'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 13
            .Text = "Series Volume"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblOriginalPublicationDate
            'UPGRADE_ISSUE: TextBox property lblOriginalPublicationDate'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(5160)
            '.TabIndex = 12
            .Text = "Original Publication Date"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1815)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 10
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 9
            .Text = "Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With


        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtEditionAndPrinting
            '.height = 285
            .Left = VB6.TwipsToPixelsX(3000)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1695)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtOriginalPublicationDate
            '.height = 285
            .Left = VB6.TwipsToPixelsX(5160)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1815)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtPublisher
            '.height = 285
            .Left = VB6.TwipsToPixelsX(7440)
            .TabIndex = 5
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(4815)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtSeriesVolume
            '.height = 285
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtTitleOfSeriesIfNotIssuedByAuthor
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1800)
            .TabIndex = 7
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(8415)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtCallNumber
            '.height = 285
            .Left = VB6.TwipsToPixelsX(10440)
            .TabIndex = 8
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1815)
            .Visible = True
            .Enabled = True
        End With

        With Me.chkYear
            .Text = "Keep year"
            '.height = 495
            .Left = VB6.TwipsToPixelsX(1680)
            '.TabIndex = 999
            .Top = VB6.TwipsToPixelsY(2860)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With

    End Sub
    Private Sub Position_Chapter_Form()
        With Me.lblPage
            'UPGRADE_ISSUE: TextBox property lblPage'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(9000)
            '.TabIndex = 2
            .Text = "Page Number"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10920)
            '.TabIndex = 18
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.chkKeepSelected
            .Text = "Check to keep same Larger Work selected for multiple entries"
            '.height = 195
            .Left = VB6.TwipsToPixelsX(2040)
            '.TabIndex = 10
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(4815)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With
        With Me.cmdNewLargerWork
            .Text = "New Larger Work"
            '.height = 315
            .Left = VB6.TwipsToPixelsX(10320)
            '.TabIndex = 9
            .Top = VB6.TwipsToPixelsY(3100)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
            .Enabled = True
        End With
        '   With Me.cmdEditLargerWork
        '   .Caption = "Edit This Larger Work"
        '   '.height = 315
        '   .Left = 10320
        '   '.TabIndex = 9
        '   .Top = 2880
        '   .Width = 1935
        '   .Visible = True
        '   .Enabled = True
        'End With
        With Me.lblLargerWorkTitle
            'UPGRADE_ISSUE: TextBox property lblLargerWorkTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 8
            .Text = "Larger Work Title"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.lblTitleOfSeriesIfNotIssuedByAuthor
            'UPGRADE_ISSUE: TextBox property lblTitleOfSeriesIfNotIssuedByAuthor'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(2040)
            '.TabIndex = 6
            .Text = "Title Of Series If Not Issued By Author"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(2775)
            .Visible = True
        End With
        With Me.lblSeriesVolume
            'UPGRADE_ISSUE: TextBox property lblSeriesVolume'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 5
            .Text = "Series Volume"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 3
            .Text = "Chapter Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With

        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbLargerWorkTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(9615)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtSeriesVolume
            '.height = 285
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtTitleOfSeriesIfNotIssuedByAuthor
            '.height = 285
            .Left = VB6.TwipsToPixelsX(2040)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(6374)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtPage
            '.height = 285
            .Left = VB6.TwipsToPixelsX(9000)
            .TabIndex = 5
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(10920)
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.chkYear
            .Text = "Keep year"
            '.height = 495
            .Left = VB6.TwipsToPixelsX(10920)
            '.TabIndex = 999
            .Top = VB6.TwipsToPixelsY(4200)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With
    End Sub
    Private Sub Position_Legislative_Form()
        With Me.lblLegislativeType
            'UPGRADE_ISSUE: TextBox property lblLegislativeType'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 21
            .Text = "Legislative Type"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblSuDocNumber
            'UPGRADE_ISSUE: TextBox property lblSuDocNumber'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(7920)
            '.TabIndex = 20
            .Text = "SuDoc Number"
            .Top = VB6.TwipsToPixelsY(3720)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblStateLegislativeSession
            'UPGRADE_ISSUE: TextBox property lblStateLegislativeSession'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10080)
            '.TabIndex = 19
            .Text = "State Legislative Session"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1815)
            .Visible = True
        End With
        With Me.lblSessionOfCongress
            'UPGRADE_ISSUE: TextBox property lblSessionOfCongress'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(2640)
            '.TabIndex = 18
            .Text = "Session of Congress"
            .Top = VB6.TwipsToPixelsY(3720)
            .Width = VB6.TwipsToPixelsX(1575)
            .Visible = True
        End With
        With Me.lblNumberOfCongress
            'UPGRADE_ISSUE: TextBox property lblNumberOfCongress'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 17
            .Text = "Number of Congress"
            .Top = VB6.TwipsToPixelsY(3720)
            .Width = VB6.TwipsToPixelsX(1575)
            .Visible = True
        End With
        With Me.lblLegislativeHouse
            'UPGRADE_ISSUE: TextBox property lblLegislativeHouse'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(6960)
            '.TabIndex = 16
            .Text = "Legislative House"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.lblReportOrDocumentNumber
            'UPGRADE_ISSUE: TextBox property lblReportOrDocumentNumber'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(5040)
            '.TabIndex = 15
            .Text = "Report or Document Number"
            .Top = VB6.TwipsToPixelsY(3720)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
        End With
        With Me.lblUSCCANCitation
            'UPGRADE_ISSUE: TextBox property lblUSCCANCitation'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10320)
            '.TabIndex = 14
            .Text = "USCCAN Citation"
            .Top = VB6.TwipsToPixelsY(3720)
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(2880)
            '.TabIndex = 5
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Text = "Legislative Work Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
            '.Enabled = True
        End With

        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbLegislativeType
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2055)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(2880)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True

        End With
        With Me.txtLegislativeHouse
            '.height = 285
            .Left = VB6.TwipsToPixelsX(6960)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtStateLegislativeSession
            '.height = 285
            .Left = VB6.TwipsToPixelsX(10080)
            .TabIndex = 5
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtNumberOfCongress
            '.height = 285
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3960)
            .Width = VB6.TwipsToPixelsX(1575)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtSessionOfCongress
            '.height = 285
            .Left = VB6.TwipsToPixelsX(2640)
            .TabIndex = 7
            .Top = VB6.TwipsToPixelsY(3960)
            .Width = VB6.TwipsToPixelsX(1695)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtReportOrDocumentNumber
            '.height = 285
            .Left = VB6.TwipsToPixelsX(5040)
            .TabIndex = 8
            .Top = VB6.TwipsToPixelsY(3960)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtSuDocNumber
            '.height = 285
            .Left = VB6.TwipsToPixelsX(7920)
            .TabIndex = 9
            .Top = VB6.TwipsToPixelsY(3960)
            .Width = VB6.TwipsToPixelsX(1695)
            .Visible = True
            .Enabled = True
        End With

        With Me.txtUSCCANCitation
            '.height = 285
            .Left = VB6.TwipsToPixelsX(10320)
            .TabIndex = 10
            .Top = VB6.TwipsToPixelsY(3960)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
            .Enabled = True
        End With


        With Me.chkYear
            .Text = "Check to keep same year"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(4080)
            '.TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With

    End Sub
    Private Sub Position_Misc_Form()

        With Me.lblMiscType
            'UPGRADE_ISSUE: TextBox property lblMiscType'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 12
            .Text = "Miscellaneous Type"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1575)
            .Visible = True
        End With
        With Me.lblLocation
            'UPGRADE_ISSUE: TextBox property lblLocation'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(5400)
            '.TabIndex = 11
            .Text = "Location"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblPublicationMonthOrSeason
            'UPGRADE_ISSUE: TextBox property lblPublicationMonthOrSeason'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 0
            .Text = "Publication Month or Season"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(2055)
            .Visible = True
        End With
        With Me.lblPublicationDay
            'UPGRADE_ISSUE: TextBox property lblPublicationDay'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3480)
            '.TabIndex = 1
            .Text = "Publication Day"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.chkYear
            .Text = "Check to keep same year"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(4680)
            '.TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 7
            .Text = "Miscellaneous Work Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3480)
            '.TabIndex = 8
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With

        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbMiscType
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2295)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(3480)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbPublicationMonthOrSeason
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2295)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtPublicationDay
            '.height = 315
            .Left = VB6.TwipsToPixelsX(3480)
            .TabIndex = 5
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtLocation
            '.height = 285
            .Left = VB6.TwipsToPixelsX(5400)
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2415)
            .Visible = True
            .Enabled = True
        End With


    End Sub
    Private Sub Position_Unpublished_Form()
        With Me.lblUnpublishedType
            'UPGRADE_ISSUE: TextBox property lblUnpublishedType'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 15
            .Text = "Unpublished Type"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1575)
            .Visible = True
        End With
        With Me.lblThesisDissertationType
            'UPGRADE_ISSUE: TextBox property lblThesisDissertationType'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3480)
            '.TabIndex = 14
            .Text = "Thesis/Dissertation Type"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1815)
            .Visible = True
        End With
        With Me.lblLocation
            'UPGRADE_ISSUE: TextBox property lblLocation'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(6240)
            '.TabIndex = 13
            .Text = "Location"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblPublicationMonthOrSeason
            'UPGRADE_ISSUE: TextBox property lblPublicationMonthOrSeason'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 0
            .Text = "Month or Season"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(2055)
            .Visible = True
        End With
        With Me.lblPublicationDay
            'UPGRADE_ISSUE: TextBox property lblPublicationDay'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(3480)
            '.TabIndex = 1
            .Text = "Publication Day"
            .Top = VB6.TwipsToPixelsY(3600)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.chkYear
            .Text = "Check to keep same year"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(7680)
            '.TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True
        End With
        With Me.lblTitle
            'UPGRADE_ISSUE: TextBox property lblTitle'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            '.TabIndex = 7
            .Text = "Unpublished Work Title"
            .Top = VB6.TwipsToPixelsY(2160)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
        End With
        With Me.lblYear
            'UPGRADE_ISSUE: TextBox property lblYear'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            .BorderStyle = System.Windows.Forms.BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(6240)
            '.TabIndex = 8
            .Text = "Publication Year"
            .Top = VB6.TwipsToPixelsY(2880)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With

        With Me.txtTitle
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(2400)
            .Width = VB6.TwipsToPixelsX(11775)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbUnpublishedType
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(2295)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbThesisDissertationType
            ''.height = 315
            .Left = VB6.TwipsToPixelsX(3480)
            .TabIndex = 3
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1935)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtYear
            '.height = 315
            .Left = VB6.TwipsToPixelsX(6240)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(3120)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmbPublicationMonthOrSeason
            '.height = 315
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 5
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2295)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtPublicationDay
            '.height = 285
            .Left = VB6.TwipsToPixelsX(3480)
            .TabIndex = 6
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtLocation
            '.height = 285
            .Left = VB6.TwipsToPixelsX(6240)
            .TabIndex = 7
            .Top = VB6.TwipsToPixelsY(3840)
            .Width = VB6.TwipsToPixelsX(2415)
            .Visible = True
            .Enabled = True
        End With

    End Sub


    Private Sub Position_Initial_Form()
        Dim Back As Object
        With Me.chkSource
            .Text = "Check to keep same type"
            If Me.tglNewRecords.get_Value() = True Then .Enabled = True Else .Enabled = False
            '.height = 255
            .Left = VB6.TwipsToPixelsX(3240)
            .TabIndex = 119
            .Top = VB6.TwipsToPixelsY(960)
            .Width = VB6.TwipsToPixelsX(2175)
            .Visible = True
            .Enabled = True
        End With


        With Me.lblKeywords
            'UPGRADE_ISSUE: TextBox property lblKeywords'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblKeywords.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 116
            .Text = "Select Keywords"
            .Top = VB6.TwipsToPixelsY(7080)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With

        With Me.lblSourceType
            'UPGRADE_ISSUE: TextBox property lblSourceType'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblSourceType.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(2160)
            .TabIndex = 113
            .Text = "Source Type"
            .Top = VB6.TwipsToPixelsY(960)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.lblInputInitials
            'UPGRADE_ISSUE: TextBox property lblInputInitials'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblInputInitials.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(11280)
            .TabIndex = 105
            .Text = "Input Initials"
            .Top = VB6.TwipsToPixelsY(960)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.lblDateUpdated
            'UPGRADE_ISSUE: TextBox property lblDateUpdated'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblDateUpdated.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(9600)
            .TabIndex = 104
            .Text = "Date Updated"
            .Top = VB6.TwipsToPixelsY(960)
            .Width = VB6.TwipsToPixelsX(1215)
            .Visible = True
        End With
        With Me.txtStatus
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(6720)
            .TabIndex = 101
            .Text = "Status:Not Saved"
            .Top = VB6.TwipsToPixelsY(11520)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1695)
            .Enabled = True
        End With
        With Me.lstNewKeywords
            '.height = 840
            .Left = VB6.TwipsToPixelsX(10080)
            .Sorted = -1 'True
            .TabIndex = 1
            .Top = VB6.TwipsToPixelsY(7320)
            .Width = VB6.TwipsToPixelsX(2535)
            .Visible = True
            .Enabled = True
        End With
        With Me.cmdGetNewKeywords
            .Text = "Suggest New Keywords"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(10080)
            .TabIndex = 2
            .Top = VB6.TwipsToPixelsY(7080)
            .Width = VB6.TwipsToPixelsX(2535)
            .Visible = True
            .Enabled = True
        End With
        'UPGRADE_WARNING: Couldn't resolve default property of object Me.cmdNewKeyword. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        
        With Me.cmdNewAuthor
            .Text = "New Author"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(3240)
            .TabIndex = 4
            .Top = VB6.TwipsToPixelsY(5160)
            .Width = VB6.TwipsToPixelsX(1455)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtMiscID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 86
            .Top = VB6.TwipsToPixelsY(10800)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)
        End With
        With Me.txtUnpublishedID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 85
            .Top = VB6.TwipsToPixelsY(10320)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)

        End With
        With Me.txtLegislativeID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(3600)
            .TabIndex = 84
            .Top = VB6.TwipsToPixelsY(10200)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)

        End With
        With Me.txtTreatiseID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 83
            .Top = VB6.TwipsToPixelsY(10560)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)
        End With
        With Me.txtChapterID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 82
            .Top = VB6.TwipsToPixelsY(10080)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)
        End With
        With Me.txtArticleID
            .Enabled = 0 'False
            '.height = 285
            .Left = VB6.TwipsToPixelsX(3600)
            .TabIndex = 81
            .Top = VB6.TwipsToPixelsY(10440)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1095)
        End With
        With Me.cmbRecordNumber
            '.height = 315
            '.itemdata        =   "frmMain.frx":0000
            .Left = VB6.TwipsToPixelsX(480)
            '.list            =   "frmMain.frx":0002
            .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'Dropdown .list
            .TabIndex = 80
            .Top = VB6.TwipsToPixelsY(1200)
            .Width = VB6.TwipsToPixelsX(1215)

            .Visible = True
            .Enabled = True
        End With
        With Me.cmdNextRecord
            .Text = "-->"
            '.height = 495
            .Left = VB6.TwipsToPixelsX(9480)
            .TabIndex = 75
            .Top = VB6.TwipsToPixelsY(10680)
            .Width = VB6.TwipsToPixelsX(1215)

            .Visible = True
            .Enabled = True
        End With
        With Me.cmdPreviousRecord
            .Text = "<--"
            '.height = 495
            .Left = VB6.TwipsToPixelsX(4560)
            .TabIndex = 74
            .Top = VB6.TwipsToPixelsY(10680)
            .Width = VB6.TwipsToPixelsX(1215)

            .Visible = True
            .Enabled = True
        End With
        With Me.cmdSave
            .Text = "Save"
            '.height = 495
            .Left = VB6.TwipsToPixelsX(6960)
            .TabIndex = 73
            .Top = VB6.TwipsToPixelsY(10680)
            .Width = VB6.TwipsToPixelsX(1215)

            .Visible = True
            .Enabled = True
        End With
        With Me.lstKeywords
            '.height = 840
            .Left = VB6.TwipsToPixelsX(480)
            .Sorted = -1 'True
            .TabIndex = 13
            .Top = VB6.TwipsToPixelsY(7320)
            .Width = VB6.TwipsToPixelsX(4215)

            .Visible = True
            .Enabled = True
        End With
        With Me.lstCurrentKeywords
            '.height = 840
            .Left = VB6.TwipsToPixelsX(5640)
            .TabIndex = 12
            .Top = VB6.TwipsToPixelsY(7320)
            .Width = VB6.TwipsToPixelsX(4215)

            .Visible = True
            .Enabled = True
        End With
        With Me.cmbAETChoice
            '.height = 315
            .Left = VB6.TwipsToPixelsX(1440)
            .TabIndex = 65
            .Top = VB6.TwipsToPixelsY(5040)
            .Width = VB6.TwipsToPixelsX(1695)
            .Visible = True
            .Enabled = True
        End With

        With Me.txtLargerWorkID
            '.height = 285
            .Left = VB6.TwipsToPixelsX(1920)
            .TabIndex = 56
            .Top = VB6.TwipsToPixelsY(12000)
            .Width = VB6.TwipsToPixelsX(1455)
            .Enabled = False
            .Visible = False
        End With
        With Me.txtNotes
            '.height = 735
            .Left = VB6.TwipsToPixelsX(480)
            .Multiline = -1 'True
            .TabIndex = 18
            .Top = VB6.TwipsToPixelsY(8880)
            .Width = VB6.TwipsToPixelsX(12255)
            .Visible = True
            .Enabled = True
        End With
        With Me.txtJournalID
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000013)
            .Enabled = 0 'False
            '.height = 315
            .Left = VB6.TwipsToPixelsX(3360)
            .TabIndex = 11
            .Top = VB6.TwipsToPixelsY(11760)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(1215)
        End With
        With Me.txtInputInitials
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000013)
            '.height = 315
            .Left = VB6.TwipsToPixelsX(11280)
            .TabIndex = 68
            If Me.tglNewRecords.Enabled = True Then .Enabled = True Else .Enabled = False
            .Top = VB6.TwipsToPixelsY(1200)
            .Width = VB6.TwipsToPixelsX(1335)

            .Visible = True
            .Enabled = True
        End With
        With Me.txtDateUpdated
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000013)
            .Enabled = 0 'False
            .ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000012)
            '.height = 315
            .Left = VB6.TwipsToPixelsX(9600)
            .TabIndex = 69
            .Top = VB6.TwipsToPixelsY(1200)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.txtDateAdded
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H80000013)
            .Enabled = 0 'False
            '.height = 315
            .Left = VB6.TwipsToPixelsX(8040)
            .TabIndex = 70
            .Top = VB6.TwipsToPixelsY(1200)
            .Width = VB6.TwipsToPixelsX(1335)
            .Visible = True
        End With
        With Me.cmbSourceType
            '.height = 315
            '.itemdata        =   "frmMain.frx":0004
            .Left = VB6.TwipsToPixelsX(2160)
            '.list            =   "frmMain.frx":0006
            .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'Dropdown .list
            .TabIndex = 0
            .Top = VB6.TwipsToPixelsY(1200)
            .Width = VB6.TwipsToPixelsX(3855)
            .Visible = True
            .Enabled = True
        End With
        With Me.lstAuthors
            '.height = 840
            .Left = VB6.TwipsToPixelsX(480)
            .Sorted = -1 'True
            .TabIndex = 71
            .Top = VB6.TwipsToPixelsY(5400)
            .Width = VB6.TwipsToPixelsX(4215)
            .Visible = True
            .Enabled = True
        End With
        With Me.lstTranslators
            .Enabled = 0 'False
            '.height = 840
            .Left = VB6.TwipsToPixelsX(480)
            .Sorted = -1 'True
            .TabIndex = 25
            .Top = VB6.TwipsToPixelsY(5400)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(4215)
            .Visible = True
        End With
        With Me.lstCurrentAuthors
            '.height = 840
            .Left = VB6.TwipsToPixelsX(5640)
            .TabIndex = 77
            .Top = VB6.TwipsToPixelsY(5400)
            .Width = VB6.TwipsToPixelsX(4215)
            .Visible = True
            .Enabled = True
        End With
        With Me.lstCurrentTranslators
            .Enabled = 0 'False
            '.height = 840
            .Left = VB6.TwipsToPixelsX(5640)
            .Sorted = -1 'True
            .TabIndex = 23
            .Top = VB6.TwipsToPixelsY(5400)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(4215)
        End With
        With Me.lstCurrentEditors
            .Enabled = 0 'False
            '.height = 840
            .Left = VB6.TwipsToPixelsX(5640)
            .Sorted = -1 'True
            .TabIndex = 78
            .Top = VB6.TwipsToPixelsY(5400)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(4215)
        End With
        With Me.lstEditors
            .Enabled = 0 'False
            '.height = 840
            .Left = VB6.TwipsToPixelsX(480)
            .Sorted = -1 'True
            .TabIndex = 21
            .Top = VB6.TwipsToPixelsY(5400)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(4215)
        End With
        With Me.lblT
            'UPGRADE_ISSUE: TextBox property lblT'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblT.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            .ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000003)
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10680)
            .TabIndex = 6
            .Text = "No Translator"
            .Top = VB6.TwipsToPixelsY(6120)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
        End With
        With Me.lblE
            'UPGRADE_ISSUE: TextBox property lblE'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblE.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = 0 'None
            .BorderStyle = BorderStyle.None
            .Enabled = 0 'False
            .ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000003)
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10680)
            .TabIndex = 8
            .Text = "No Editor"
            .Top = VB6.TwipsToPixelsY(5640)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
        End With
        With Me.lblA
            'UPGRADE_ISSUE: TextBox property lblA'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblA.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            .ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000003)
            '.height = 195
            .Left = VB6.TwipsToPixelsX(10680)
            .TabIndex = 62
            .Text = "No Author"
            .Top = VB6.TwipsToPixelsY(5160)
            .Width = VB6.TwipsToPixelsX(1095)
            .Visible = True
        End With
        With Me.lblRecordNumber
            'UPGRADE_ISSUE: CommandButton property lblRecordNumber'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblRecordNumber.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

            '.BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(480)
            .TabIndex = 63
            .Text = "Record Number"
            .Top = VB6.TwipsToPixelsY(960)
            .Width = VB6.TwipsToPixelsX(1335)
        End With
        With Me.lblAETChoice
            'UPGRADE_ISSUE: TextBox property lblAETChoice'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblAETChoice.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(720)
            .TabIndex = 55
            .Text = "Select"
            .Top = VB6.TwipsToPixelsY(5040)
            .Width = VB6.TwipsToPixelsX(735)

            .Visible = True
        End With
        With Me.lblDoubleClickToAdd
            'UPGRADE_ISSUE: TextBox property lblDoubleClickToAdd'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblDoubleClickToAdd.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 675
            .Left = VB6.TwipsToPixelsX(4680)
            .Multiline = -1 'True
            .TabIndex = 72
            '.Text            =   "frmMain.frx":0008
            .Top = VB6.TwipsToPixelsY(7560)
            .Width = VB6.TwipsToPixelsX(975)

            .Visible = True
        End With
        With Me.lblArrow
            'UPGRADE_ISSUE: TextBox property lblArrow'.Appearance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            '.Appearance = 0 'Flat
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblArrow.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.None 'None
            .Enabled = 0 'False
            '.height = 195
            .Left = VB6.TwipsToPixelsX(4680)
            .TabIndex = 57
            .Text = "<<<--------->>>"
            .Top = VB6.TwipsToPixelsY(5400)
            .Width = VB6.TwipsToPixelsX(975)
            .Visible = True
        End With
        With Me.frmEntryInfo
            .Text = "Entry Information"
            '.height = 975
            .Left = VB6.TwipsToPixelsX(7800)
            .TabIndex = 42
            .Top = VB6.TwipsToPixelsY(720)
            .Width = VB6.TwipsToPixelsX(5175)
            .Visible = True
        End With
        With Me.frmRecordInfo
            .Text = "Record Information"
            '.height = 975
            .Left = VB6.TwipsToPixelsX(240)
            .TabIndex = 79
            .Top = VB6.TwipsToPixelsY(720)
            .Width = VB6.TwipsToPixelsX(6615)
            .Visible = True
        End With
        With Me.frmCitationInfo
            .Text = "Citation Information"
            '.height = 2775
            .Left = VB6.TwipsToPixelsX(240)
            .TabIndex = 96
            .Top = VB6.TwipsToPixelsY(1800)
            .Width = VB6.TwipsToPixelsX(12735)
            .Visible = True
        End With
        With Me.frmAuthorInfo
            .Text = "Author Information"
            '.height = 1935
            .Left = VB6.TwipsToPixelsX(240)
            .TabIndex = 97
            .Top = VB6.TwipsToPixelsY(4680)
            .Width = VB6.TwipsToPixelsX(12735)
            .Visible = True
        End With
        With Me.frmKeywordInfo
            .Text = "Keyword Information"
            '.height = 1815
            .Left = VB6.TwipsToPixelsX(240)
            .TabIndex = 66
            .Top = VB6.TwipsToPixelsY(6720)
            .Width = VB6.TwipsToPixelsX(12735)
            .Visible = True
        End With
        With Me.frmNotes
            .Text = "Notes"
            '.height = 1095
            .Left = VB6.TwipsToPixelsX(240)
            .TabIndex = 67
            .Top = VB6.TwipsToPixelsY(8640)
            .Width = VB6.TwipsToPixelsX(12735)
            .Visible = True
        End With
        With Me.lblSeparateBottom
            'UPGRADE_WARNING: Couldn't resolve default property of object Back.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Back.Style = 0 'Transparent
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblSeparateBottom.Border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .BorderStyle = BorderStyle.FixedSingle 'Fixed Single
            '.height = 5175
            .Left = VB6.TwipsToPixelsX(-3600)
            .TabIndex = 98
            .Top = VB6.TwipsToPixelsY(12120)
            .Width = VB6.TwipsToPixelsX(30000)
            .Visible = True
        End With
        'UPGRADE_WARNING: Couldn't resolve default property of object Me.lblSeparate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        
        With Me.lblMiscID
            .Text = "Misc ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(600)
            .TabIndex = 95
            .Top = VB6.TwipsToPixelsY(10800)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.lblTreatiseID
            .Text = "Treatise ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(600)
            .TabIndex = 94
            .Top = VB6.TwipsToPixelsY(10560)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.lblUnpublishedID
            .Text = "Unpublished ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(600)
            .TabIndex = 93
            .Top = VB6.TwipsToPixelsY(10320)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.lblChapterID
            .Text = "Chapter ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(600)
            .TabIndex = 92
            .Top = VB6.TwipsToPixelsY(10080)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.lblArticleID
            .Text = "Article ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(2760)
            .TabIndex = 91
            .Top = VB6.TwipsToPixelsY(10440)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.lblLegisID
            .Text = "Legis ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(2760)
            .TabIndex = 90
            .Top = VB6.TwipsToPixelsY(10200)
            .Visible = 0 'False
            .Width = VB6.TwipsToPixelsX(855)
        End With
        With Me.tglNewRecords
            '.height = 375
            .Left = VB6.TwipsToPixelsX(2040)
            .TabIndex = 100
            .Top = VB6.TwipsToPixelsY(120)
            .Width = VB6.TwipsToPixelsX(1455)
            '.BackColor = -2147483633
            '.ForeColor = -2147483630
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglNewRecords.Display. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.DisplayStyle = Microsoft.Vbe.Interop.Forms.fmDisplayStyle.fmDisplayStyleToggle
            '.Size = 2566;661
            .set_Value("0")
            .Text = "New Entries"
            '.Font(VB6.PixelsToTwipsY(Height) = 165)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglNewRecords.FontCharSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontCharSet = 204
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglNewRecords.FontPitchAndFamily. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontPitchAndFamily = 2
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglNewRecords.ParagraphAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.ParagraphAlign = 3
            .Visible = True
        End With
        With Me.tglUpdateRecords
            '.height = 375
            .Left = VB6.TwipsToPixelsX(6240)
            .TabIndex = 36
            .Top = VB6.TwipsToPixelsY(120)
            .Width = VB6.TwipsToPixelsX(1455)
            '.BackColor = -2147483633
            '.ForeColor = -2147483630
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglUpdateRecords.Display. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.Display.Style = 6
            '.Size = "2566;661"
            .set_Value("0")
            .Text = "Update Records"
            '.Font(VB6.PixelsToTwipsY(Height) = 165)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglUpdateRecords.FontCharSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontCharSet = 204
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglUpdateRecords.FontPitchAndFamily. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontPitchAndFamily = 2
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglUpdateRecords.ParagraphAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.ParagraphAlign = 3
            .Visible = True
        End With
        With Me.tglImportRecords
            '.height = 375
            .Left = VB6.TwipsToPixelsX(10320)
            .TabIndex = 35
            .Top = VB6.TwipsToPixelsY(120)
            .Width = VB6.TwipsToPixelsX(1455)
            '.BackColor = -2147483633
            '.ForeColor = -2147483630
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglImportRecords.Display. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.Display.Style = 6
            '.Size = "2566;661"
            .set_Value("0")
            .Text = "Import Records"
            '.Font(VB6.PixelsToTwipsY(Height) = 165)
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglImportRecords.FontCharSet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontCharSet = 204
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglImportRecords.FontPitchAndFamily. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.FontPitchAndFamily = 2
            'UPGRADE_WARNING: Couldn't resolve default property of object Me.tglImportRecords.ParagraphAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '.ParagraphAlign = 3
            .Visible = True
        End With
        With Me.lblLargerWorkID
            .Text = "Larger Work ID"
            '.height = 255
            .Left = VB6.TwipsToPixelsX(7320)
            .TabIndex = 87
            .Top = VB6.TwipsToPixelsY(10080)
            .Width = VB6.TwipsToPixelsX(1335)
        End With
    End Sub
	
	
	
	
	
	Private Sub Populate_Comboboxes()
		Dim icounter As Short
		
		Call Populate_Journal_Combobox()
		
		Call Populate_LargerWork_Combobox()
		
		Call Populate_AET_Lists()
		
		Call Populate_Keyword_List()
		
		Call populate_RecordID_List()
		
		cmbAETChoice.Items.Add("Authors")
		cmbAETChoice.Items.Add("Editors")
		cmbAETChoice.Items.Add("Translators")
		
		
		cmbSourceType.Items.Add("Journal Article")
		cmbSourceType.Items.Add("Treatise")
		cmbSourceType.Items.Add("Chapter in Treatise")
		cmbSourceType.Items.Add("Unpublished Work")
		cmbSourceType.Items.Add("Legislative Material")
		cmbSourceType.Items.Add("Nonprint Material")
		
		
		cmbArticleDesignation.Items.Add("Abstract")
		cmbArticleDesignation.Items.Add("Annotation")
		cmbArticleDesignation.Items.Add("Book Note")
		cmbArticleDesignation.Items.Add("Book Review")
		cmbArticleDesignation.Items.Add("Case Comment")
		cmbArticleDesignation.Items.Add("Case Note")
		cmbArticleDesignation.Items.Add("Comment")
		cmbArticleDesignation.Items.Add("Note")
		cmbArticleDesignation.Items.Add("Recent Case")
		cmbArticleDesignation.Items.Add("Recent Decision")
		cmbArticleDesignation.Items.Add("Recent Development")
		cmbArticleDesignation.Items.Add("Recent Statute")
		cmbArticleDesignation.Items.Add("Symposium")
		
		cmbPublicationMonthOrSeason.Items.Add("Jan.")
		cmbPublicationMonthOrSeason.Items.Add("Jan./Feb.")
		cmbPublicationMonthOrSeason.Items.Add("Feb.")
		cmbPublicationMonthOrSeason.Items.Add("Feb./Mar.")
		cmbPublicationMonthOrSeason.Items.Add("Mar.")
		cmbPublicationMonthOrSeason.Items.Add("Mar./Apr.")
		cmbPublicationMonthOrSeason.Items.Add("Apr.")
		cmbPublicationMonthOrSeason.Items.Add("Apr./May")
		cmbPublicationMonthOrSeason.Items.Add("May")
		cmbPublicationMonthOrSeason.Items.Add("May/June")
		cmbPublicationMonthOrSeason.Items.Add("June")
		cmbPublicationMonthOrSeason.Items.Add("June/July")
		cmbPublicationMonthOrSeason.Items.Add("July")
		cmbPublicationMonthOrSeason.Items.Add("July/Aug.")
		cmbPublicationMonthOrSeason.Items.Add("Aug.")
		cmbPublicationMonthOrSeason.Items.Add("Aug./Sept.")
		cmbPublicationMonthOrSeason.Items.Add("Sept.")
		cmbPublicationMonthOrSeason.Items.Add("Sept./Oct.")
		cmbPublicationMonthOrSeason.Items.Add("Oct.")
		cmbPublicationMonthOrSeason.Items.Add("Oct./Nov.")
		cmbPublicationMonthOrSeason.Items.Add("Nov.")
		cmbPublicationMonthOrSeason.Items.Add("Nov./Dec.")
		cmbPublicationMonthOrSeason.Items.Add("Dec.")
		cmbPublicationMonthOrSeason.Items.Add("Dec./Jan.")
		cmbPublicationMonthOrSeason.Items.Add("Spring")
		cmbPublicationMonthOrSeason.Items.Add("Summer")
		cmbPublicationMonthOrSeason.Items.Add("Fall")
		cmbPublicationMonthOrSeason.Items.Add("Winter")
		
		cmbPagination.Items.Add("Consecutive")
		cmbPagination.Items.Add("Nonconsecutive")
		
		cmbMiscType.Items.Add("Electronic Paginated")
		cmbMiscType.Items.Add("Internet Site")
		cmbMiscType.Items.Add("Film")
		cmbMiscType.Items.Add("Audio Recording")
		cmbMiscType.Items.Add("Electronic Material")
		
		cmbLegislativeType.Items.Add("Committee Hearing")
		cmbLegislativeType.Items.Add("Report")
		cmbLegislativeType.Items.Add("Conference Report")
		cmbLegislativeType.Items.Add("Committee Print")
		cmbLegislativeType.Items.Add("Executive Document")
		cmbLegislativeType.Items.Add("Miscellaneous Document")
		cmbLegislativeType.Items.Add("State Material")
		
		cmbUnpublishedType.Items.Add("Manuscript")
		cmbUnpublishedType.Items.Add("Dissertation")
		cmbUnpublishedType.Items.Add("Thesis")
		
		cmbThesisDissertationType.Items.Add("Ph.D")
		cmbThesisDissertationType.Items.Add("M.A.")
		cmbThesisDissertationType.Items.Add("M.S.")
		cmbThesisDissertationType.Items.Add("A.B.")
		cmbThesisDissertationType.Items.Add("B.A.")
		cmbThesisDissertationType.Items.Add("B.S.")
		cmbThesisDissertationType.Items.Add("M.B.A.")
		cmbThesisDissertationType.Items.Add("B.B.A.")
	End Sub
	
	Public Sub Populate_Journal_Combobox()
		Dim rstTempJournals As ADODB.Recordset
		Dim sTempJournalSource As String
		cmbJournalTitle.Items.Clear()
		rstTempJournals = New ADODB.Recordset
		sTempJournalSource = "SELECT * FROM tblJournals"
		rstTempJournals.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstTempJournals.Open(sTempJournalSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		Do While Not rstTempJournals.EOF
			If rstTempJournals.Fields("JournalTitle").Value <> "" Then
				cmbJournalTitle.Items.Add(rstTempJournals.Fields("JournalTitle").Value)
			End If
			rstTempJournals.MoveNext()
		Loop 
		rstTempJournals.Close()
		'UPGRADE_NOTE: Object rstTempJournals may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempJournals = Nothing
	End Sub
	Public Sub Populate_LargerWork_Combobox()
		Dim rstTempLargerWorks As ADODB.Recordset
		Dim sTempLargerWorkSource As String
		cmbLargerWorkTitle.Items.Clear()
		rstTempLargerWorks = New ADODB.Recordset
		sTempLargerWorkSource = "SELECT * FROM tblLargerWorks"
		rstTempLargerWorks.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rstTempLargerWorks.Open(sTempLargerWorkSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        Do While Not rstTempLargerWorks.EOF
            'If Not IsDBNull(rstTempLargerWorks("LargerWorkTitle")) AndAlso rstTempLargerWorks.Fields("LargerWorkTitle"). <> "" Then
            If Not IsDBNull(rstTempLargerWorks("LargerWorkTitle")) Then

                cmbLargerWorkTitle.Items.Add(rstTempLargerWorks.Fields("LargerWorkTitle").Value)
            End If
            rstTempLargerWorks.MoveNext()

        Loop
		rstTempLargerWorks.Close()
		'UPGRADE_NOTE: Object rstTempLargerWorks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempLargerWorks = Nothing
		
		
	End Sub
	
	Public Sub Populate_AET_Lists()
		Dim rstTempAET As ADODB.Recordset
		Dim sTempAETSource As String
		Dim sTempAET As String
		
		lstAuthors.Items.Clear()
		lstEditors.Items.Clear()
		lstTranslators.Items.Clear()
		'With rstAuthors
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from qryAETRecords WHERE AETType='Author'")
		'End With
		rstTempAET = New ADODB.Recordset
		rstTempAET.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		
		sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Author'"
		rstTempAET.Open(sTempAETSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		Do While Not rstTempAET.EOF
			'rstAETLMFRecords.Open "SELECT * FROM qryAETRecords WHERE RecordID=" & iRecNum, cnDatabase, adOpenStatic, adLockPessimistic
			
			sTempAET = Full_AET_Name(rstTempAET)
			'If rstAuthors.Fields("FullName").Value <> "" Then
			'lstAuthors.AddItem rstAuthors.Fields("FullName").Value & " (ID: " & rstAuthors!AETID & ")"
			
			'End If
			lstAuthors.Items.Add(sTempAET & " (ID: " & rstTempAET.Fields("AETID").Value & ")")
			
			rstTempAET.MoveNext()
		Loop 
		
		rstTempAET.Close()
		'UPGRADE_NOTE: Object rstTempAET may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempAET = Nothing
		rstTempAET = New ADODB.Recordset
		rstTempAET.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Editor'"
		rstTempAET.Open(sTempAETSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		Do While Not rstTempAET.EOF
			sTempAET = Full_AET_Name(rstTempAET)
			'If rstEditors.Fields("FullName").Value <> "" Then
			'lstEditors.AddItem rstEditors.Fields("FullName").Value & " (ID: " & rstEditors!AETID & ")"
			'End If
			lstEditors.Items.Add(sTempAET & " (ID: " & rstTempAET.Fields("AETID").Value & ")")
			
			rstTempAET.MoveNext()
		Loop 
		
		rstTempAET.Close()
		'UPGRADE_NOTE: Object rstTempAET may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempAET = Nothing
		rstTempAET = New ADODB.Recordset
		rstTempAET.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		sTempAETSource = "SELECT * from tblAuthorsEditorsTranslators WHERE AETType='Translator'"
		rstTempAET.Open(sTempAETSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		
		Do While Not rstTempAET.EOF
			sTempAET = Full_AET_Name(rstTempAET)
			
			'If rstTranslators.Fields("FullName").Value <> "" Then
			'lstTranslators.AddItem rstTranslators.Fields("FullName").Value & " (ID: " & rstTranslators!AETID & ")"
			'End If
			lstTranslators.Items.Add(sTempAET & " (ID: " & rstTempAET.Fields("AETID").Value & ")")
			
			rstTempAET.MoveNext()
		Loop 
		rstTempAET.Close()
		'UPGRADE_NOTE: Object rstTempAET may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempAET = Nothing
		
	End Sub
	
	Public Function Full_AET_Name(ByRef rstRecordset As ADODB.Recordset) As String
		Dim sCurrentAET As String
        sCurrentAET = ""
        If Not IsDBNull(rstRecordset.Fields("InstitutionalEntity")) Then sCurrentAET = sCurrentAET & rstRecordset.Fields("InstitutionalEntity").Value
        'If rstRecordset.Fields("InstitutionalEntity").Value <> "" Then sCurrentAET = sCurrentAET & rstRecordset.Fields("InstitutionalEntity").Value
        'If rstRecordset.Fields("LastName").Value <> "" Then
        If Not IsDBNull(rstRecordset.Fields("LastName")) Then

            If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
            sCurrentAET = sCurrentAET & rstRecordset.Fields("LastName").Value
        End If
        'If rstRecordset.Fields("FirstName").Value <> "" Then
        If Not IsDBNull(rstRecordset.Fields("FirstName")) Then

            If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
            sCurrentAET = sCurrentAET & rstRecordset.Fields("FirstName").Value
        End If
        If Not IsDBNull(rstRecordset.Fields("MiddleName")) Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("MiddleName").Value

        If Not IsDBNull(rstRecordset.Fields("Suffix")) Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("Suffix").Value
        'If rstRecordset.Fields("MiddleName").Value <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("MiddleName").Value
        'If rstRecordset.Fields("Suffix").Value <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("Suffix").Value
        Full_AET_Name = sCurrentAET
    End Function
	Private Function Full_AET(ByRef rstRecordset As ADODB.Recordset, ByRef sType As String) As String
		Dim sCurrentAET As String
		sCurrentAET = ""
		Select Case sType
			Case "FMLS", "FML", "FL"
				If rstRecordset.Fields("FirstName").Value <> "" Then
					'If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
					sCurrentAET = sCurrentAET & rstRecordset.Fields("FirstName").Value
				End If
				If (sType = "FMLS") Or (sType = "FML") Then
					If rstRecordset.Fields("MiddleName").Value <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("MiddleName").Value
				End If
				If rstRecordset.Fields("LastName").Value <> "" Then
					If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & " "
					sCurrentAET = sCurrentAET & rstRecordset.Fields("LastName").Value
				End If
				If (rstRecordset.Fields("Suffix").Value <> "") And (sType = "FMLS") Then
					If rstRecordset.Fields("Suffix").Value = "Jr." Then sCurrentAET = sCurrentAET & ","
					sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("Suffix").Value
				End If
			Case "LFM"
				If rstRecordset.Fields("LastName").Value <> "" Then
					sCurrentAET = sCurrentAET & rstRecordset.Fields("LastName").Value
				End If
				If rstRecordset.Fields("FirstName").Value <> "" Then
					If sCurrentAET <> "" Then sCurrentAET = sCurrentAET & ", "
					sCurrentAET = sCurrentAET & rstRecordset.Fields("FirstName").Value
				End If
				If rstRecordset.Fields("MiddleName").Value <> "" Then sCurrentAET = sCurrentAET & " " & rstRecordset.Fields("MiddleName").Value
		End Select
		Full_AET = sCurrentAET
	End Function
	
	Public Sub Populate_Keyword_List()
		Dim rstTempKeywords As ADODB.Recordset
		Dim sTempSource As String
		
		lstKeywords.Items.Clear()
		sTempSource = "SELECT * FROM tblKeywords"
		rstTempKeywords = New ADODB.Recordset
		rstTempKeywords.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstTempKeywords.Open(sTempSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		Do While Not rstTempKeywords.EOF
			If rstTempKeywords.Fields("KeywordOrCodeSection").Value <> "" Then
				lstKeywords.Items.Add(rstTempKeywords.Fields("KeywordOrCodeSection").Value & " (ID: " & rstTempKeywords.Fields("KeywordID").Value & ")")
			End If
			rstTempKeywords.MoveNext()
		Loop 
		rstTempKeywords.Close()
		'UPGRADE_NOTE: Object rstTempKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTempKeywords = Nothing
	End Sub
	
	Public Sub populate_RecordID_List()
		Dim iRecNum As Short
		Me.cmbRecordNumber.Items.Clear()
		If Not rstRecords.EOF Then rstRecords.MoveFirst()
		Do While Not rstRecords.EOF
			iRecNum = rstRecords.Fields("RecordID").Value
			If VB6.GetItemString(cmbRecordNumber, Me.cmbRecordNumber.Items.Count - 1) = "" Then
				cmbRecordNumber.Items.Add(CStr(iRecNum))
			Else
				If Str(CDbl(VB6.GetItemString(cmbRecordNumber, cmbRecordNumber.Items.Count - 1))) <> Str(iRecNum) Then
					cmbRecordNumber.Items.Add(CStr(iRecNum))
				End If
			End If
			rstRecords.MoveNext()
		Loop 
		cmbRecordNumber.Items.Add(("New Record"))
		If Me.tglImportRecords.get_Value() = True Then MsgBox("Query executed. " & Me.cmbRecordNumber.Items.Count - 1 & " records found.")
		
	End Sub
	
	
	'UPGRADE_WARNING: Event chkLibraryCollection.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkLibraryCollection_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLibraryCollection.CheckStateChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event chkRepublished.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkRepublished_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRepublished.CheckStateChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbAETChoice.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbAETChoice_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAETChoice.SelectedIndexChanged
		Dim sChoice As String
		sChoice = Me.cmbAETChoice.Text
		Me.lblArrow.Visible = True
		Me.lblAETChoice.Visible = True
		Select Case sChoice
			Case "Authors"
				Erase_Object(lstEditors)
				Erase_Object(lstTranslators)
				Erase_Object(lstCurrentEditors)
				Erase_Object(lstCurrentTranslators)
				lstAuthors.Visible = True
				lstAuthors.Enabled = True
				lstCurrentAuthors.Visible = True
				lstCurrentAuthors.Enabled = True
				Me.cmdNewAuthor.Text = "New Author"
			Case "Editors"
				
				Erase_Object(lstAuthors)
				Erase_Object(lstTranslators)
				Erase_Object(lstCurrentAuthors)
				Erase_Object(lstCurrentTranslators)
				lstEditors.Visible = True
				lstEditors.Enabled = True
				lstCurrentEditors.Visible = True
				lstCurrentEditors.Enabled = True
				Me.cmdNewAuthor.Text = "New Editor"
				
			Case "Translators"
				
				Erase_Object(lstEditors)
				Erase_Object(lstAuthors)
				Erase_Object(lstCurrentEditors)
				Erase_Object(lstCurrentAuthors)
				lstCurrentTranslators.Visible = True
				lstCurrentTranslators.Enabled = True
				lstTranslators.Visible = True
				lstTranslators.Enabled = True
				Me.cmdNewAuthor.Text = "New Translator"
				
		End Select
	End Sub
	
	
	
	'UPGRADE_WARNING: Event cmbArticleDesignation.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbArticleDesignation_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbArticleDesignation.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'Private Sub cmbJournalTitle_click()
	
	'    Call Lookup_Journal
	'End Sub
	
	'UPGRADE_WARNING: Event cmbJournalTitle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbJournalTitle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbJournalTitle.SelectedIndexChanged
		Call Lookup_Journal()
		Call Position_Article_Form()
		Me.txtStatus.Text = "Not Saved"
	End Sub
	Private Sub Lookup_Journal()
		Dim rstJournalLookup As ADODB.Recordset
		Dim sJournalSource As String
		Dim sJournalTitle As String
		Dim sJournalTitleShortForm As String
		
		sJournalTitle = cmbJournalTitle.Text
		'sJournalTitle = Replace(sJournalTitle, "'", "*")
		If Not sJournalTitle = "" Then
			sJournalSource = "SELECT * from tblJournals"
			rstJournalLookup = New ADODB.Recordset
			rstJournalLookup.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstJournalLookup.Open(sJournalSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			rstJournalLookup.MoveFirst()
			On Error GoTo Lookup_Journal_Error
			Do Until (rstJournalLookup.EOF) Or (rstJournalLookup.Fields("JournalTitle").Value = sJournalTitle)
				rstJournalLookup.MoveNext()
			Loop 
			
			'rstjournallookup.Find "JournalTitle LIKE '" & sJournalTitle & "'"
Lookup_EOF: 
			If Not rstJournalLookup.EOF Then
				Me.txtJournalID.Text = rstJournalLookup.Fields("JournalID").Value
				Me.txtJournaTitleShortForm.Text = rstJournalLookup.Fields("JournalTitleShortFOrm").Value
				'If rstjournallookup!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstjournallookup!JournalTitleShortForm
				If rstJournalLookup.Fields("Pagination").Value <> "" Then Me.cmbPagination.Text = rstJournalLookup.Fields("Pagination").Value
				'    If rstjournallookup!CallNumber <> "" Then Me.txtCallNumber = rstjournallookup!CallNumber
				'If rstjournallookup!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstjournallookup!PLaceOfPublication
				frmNewJournal.txtJournalID.Text = rstJournalLookup.Fields("JournalID").Value
				frmNewJournal.txtNewJournal.Text = rstJournalLookup.Fields("JournalTitle").Value
				frmNewJournal.txtNewJournalShortForm.Text = rstJournalLookup.Fields("JournalTitleShortFOrm").Value
				frmNewJournal.cmbPagination.Text = rstJournalLookup.Fields("Pagination").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If rstJournalLookup.Fields("CallNumber").Value Is System.DBNull.Value Then frmNewJournal.txtCallNumber.Text = rstJournalLookup.Fields("CallNumber").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If rstJournalLookup.Fields("PlaceOfPublication").Value Is System.DBNull.Value Then frmNewJournal.txtPlaceOfPublication.Text = rstJournalLookup.Fields("PlaceOfPublication").Value
				
			End If
			rstJournalLookup.Close()
			'UPGRADE_NOTE: Object rstJournalLookup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstJournalLookup = Nothing
		End If
		
Lookup_Journal_Error: 
		Select Case Err.Number
			Case 0
			Case 3021
				Resume Lookup_EOF
				'Case Else
				'    MsgBox Err.Number & " " & Err.Description
		End Select
		
		
	End Sub
	
	'UPGRADE_WARNING: Event cmbLargerWorkTitle.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbLargerWorkTitle.Change was upgraded to cmbLargerWorkTitle.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbLargerWorkTitle_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLargerWorkTitle.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbLargerWorkTitle.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbLargerWorkTitle_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLargerWorkTitle.SelectedIndexChanged
		Dim sTempSource As String
		Dim sLargerWorkTitle As String
		sLargerWorkTitle = cmbLargerWorkTitle.Text
		'sJournalTitle = Replace(sJournalTitle, "'", "*")
		rstLargerWorks = New ADODB.Recordset
		sTempSource = "SELECT * FROM tblLargerWorks"
		rstLargerWorks.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstLargerWorks.Open(sTempSource, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		rstLargerWorks.MoveFirst()
		Do Until (rstLargerWorks.Fields("LargerWorkTitle").Value = sLargerWorkTitle) Or rstLargerWorks.EOF
			rstLargerWorks.MoveNext()
		Loop 
		
		'rstJournals.Find "JournalTitle LIKE '" & sJournalTitle & "'"
		If Not rstLargerWorks.EOF Then
			Me.txtLargerWorkID.Text = rstLargerWorks.Fields("LargerWorkID").Value
			If rstLargerWorks.Fields("CallNumber").Value <> "" Then Me.txtCallNumber.Text = rstLargerWorks.Fields("CallNumber").Value
			If rstLargerWorks.Fields("EditionAndPrinting").Value <> "" Then Me.txtEditionandPrinting.Text = rstLargerWorks.Fields("EditionAndPrinting").Value
			If rstLargerWorks.Fields("Publisher").Value <> "" Then Me.txtPublisher.Text = rstLargerWorks.Fields("Publisher").Value
			If rstLargerWorks.Fields("OriginalPublicationDate").Value <> "" Then Me.txtOriginalPublicationDate.Text = rstLargerWorks.Fields("OriginalPublicationDate").Value
			If rstLargerWorks.Fields("SeriesVolume").Value <> "" Then Me.txtSeriesVolume.Text = rstLargerWorks.Fields("SeriesVolume").Value
			If rstLargerWorks.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = rstLargerWorks.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value
			
			'TitleOfSeriesIfNotIssuedByAuthor
			
			
		End If
	End Sub
	
	
	
	'UPGRADE_WARNING: Event cmbLegislativeType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbLegislativeType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLegislativeType.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbMiscType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbMiscType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbMiscType.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbPagination.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbPagination.Change was upgraded to cmbPagination.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbPagination_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPagination.TextChanged
		Call Position_Article_Form()
	End Sub
	
	
	'UPGRADE_WARNING: Event cmbPublicationMonthOrSeason.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbPublicationMonthOrSeason_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPublicationMonthOrSeason.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbRecordNumber.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbRecordNumber_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbRecordNumber.SelectedIndexChanged
		Dim iRecNum As Short
		If Me.cmbRecordNumber.SelectedIndex <> cmbRecordNumber.Items.Count - 1 Then
			'Me.tglUpdateRecords.Value = True
			If IsNumeric(Me.cmbRecordNumber.Text) Then iRecNum = CShort(Me.cmbRecordNumber.Text)
			rstRecords.MoveFirst()
			Do Until rstRecords.Fields("RecordID").Value = iRecNum
				rstRecords.MoveNext()
			Loop 
			'rstRecords.Find "RecordID=" & iRecNum
			Call Erase_Form()
			Call Clear_Form()
			Me.Refresh()
			
			Call Change_Record_Lists()
			Call Fill_Form()
		End If
		If Me.cmbRecordNumber.Text = "New Record" Then
			Me.tglNewRecords.set_Value(True)
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbSourceType.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbSourceType.Change was upgraded to cmbSourceType.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbSourceType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSourceType.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbSourceType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbSourceType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSourceType.SelectedIndexChanged
		Dim sSourceType As String
		
		
		'Call Fill_Form
		sSourceType = cmbSourceType.Text
		Select Case sSourceType
			Case "Journal Article"
				'Call Article_Form
				Call Erase_Form()
				Call Position_Article_Form()
				
			Case "Treatise"
				'Call Treatise_Form
				Call Erase_Form()
				Call Position_Treatise_Form()
				
			Case "Chapter in Treatise"
				'Call Chapter_Form
				Call Erase_Form()
				Call Position_Chapter_Form()
				
			Case "Unpublished Work"
				'Call Unpublished_Form
				Call Erase_Form()
				Call Position_Unpublished_Form()
				
			Case "Legislative Material"
				'Call Legislative_Form
				Call Erase_Form()
				Call Position_Legislative_Form()
				
			Case "Nonprint Material"
				'Call Misc_Form
				Call Erase_Form()
				Call Position_Misc_Form()
				
		End Select
	End Sub
	
	Private Sub cmbSourceType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbSourceType.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		
		If cmbSourceType.Text = "" Then
			MsgBox("Please Enter a Source Type.")
			Cancel = True
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event cmbThesisDissertationType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbThesisDissertationType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbThesisDissertationType.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event cmbUnpublishedType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbUnpublishedType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbUnpublishedType.SelectedIndexChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
		Dim bYN As String
		Dim iRecordNumber As Short
		Dim iListindex As Short
		bYN = CStr(MsgBox("Permanently Delete Record", MsgBoxStyle.OKCancel, "Confirm Deletion"))
		Select Case bYN
			Case CStr(MsgBoxResult.OK)
				'MsgBox "Yes"
				iRecordNumber = CShort(Me.cmbRecordNumber.Text)
				iListindex = Me.cmbRecordNumber.SelectedIndex
				rstRecords.MoveFirst()
				Do Until rstRecords.Fields("RecordID").Value = iRecordNumber
					rstRecords.MoveNext()
				Loop 
				If Not rstRecords.EOF Then
					On Error GoTo CancelErr
					cnWriteDatabase.BeginTrans()
					rstRecords.Delete()
					rstRecords.Update()
					cnWriteDatabase.CommitTrans()
					rstRecords.Requery()
					Me.cmbRecordNumber.Items.RemoveAt((iListindex))
				End If
				Call cmdNextRecord_Click(cmdNextRecord, New System.EventArgs())
			Case CStr(MsgBoxResult.Cancel)
				'MsgBox "No"
		End Select
CancelErr: 
		Select Case Err.Number
			Case 0
			Case Else
				cnWriteDatabase.RollbackTrans()
		End Select
	End Sub
	
	Private Sub cmdEditJournal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEditJournal.Click
		frmNewJournal.Close()
		frmNewJournal.bEdit = True
		frmNewJournal.Show()
		frmNewJournal.Text = "Edit Journal"
		
		frmNewJournal.bEdit = False
	End Sub
	
	Private Sub cmdGetNewKeywords_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGetNewKeywords.Click
		Call suggest_keywords()
	End Sub
	Private Sub suggest_keywords()
		Dim rstOldKeywords As ADODB.Recordset
		Dim rstKeywordCheck As ADODB.Recordset
		Dim rstThesaurusCheck As ADODB.Recordset
		Dim sKeywordText As String
		Dim sThesaurusText As String
		Dim sTitleText As String
		Dim cSuggestedKeywords As Collection
		Dim i As Short
		Dim sOldKeywordString As String
		Dim iCurrentRecnum As Short
		Dim bDuplicate As Boolean
		Dim rstBigCategory As ADODB.Recordset
		Dim sTempText As String
		Dim rstExistingKeywordBigCat As ADODB.Recordset
		Dim rstJournalKeyword As ADODB.Recordset
		Dim sJournalLocation As String
		
		'the following part will be removed later
		'iCurrentRecnum = Me.cmbRecordNumber.Text
		'Set rstOldKeywords = New Recordset
		'With rstOldKeywords
		'    .ActiveConnection = cnDatabase
		'    .CursorType = adOpenForwardOnly
		'    .LockType = adLockReadOnly
		'    .Open ("SELECT AllKeywords from tblRecordsAllKeywords WHERE RecordID=" & iCurrentRecnum)
		'End With
		'If Not rstOldKeywords.EOF Then sOldKeywordString = rstOldKeywords!AllKeywords
		
		'Set rstOldKeywords = Nothing
		
		cSuggestedKeywords = New Collection
		sTitleText = Me.txtTitle.Text
		'next line taken out later
		'sTitleText = sTitleText & " " & sOldKeywordString
		Me.lstNewKeywords.Items.Clear()
		rstKeywordCheck = New ADODB.Recordset
		rstKeywordCheck.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		With rstKeywordCheck
			.let_ActiveConnection(cnReadDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
			.LockType = ADODB.LockTypeEnum.adLockReadOnly
			.Open(("SELECT * from tblKeywords"))
		End With
		
		rstThesaurusCheck = New ADODB.Recordset
		
		Do While Not rstKeywordCheck.EOF
			sKeywordText = rstKeywordCheck.Fields("keywordorcodesection").Value
			If InStr(1, sTitleText, sKeywordText) Then
				sKeywordText = sKeywordText & " (ID: " & rstKeywordCheck.Fields("KeywordID").Value & ")"
				bDuplicate = False
				For i = 0 To (Me.lstCurrentKeywords.Items.Count - 1)
					If VB6.GetItemString(Me.lstCurrentKeywords, i) = sKeywordText Then bDuplicate = True
				Next 
				If Not bDuplicate Then cSuggestedKeywords.Add(sKeywordText)
				'Me.lstNewKeywords.AddItem sKeywordText
			End If
			rstKeywordCheck.MoveNext()
		Loop 
		
		rstExistingKeywordBigCat = New ADODB.Recordset
		rstExistingKeywordBigCat.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		With rstExistingKeywordBigCat
			.let_ActiveConnection(cnReadDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
			.LockType = ADODB.LockTypeEnum.adLockReadOnly
			.Open(("SELECT * from qryRecordsKeywordsThesaurus WHERE LargerCategory=1 AND RecordID=" & iCurrentRecnum))
		End With
		
		Do While Not rstExistingKeywordBigCat.EOF
			rstBigCategory = New ADODB.Recordset
			sTempText = rstExistingKeywordBigCat.Fields("keywordorcodesection").Value
			rstBigCategory.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			With rstBigCategory
				.let_ActiveConnection(cnReadDatabase)
				.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
				.LockType = ADODB.LockTypeEnum.adLockReadOnly
				.Open(("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'"))
			End With
			If Not rstBigCategory.EOF Then
				sThesaurusText = rstThesaurusCheck.Fields("ThesaurusEquivalent").Value
				sKeywordText = rstThesaurusCheck.Fields("keywordorcodesection").Value & " (ID: " & rstThesaurusCheck.Fields("KeywordID").Value & ")"
				If InStr(1, sTitleText, sThesaurusText) Then
					For i = 0 To (Me.lstCurrentKeywords.Items.Count - 1)
						If VB6.GetItemString(Me.lstCurrentKeywords, i) = sKeywordText Then bDuplicate = True
					Next 
					For i = 1 To cSuggestedKeywords.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object cSuggestedKeywords.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
					Next 
					If Not bDuplicate Then cSuggestedKeywords.Add(sKeywordText)
				End If
			End If
			
			'UPGRADE_NOTE: Object rstBigCategory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstBigCategory = Nothing
		Loop 
		
		'UPGRADE_NOTE: Object rstExistingKeywordBigCat may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstExistingKeywordBigCat = Nothing
		'sJournalLocation = Me.txtPlaceOfPublication
		If Not (sJournalLocation = "") Then
			
			
			rstJournalKeyword = New ADODB.Recordset
			With rstJournalKeyword
				.let_ActiveConnection(cnReadDatabase)
				.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
				.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				.LockType = ADODB.LockTypeEnum.adLockReadOnly
				.Open(("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sJournalLocation & "'"))
			End With
			sKeywordText = sJournalLocation
			If Not rstJournalKeyword.EOF Then sKeywordText = sKeywordText & " (ID: " & rstJournalKeyword.Fields("KeywordID").Value & ")"
			bDuplicate = False
			For i = 0 To (Me.lstCurrentKeywords.Items.Count - 1)
				If VB6.GetItemString(Me.lstCurrentKeywords, i) = sKeywordText Then bDuplicate = True
			Next 
			If Not bDuplicate Then cSuggestedKeywords.Add(sKeywordText)
			
			'UPGRADE_NOTE: Object rstJournalKeyword may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstJournalKeyword = Nothing
		End If
		
		With rstThesaurusCheck
			.let_ActiveConnection(cnReadDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockReadOnly
			.Open(("SELECT * from qryKeywordThesaurus where (NOT (ThesaurusEquivalent IS NULL))"))
		End With
		
		Do While Not rstThesaurusCheck.EOF
			If rstThesaurusCheck.Fields("largercategory").Value = 1 Then
				rstBigCategory = New ADODB.Recordset
				sTempText = rstThesaurusCheck.Fields("keywordorcodesection").Value
				With rstBigCategory
					.let_ActiveConnection(cnReadDatabase)
					.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.LockType = ADODB.LockTypeEnum.adLockReadOnly
					.Open(("SELECT * from tblKeywords where KeyWordOrCodeSection='" & sTempText & "'"))
				End With
				If Not rstBigCategory.EOF Then
					sThesaurusText = rstThesaurusCheck.Fields("ThesaurusEquivalent").Value
					sKeywordText = rstThesaurusCheck.Fields("keywordorcodesection").Value & " (ID: " & rstThesaurusCheck.Fields("KeywordID").Value & ")"
					If InStr(1, sTitleText, sThesaurusText) Then
						For i = 0 To (Me.lstCurrentKeywords.Items.Count - 1)
							If VB6.GetItemString(Me.lstCurrentKeywords, i) = sKeywordText Then bDuplicate = True
						Next 
						For i = 1 To cSuggestedKeywords.Count()
							'UPGRADE_WARNING: Couldn't resolve default property of object cSuggestedKeywords.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
						Next 
						If Not bDuplicate Then cSuggestedKeywords.Add(sKeywordText)
					End If
				End If
				'UPGRADE_NOTE: Object rstBigCategory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstBigCategory = Nothing
			End If
			If rstThesaurusCheck.Fields("largercategory").Value = 0 Then
				bDuplicate = False
				sThesaurusText = rstThesaurusCheck.Fields("ThesaurusEquivalent").Value
				sKeywordText = rstThesaurusCheck.Fields("keywordorcodesection").Value & " (ID: " & rstThesaurusCheck.Fields("KeywordID").Value & ")"
				If InStr(1, sTitleText, sThesaurusText) Then
					For i = 0 To (Me.lstCurrentKeywords.Items.Count - 1)
						If VB6.GetItemString(Me.lstCurrentKeywords, i) = sKeywordText Then bDuplicate = True
					Next 
					For i = 1 To cSuggestedKeywords.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object cSuggestedKeywords.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If cSuggestedKeywords.Item(i) = sKeywordText Then bDuplicate = True
					Next 
					If Not bDuplicate Then cSuggestedKeywords.Add(sKeywordText)
					
					'Me.lstNewKeywords.AddItem sKeywordText
				End If
			End If
			rstThesaurusCheck.MoveNext()
		Loop 
		
		
		
		For i = 1 To cSuggestedKeywords.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object cSuggestedKeywords.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Me.lstNewKeywords.Items.Add(cSuggestedKeywords.Item(i))
		Next 
		'UPGRADE_NOTE: Object rstThesaurusCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstThesaurusCheck = Nothing
		'UPGRADE_NOTE: Object rstKeywordCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstKeywordCheck = Nothing
		'UPGRADE_NOTE: Object cSuggestedKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cSuggestedKeywords = Nothing
	End Sub
	
	Private Sub cmdNewAuthor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNewAuthor.Click
		frmNewAuthor.Close()
		frmNewAuthor.Show()
	End Sub
	
	Private Sub cmdNewJournal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNewJournal.Click
		frmNewJournal.Close()
		frmNewJournal.bEdit = False
		frmNewJournal.Text = "New Journal"
		frmNewJournal.Show()
	End Sub
	
	Private Sub cmdNewLargerWork_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNewLargerWork.Click
		frmNewLargerWork.Show()
	End Sub
	
	Private Sub cmdNextRecord_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNextRecord.Click
		Dim icounter As Short
		Dim iYN As Short
		iYN = MsgBoxResult.OK
		If Me.txtStatus.Text = "Not Saved" Then
			iYN = (MsgBox("You have made changes without saving record. Continue without saving?", MsgBoxStyle.OKCancel, "Confirm No Save"))
		End If
		If iYN = MsgBoxResult.OK Then
			icounter = (Me.cmbRecordNumber.SelectedIndex) + 1
			If icounter < Me.cmbRecordNumber.Items.Count Then
				Me.cmbRecordNumber.SelectedIndex = icounter
			End If
		End If
		'rstRecords.MoveNext
		'If rstRecords.EOF Then rstRecords.MoveLast
		'Call Erase_Form
		'Call Change_Record_Lists
		'Call Fill_Form
	End Sub
	
	Private Sub cmdPreview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreview.Click
		
		Dim report As procs
		Dim iSourceType As Short
		Dim sAuthor As String
		Dim sEditor As String
		Dim iRecordID As Short
		Dim sArticleDesignation As String
		Dim sTitle As String
		Dim sJournalTitle As String
		Dim sPage As String
		Dim sMonth As String
		Dim sDay As String
		Dim sYear As String
		Dim sVolume As String
		Dim sSeriesTitle As String
		Dim sEdition As String
		Dim sLegislativeMaterialType As String
		Dim sNameOfHouse As String
		Dim sNumberOfCongress As String
		Dim SessionOfCongress As String
		Dim sStateLegislativeSession As String
		Dim sUSCCANCitation As String
		Dim sReportOrDocumentNumber As String
		Dim sSuDocNumber As String
		
		Dim wDocument As Word.Application
		Dim i As Short
		Dim cAETIDs As Collection
		
		If IsNumeric(Me.cmbRecordNumber.Text) Then iRecordID = CShort(Me.cmbRecordNumber.Text) Else iRecordID = 0
		
		If Me.cmbSourceType.Text = "Journal Article" Then
			If Me.cmbPagination.Text = "Consecutive" Then iSourceType = 1
			If Me.cmbPagination.Text = "Nonconsecutive" Then iSourceType = 2
			'sJournalTitle = frmNewJournal.txtNewJournalShortForm.Text
			sJournalTitle = Me.txtJournaTitleShortForm.Text
			
			sVolume = Me.txtVolume.Text
		End If
		If Me.cmbSourceType.Text = "Treatise" Then iSourceType = 3
		If Me.cmbSourceType.Text = "Chapter in Treatise" Then
			iSourceType = 4
			sVolume = Me.txtSeriesVolume.Text
			sJournalTitle = Me.cmbLargerWorkTitle.Text
		End If
		If Me.cmbSourceType.Text = "Legislative Material" Then iSourceType = 5
		If Me.cmbSourceType.Text = "Unpublished Work" Then iSourceType = 7
		If Me.cmbSourceType.Text = "Nonprint Material" Then
			iSourceType = 6
			sJournalTitle = Me.txtLocation.Text
		End If
		report = New procs
		If iRecordID = 0 Then 'build a collection of AETIDs
			cAETIDs = New Collection
			For i = 1 To cAuthors.Count()
				cAETIDs.Add(cAuthors.Item(i))
			Next 
			For i = 1 To cEditors.Count()
				cAETIDs.Add(cEditors.Item(i))
			Next 
		End If
		If iRecordID <> 0 Then Call report.Get_AET_String(iRecordID, (Me.cnReadDatabase), sAuthor, sEditor, cAuthors.Count(), cEditors.Count()) Else If (cAETIDs.Count() > 0) Then Call report.Get_AET_String(iRecordID, (Me.cnReadDatabase), sAuthor, sEditor, cAuthors.Count(), cEditors.Count(), cAETIDs)
		
		sArticleDesignation = Me.cmbArticleDesignation.Text
		sTitle = Me.txtTitle.Text
		sDay = Me.txtPublicationDay.Text
		sPage = Me.txtPage.Text
		sMonth = Me.cmbPublicationMonthOrSeason.Text
		sYear = Me.txtYear.Text
		sSeriesTitle = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
		sEdition = Me.txtEditionandPrinting.Text
		wDocument = New Word.Application
		wDocument.Documents.Add()
		
		'frmWordPreview.OLEWord.Verb = -2
		
		'frmWordPreview.OLEWord.Verb = -1
		
		'frmWordPreview.OLEWord.Action = 7
		
		'frmWordPreview.OLEWord.object.Application.Visible = False
		
		'frmWordPreview.OLEWord.Verb = -2
		
		
		
		'Set wDocument = frmWordPreview.OLEWord.object.Application
		'wDocument.Visible = False
		
		
		sLegislativeMaterialType = Me.cmbLegislativeType.Text
		
		
		sNameOfHouse = Me.txtLegislativeHouse.Text
		sNumberOfCongress = Me.txtNumberOfCongress.Text
		SessionOfCongress = Me.txtNumberOfCongress.Text
		sStateLegislativeSession = Me.txtStateLegislativeSession.Text
		sUSCCANCitation = Me.txtUSCCANCitation.Text
		sReportOrDocumentNumber = Me.txtReportOrDocumentNumber.Text
		sSuDocNumber = Me.txtSuDocNumber.Text
		
		
		'    Call report.Process_Word_Line(iSourceType, sAuthor, iRecordID, sArticleDesignation, sTitle, sVolume, sJournalTitle, _
		'sPage, sMonth, sDay, sYear, sSeriesTitle, sEditor, sEdition, cEditors.Count, False, frmWordPreview.OLEWord.object.Application _
		', sLegislativeMaterialType, sNameOfHouse, sNumberOfCongress, SessionOfCongress, sStateLegislativeSession, _
		'sUSCCANCitation, sReportOrDocumentNumber, sSuDocNumber)
		
		Call report.Process_Word_Line(iSourceType, sAuthor, iRecordID, sArticleDesignation, sTitle, sVolume, sJournalTitle, sPage, sMonth, sDay, sYear, sSeriesTitle, sEditor, sEdition, cEditors.Count(), False, wDocument, sLegislativeMaterialType, sNameOfHouse, sNumberOfCongress, SessionOfCongress, sStateLegislativeSession, sUSCCANCitation, sReportOrDocumentNumber, sSuDocNumber, "")
		
		
		wDocument.Visible = True
		
		'frmWordPreview.Show
		'frmWordPreview.OLEWord.object.Application.Documents(1).Close
		'frmWordPreview.OLEWord.object.Application.Application.Quit
		'Set frmWordPreview.OLEWord.object.Application = Nothing
		
		'wDocument.Documents(1).Close
		'wDocument.Application.Quit
		'Set wDocument = Nothing
		'frmWordPreview.OLEWord.object.Application.Documents(1).Close
		'frmWordPreview.OLEWord.object.Application.Application.Quit
		'Set frmWordPreview.OLEWord.object.Application = Nothing
	End Sub
	
	Private Sub cmdPreviousRecord_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviousRecord.Click
		Dim icounter As Short
		Dim iYN As Short
		iYN = MsgBoxResult.OK
		If Me.txtStatus.Text = "Not Saved" Then
			iYN = (MsgBox("You have made changes without saving record. Continue without saving?", MsgBoxStyle.OKCancel, "Confirm No Save"))
		End If
		If iYN = MsgBoxResult.OK Then
			icounter = (Me.cmbRecordNumber.SelectedIndex) - 1
			If icounter > -1 Then
				Me.cmbRecordNumber.SelectedIndex = icounter
			End If
		End If
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim sConnectionString As String
		Dim sSourceType As String
		Dim sDateAdded As String
		Dim sDateUpdated As String
		Dim sInputInitials As String
		Dim sTitle As String
		Dim sYear As String
		Dim iLargerWorkID As Short
		Dim sArticleDesignation As String
		Dim iJournalID As Short
		Dim sPublicationDay As String
		Dim sPageNumber As String
		Dim sVolume As String
		Dim sCallNumber As String
		Dim sPublicationMonth As String
		Dim sNotes As String
		Dim sEditionAndPrinting As String
		Dim sPublisher As String
		Dim sOriginalPublicationDate As String
		Dim sSeriesVolume As String
		Dim sTitleOfSeriesIfNotIssuedByAuthor As String
		Dim bAllChaptersBySameAuthor As String
		Dim sUnpublishedWorkType As String
		Dim sThesisDissertationType As String
		Dim sLocation As String
		Dim sMiscellaneousType As String
		Dim sLegislativeType As String
		Dim sNameOfHouse As String
		Dim sNumberOfCongress As String
		Dim sSessionOfCongress As String
		Dim sStateLegislativeSession As String
		Dim sUSCCANCitation As String
		Dim sReportDocumentNumber As String
		Dim sSuDocNumber As String
		Dim rstAETDelete As ADODB.Recordset
		Dim rstKeywordDelete As ADODB.Recordset
		Dim rstBigRecordIndex As ADODB.Recordset
		Dim iRecordID As Short
		Dim icounter As Short
		'Dim lAuthorID As Long
		'Dim rstJournalCheck As ADODB.Recordset
		Dim rstCheck As ADODB.Recordset
		Dim sCheckString As String
		Dim sCheckTitle As String
		Dim bDuplicate As Boolean
		'Dim sSQL As String
		'Dim sArticleDesignation As String
		Dim iArticleID As Short
		Dim iLegislativeID As Short
		Dim iChapterID As Short
		Dim iTreatiseID As Short
		Dim iUnpublishedID As Short
		Dim iMiscID As Short
		Dim sSQLString As String
		Dim rstRecordsKeywordsThesaurus As ADODB.Recordset
		Dim rstAllKeyword As ADODB.Recordset
		Dim sAllKeywordString As String
		Dim cAllKeywords As Collection
		Dim sCurrentKeyword As String
		Dim i As Short
		Dim rstRecordsAuthors As ADODB.Recordset
		Dim rstAllAuthor As ADODB.Recordset
		Dim sFullAuthorString As String
		Dim rstAuthorCiteForm As ADODB.Recordset
		Dim iAuthorCount As Short
		Dim sAuthorString As String
		Dim rstAuthorLast As ADODB.Recordset
		Dim sAuthorLastString As String
		Dim bLibraryCollection As Boolean
		Dim iRecordsAETID As Short
		Dim sAETFMLS As String
		Dim rstAETCiteForm As ADODB.Recordset
		Dim sAuthorCiteForm As String
		Dim sEditorCiteForm As String
		Dim iEditorCount As Short
		Dim report As procs
		Dim sJournalTitle As String
		Dim sJournalTitleShortForm As String
		Dim brepublished As Boolean
		'Dim sDay As String
		report = New procs
		
		
		'Dim iAuthorCounter As Integer
		
		'Dim iRecordID As Integer
		Dim dDate As Date
		
		sSourceType = Me.cmbSourceType.Text
		sDateAdded = Me.txtDateAdded.Text
		sDateUpdated = Me.txtDateUpdated.Text
		sInputInitials = Me.txtInputInitials.Text
		sTitle = Me.txtTitle.Text
		sYear = Me.txtYear.Text
		iLargerWorkID = Val(Me.txtLargerWorkID.Text)
		sArticleDesignation = Me.cmbArticleDesignation.Text
		iJournalID = Val(Me.txtJournalID.Text)
		sPublicationDay = Me.txtPublicationDay.Text
		sPageNumber = Me.txtPage.Text
		sVolume = Me.txtVolume.Text
		sCallNumber = Me.txtCallNumber.Text
		sPublicationMonth = Me.cmbPublicationMonthOrSeason.Text
		sNotes = Me.txtNotes.Text
		sEditionAndPrinting = Me.txtEditionandPrinting.Text
		sPublisher = Me.txtPublisher.Text
		sOriginalPublicationDate = Me.txtOriginalPublicationDate.Text
		sSeriesVolume = Me.txtSeriesVolume.Text
		sTitleOfSeriesIfNotIssuedByAuthor = Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text
		bAllChaptersBySameAuthor = CStr(Me.chkAllChaptersBySameAuthor.CheckState)
		sUnpublishedWorkType = Me.cmbUnpublishedType.Text
		sThesisDissertationType = Me.cmbThesisDissertationType.Text
		sLocation = Me.txtLocation.Text
		sMiscellaneousType = Me.cmbMiscType.Text
		sLegislativeType = Me.cmbLegislativeType.Text
		sNameOfHouse = Me.txtLegislativeHouse.Text
		sNumberOfCongress = Me.txtNumberOfCongress.Text
		sSessionOfCongress = Me.txtSessionOfCongress.Text
		sStateLegislativeSession = Me.txtStateLegislativeSession.Text
		sUSCCANCitation = Me.txtUSCCANCitation.Text
		sReportDocumentNumber = Me.txtReportOrDocumentNumber.Text
		sSuDocNumber = Me.txtSuDocNumber.Text
		sJournalTitle = Me.cmbJournalTitle.Text
		sJournalTitleShortForm = Me.txtJournaTitleShortForm.Text
		bLibraryCollection = Me.chkLibraryCollection.CheckState
		brepublished = Me.chkRepublished.CheckState
		
		If (Me.txtTitle.Text = "") Or (Me.cmbSourceType.Text = "") Then
			MsgBox("Some required fields were left blank.")
			GoTo CancelErr
			
		End If
		
		Select Case Me.cmbSourceType.Text
			Case "Journal Article"
				If (Me.cmbJournalTitle.Text = "") Then
					MsgBox("Some required fields were left blank.")
					GoTo CancelErr
					
				End If
				
			Case "Nonprint Material"
				If (Me.cmbMiscType.Text = "") Then
					MsgBox("Some required fields were left blank.")
					GoTo CancelErr
					
				End If
		End Select
		If Me.txtArticleID.Text <> "" Then iArticleID = CShort(Me.txtArticleID.Text)
		If Me.txtLegislativeID.Text <> "" Then iLegislativeID = CShort(Me.txtLegislativeID.Text)
		If Me.txtChapterID.Text <> "" Then iChapterID = CShort(Me.txtChapterID.Text)
		If Me.txtTreatiseID.Text <> "" Then iTreatiseID = CShort(Me.txtTreatiseID.Text)
		If Me.txtUnpublishedID.Text <> "" Then iUnpublishedID = CShort(Me.txtUnpublishedID.Text)
		If Me.txtMiscID.Text <> "" Then iMiscID = CShort(Me.txtMiscID.Text)
		If IsNumeric(Me.cmbRecordNumber.Text) Then iRecordID = CShort(Me.cmbRecordNumber.Text)
		'sCheckTitle = Replace(sTitle, "'", "%")
		bDuplicate = False
		
		' GoTo SaveErr
		'Check to see if this is a duplicate entry
		If Me.tglNewRecords.get_Value() = True Then
			rstCheck = New ADODB.Recordset
			sCheckString = "SELECT * FROM tblrecords WHERE (PublicationYear='" & sYear & "')"
			
			If sPageNumber <> "" Then sCheckString = sCheckString & " AND (PageNumber = '" & sPageNumber & "')"
			'.CursorType = adOpenForwardOnly
			'.LockType = adLockReadOnly
			rstCheck.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCheck.Open(sCheckString, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
			Do While Not rstCheck.EOF
				If sTitle = rstCheck.Fields("Title").Value Then bDuplicate = True
				rstCheck.MoveNext()
			Loop 
			If bDuplicate Then
				MsgBox("Duplicate Record Exists. Cannot Save.", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
				
				rstCheck.Close()
				'UPGRADE_NOTE: Object rstCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstCheck = Nothing
				Exit Sub
			End If
			
			rstCheck.Close()
			'UPGRADE_NOTE: Object rstCheck may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstCheck = Nothing
		End If
		
		dDate = Now
		On Error GoTo CancelErr
		cnWriteDatabase.BeginTrans()
		If Me.tglNewRecords.get_Value() = True Then rstRecords.AddNew()
		
		If sDateAdded <> "" Then rstRecords.Fields("DateRecordAdded").Value = sDateAdded
		If (Me.tglUpdateRecords.get_Value() = True) Or (Me.tglImportRecords.get_Value() = True) Then rstRecords.Fields("dateRecordUpdated").Value = dDate
		If sInputInitials <> "" Then rstRecords.Fields("InputInitials").Value = sInputInitials
		If sSourceType <> "" Then rstRecords.Fields("DocumentType").Value = sSourceType
		If sTitle <> "" Then rstRecords.Fields("Title").Value = sTitle
		If sPageNumber <> "" Then rstRecords.Fields("PageNumber").Value = sPageNumber
		If sYear <> "" Then rstRecords.Fields("PublicationYear").Value = sYear
		If sNotes <> "" Then rstRecords.Fields("Notes").Value = sNotes
		rstRecords.Fields("LibraryCOllection").Value = Me.chkLibraryCollection.CheckState
		rstRecords.Fields("Republished").Value = Me.chkRepublished.CheckState
		
		
		rstRecords.Update()
		cnWriteDatabase.CommitTrans()
		If Me.tglNewRecords.get_Value() = True Then iRecordID = rstRecords.Fields("RecordID").Value
		Select Case sSourceType
			
			Case "Chapter in Treatise"
				rstChapters = New ADODB.Recordset
				With rstChapters
					.let_ActiveConnection(cnWriteDatabase)
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblChapters"))
				End With
				
				cnWriteDatabase.BeginTrans()
				If Me.tglNewRecords.get_Value() = True Then rstChapters.AddNew()
				
				If Me.tglNewRecords.get_Value() = False Then
					rstChapters.MoveFirst()
					Do Until rstChapters.Fields("RecordID").Value = iRecordID
						rstChapters.MoveNext()
					Loop 
				End If
				If Not rstChapters.EOF Then
					'If Me.tglUpdateRecords = True Then rstChapters!chapterID = iChapterID
					
					rstChapters.Fields("RecordID").Value = iRecordID
					rstChapters.Fields("LargerWorkID").Value = iLargerWorkID
					If sSeriesVolume <> "" Then rstChapters.Fields("SeriesVolume").Value = sSeriesVolume
					If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstChapters.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value = sTitleOfSeriesIfNotIssuedByAuthor
					
				End If
				rstChapters.Update()
				cnWriteDatabase.CommitTrans()
				rstChapters.Close()
				'UPGRADE_NOTE: Object rstChapters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstChapters = Nothing
			Case "Journal Article"
				cnWriteDatabase.BeginTrans()
				
				rstArticles = New ADODB.Recordset
				With rstArticles
					.let_ActiveConnection(cnWriteDatabase)
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblArticles"))
				End With
				
				
				dDate = Now
				If Me.tglNewRecords.get_Value() = True Then rstArticles.AddNew()
				If Me.tglNewRecords.get_Value() = False Then
					rstArticles.MoveFirst()
					Do Until rstArticles.Fields("RecordID").Value = iRecordID
						rstArticles.MoveNext()
					Loop 
				End If
				If Not rstArticles.EOF Then
					rstArticles.Fields("RecordID").Value = iRecordID
					'rstArticles!recordID = rstRecords!recordID
					
					'If Me.tglUpdateRecords = True Then rstArticles!articleID = iArticleID
					
					If sVolume <> "" Then
						rstArticles.Fields("Volume").Value = sVolume
					Else
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						rstArticles.Fields("Volume").Value = System.DBNull.Value
					End If
					If sPublicationMonth <> "" Then
						rstArticles.Fields("PublicationMonthOrSeason").Value = sPublicationMonth
					Else
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						rstArticles.Fields("PublicationMonthOrSeason").Value = System.DBNull.Value
					End If
					If sPublicationDay <> "" Then
						rstArticles.Fields("PublicationDay").Value = sPublicationDay
					Else
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						rstArticles.Fields("PublicationDay").Value = System.DBNull.Value
					End If
					If sArticleDesignation <> "" Then
						rstArticles.Fields("ArticleDesignationForCitation").Value = sArticleDesignation
					Else
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						rstArticles.Fields("ArticleDesignationForCitation").Value = System.DBNull.Value
					End If
					
					rstArticles.Fields("JournalID").Value = iJournalID
				End If
				rstArticles.Update()
				cnWriteDatabase.CommitTrans()
				rstArticles.Close()
			Case "Legislative Material"
				rstLegislativeMaterial = New ADODB.Recordset
				
				With rstLegislativeMaterial
					.let_ActiveConnection(cnWriteDatabase)
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblLegislativeMaterial"))
				End With
				
				dDate = Now
				cnWriteDatabase.BeginTrans()
				If Me.tglNewRecords.get_Value() = True Then rstLegislativeMaterial.AddNew()
				If Me.tglNewRecords.get_Value() = False Then
					rstLegislativeMaterial.MoveFirst()
					Do Until rstLegislativeMaterial.Fields("RecordID").Value = iRecordID
						rstLegislativeMaterial.MoveNext()
					Loop 
				End If
				If Not rstLegislativeMaterial.EOF Then
					rstLegislativeMaterial.Fields("RecordID").Value = iRecordID
					'If Me.tglUpdateRecords = True Then rstChapters!chapterID = iChapterID
					
					rstLegislativeMaterial.Fields("materialtype").Value = sLegislativeType
					If sNameOfHouse <> "" Then rstLegislativeMaterial.Fields("NameOfHouse").Value = sNameOfHouse
					If sLegislativeType <> "" Then rstLegislativeMaterial.Fields("materialtype").Value = sLegislativeType
					If sNumberOfCongress <> "" Then rstLegislativeMaterial.Fields("NumberOfCongress").Value = sNumberOfCongress
					If sSessionOfCongress <> "" Then rstLegislativeMaterial.Fields("SessionOfCongress").Value = sSessionOfCongress
					If sStateLegislativeSession <> "" Then rstLegislativeMaterial.Fields("StateLegislativeSession").Value = sStateLegislativeSession
					If sUSCCANCitation <> "" Then rstLegislativeMaterial.Fields("USCCANCitation").Value = sUSCCANCitation
					If sReportDocumentNumber <> "" Then rstLegislativeMaterial.Fields("ReportOrDocumentNumber").Value = sReportDocumentNumber
					If sSuDocNumber <> "" Then rstLegislativeMaterial.Fields("SuDocNumber").Value = sSuDocNumber
				End If
				rstLegislativeMaterial.Update()
				cnWriteDatabase.CommitTrans()
				rstLegislativeMaterial.Close()
				'UPGRADE_NOTE: Object rstLegislativeMaterial may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstLegislativeMaterial = Nothing
				
			Case "Treatise"
				rstTreatises = New ADODB.Recordset
				
				With rstTreatises
					.let_ActiveConnection(cnWriteDatabase)
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblTreatises"))
				End With
				
				cnWriteDatabase.BeginTrans()
				If Me.tglNewRecords.get_Value() = True Then rstTreatises.AddNew()
				If Me.tglNewRecords.get_Value() = False Then
					rstTreatises.MoveFirst()
					Do Until rstTreatises.Fields("RecordID").Value = iRecordID
						rstTreatises.MoveNext()
					Loop 
				End If
				If Not rstTreatises.EOF Then
					'If Me.tglNewRecords = True Then rstTreatises.AddNew
					rstTreatises.Fields("RecordID").Value = iRecordID
					If sEditionAndPrinting <> "" Then rstTreatises.Fields("EditionAndPrinting").Value = sEditionAndPrinting
					If sPublisher <> "" Then rstTreatises.Fields("Publisher").Value = sPublisher
					If sOriginalPublicationDate <> "" Then rstTreatises.Fields("OriginalPublicationDate").Value = sOriginalPublicationDate
					'If sSeriesVolume <> "" Then
					rstTreatises.Fields("SeriesVolume").Value = sSeriesVolume
					If sTitleOfSeriesIfNotIssuedByAuthor <> "" Then rstTreatises.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value = sTitleOfSeriesIfNotIssuedByAuthor
					If sCallNumber <> "" Then rstTreatises.Fields("CallNumber").Value = sCallNumber
				End If
				rstTreatises.Update()
				cnWriteDatabase.CommitTrans()
				rstTreatises.Close()
				'UPGRADE_NOTE: Object rstTreatises may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstTreatises = Nothing
				
			Case "Unpublished Work"
				rstUnpublishedWork = New ADODB.Recordset
				With rstUnpublishedWork
					.let_ActiveConnection(cnWriteDatabase)
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblUnpublishedWork"))
				End With
				
				cnWriteDatabase.BeginTrans()
				If Me.tglNewRecords.get_Value() = True Then rstUnpublishedWork.AddNew()
				If Me.tglNewRecords.get_Value() = False Then
					rstUnpublishedWork.MoveFirst()
					Do Until rstUnpublishedWork.Fields("RecordID").Value = iRecordID
						rstUnpublishedWork.MoveNext()
					Loop 
				End If
				If Not rstUnpublishedWork.EOF Then
					rstUnpublishedWork.Fields("RecordID").Value = rstRecords.Fields("RecordID").Value
					If sUnpublishedWorkType <> "" Then rstUnpublishedWork.Fields("Type").Value = sUnpublishedWorkType
					If sThesisDissertationType <> "" Then rstUnpublishedWork.Fields("Thesis/Dissertation Type").Value = sThesisDissertationType
					If sPublicationMonth <> "" Then rstUnpublishedWork.Fields("PublicationMonth").Value = sPublicationMonth
					If sPublicationDay <> "" Then rstUnpublishedWork.Fields("PublicationDay").Value = sPublicationDay
					If sLocation <> "" Then rstUnpublishedWork.Fields("Location").Value = sLocation
				End If
				rstUnpublishedWork.Update()
				cnWriteDatabase.CommitTrans()
				rstUnpublishedWork.Close()
				'UPGRADE_NOTE: Object rstUnpublishedWork may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstUnpublishedWork = Nothing
			Case "Nonprint Material"
				rstMisc = New ADODB.Recordset
				With rstMisc
					.let_ActiveConnection(cnWriteDatabase)
					.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
					.CursorLocation = ADODB.CursorLocationEnum.adUseClient
					.LockType = ADODB.LockTypeEnum.adLockOptimistic
					.Open(("SELECT * from tblMisc"))
				End With
				cnWriteDatabase.BeginTrans()
				
				If ((Me.tglNewRecords.get_Value() = True)) Then rstMisc.AddNew()
				If Me.tglNewRecords.get_Value() = False Then
					rstMisc.MoveFirst()
					Do Until rstMisc.Fields("RecordID").Value = iRecordID
						rstMisc.MoveNext()
					Loop 
				End If
				If Not rstMisc.EOF Then
					rstMisc.Fields("RecordID").Value = iRecordID
					If sMiscellaneousType <> "" Then rstMisc.Fields("RecordType").Value = sMiscellaneousType
					If sLocation <> "" Then rstMisc.Fields("Location").Value = sLocation
					If sPublicationMonth <> "" Then rstMisc.Fields("Month").Value = sPublicationMonth
					If sPublicationDay <> "" Then rstMisc.Fields("Day").Value = sPublicationDay
				End If
				rstMisc.Update()
				cnWriteDatabase.CommitTrans()
				rstMisc.Close()
				'UPGRADE_NOTE: Object rstMisc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rstMisc = Nothing
		End Select
		
		If (Me.tglUpdateRecords.get_Value() = True) Or (Me.tglImportRecords.get_Value() = True) Then
			rstAETDelete = New ADODB.Recordset
			rstKeywordDelete = New ADODB.Recordset
			rstAETDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstKeywordDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstAETDelete.Open("Select * from tblRecordsAET WHERE RecordID=" & iRecordID, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			rstKeywordDelete.Open("Select * from tblRecordsKeywords WHERE RecordID=" & iRecordID, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			Do While Not rstAETDelete.EOF
				rstAETDelete.Delete()
				rstAETDelete.Update()
				rstAETDelete.MoveNext()
			Loop 
			
			Do While Not rstKeywordDelete.EOF
				rstKeywordDelete.Delete()
				rstKeywordDelete.Update()
				rstKeywordDelete.MoveNext()
			Loop 
			
			rstAETDelete.Close()
			rstKeywordDelete.Close()
			'UPGRADE_NOTE: Object rstAETDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstAETDelete = Nothing
			'UPGRADE_NOTE: Object rstKeywordDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rstKeywordDelete = Nothing
			
		End If
		
		
		'If Me.tglNewRecords.Value = True Then
		rstRecordsAET = New ADODB.Recordset
		With rstRecordsAET
			.let_ActiveConnection(cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from tblRecordsAET"))
		End With
		rstRecordsAET.MoveLast()
		'iRecordsAETID = rstRecordsAET!RecordsAETID
		
		For icounter = 1 To cAuthors.Count()
			'    iRecordsAETID = iRecordsAETID + 1
			rstRecordsAET.AddNew()
			rstRecordsAET.Fields("RecordID").Value = iRecordID
			'UPGRADE_WARNING: Couldn't resolve default property of object cAuthors.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstRecordsAET.Fields("AETID").Value = cAuthors.Item(icounter)
			'        rstRecordsAET!RecordsAETID = iRecordsAETID
			rstRecordsAET.Update()
		Next 
		
		
		For icounter = 1 To cEditors.Count()
			'    iRecordsAETID = iRecordsAETID + 1
			rstRecordsAET.AddNew()
			rstRecordsAET.Fields("RecordID").Value = iRecordID
			'UPGRADE_WARNING: Couldn't resolve default property of object cEditors.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstRecordsAET.Fields("AETID").Value = cEditors.Item(icounter)
			'        rstRecordsAET!RecordsAETID = iRecordsAETID
			
			rstRecordsAET.Update()
		Next 
		
		For icounter = 1 To cTranslators.Count()
			'    iRecordsAETID = iRecordsAETID + 1
			rstRecordsAET.AddNew()
			rstRecordsAET.Fields("RecordID").Value = iRecordID
			'UPGRADE_WARNING: Couldn't resolve default property of object cTranslators.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstRecordsAET.Fields("AETID").Value = cTranslators.Item(icounter)
			'        rstRecordsAET!RecordsAETID = iRecordsAETID
			rstRecordsAET.Update()
		Next 
		
		rstRecordsAET.Close()
		'UPGRADE_NOTE: Object rstRecordsAET may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecordsAET = Nothing
		
		rstRecordsKeywords = New ADODB.Recordset
		With rstRecordsKeywords
			.let_ActiveConnection(cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from tblRecordsKeywords"))
		End With
		
		
		For icounter = 1 To cKeywords.Count()
			rstRecordsKeywords.AddNew()
			rstRecordsKeywords.Fields("RecordID").Value = iRecordID
			'UPGRADE_WARNING: Couldn't resolve default property of object cKeywords.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rstRecordsKeywords.Fields("KeywordID").Value = cKeywords.Item(icounter)
			rstRecordsKeywords.Update()
		Next 
		rstRecordsKeywords.Close()
		'UPGRADE_NOTE: Object rstRecordsKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecordsKeywords = Nothing
		
		'keyword subpart
		sSQLString = "select * from qryRecordsKeywordsThesaurus where RecordID=" & iRecordID
		rstRecordsKeywordsThesaurus = New ADODB.Recordset
		rstAllKeyword = New ADODB.Recordset
		cAllKeywords = New Collection
		rstRecordsKeywordsThesaurus.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstRecordsKeywordsThesaurus.Open(sSQLString, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		sSQLString = "select * from tblRecordsAllKeywords where RecordID=" & iRecordID
		rstAllKeyword.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstAllKeyword.Open(sSQLString, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		On Error GoTo KeywordSaveErr
		Do While Not rstRecordsKeywordsThesaurus.EOF
			'iRecordID = rstRecordsKeywordsThesaurus!RecordID
			sAllKeywordString = ""
			
			Do While iRecordID = rstRecordsKeywordsThesaurus.Fields("RecordID").Value
				bDuplicate = False
				If rstRecordsKeywordsThesaurus.Fields("keywordorcodesection").Value <> "" Then
					bDuplicate = False
					sCurrentKeyword = rstRecordsKeywordsThesaurus.Fields("keywordorcodesection").Value
					For i = 1 To cAllKeywords.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object cAllKeywords.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If cAllKeywords.Item(i) = sCurrentKeyword Then bDuplicate = True
					Next 
					If Not bDuplicate Then cAllKeywords.Add(sCurrentKeyword)
				End If
				'
				bDuplicate = False
				'
				If rstRecordsKeywordsThesaurus.Fields("ThesaurusEquivalent").Value <> "" Then
					bDuplicate = False
					sCurrentKeyword = rstRecordsKeywordsThesaurus.Fields("ThesaurusEquivalent").Value
					For i = 1 To cAllKeywords.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object cAllKeywords.Item(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If cAllKeywords.Item(i) = sCurrentKeyword Then bDuplicate = True
					Next 
					If Not bDuplicate Then cAllKeywords.Add(sCurrentKeyword)
				End If
				rstRecordsKeywordsThesaurus.MoveNext()
			Loop 
			'
KeywordEOFErr: 
			For i = 1 To cAllKeywords.Count()
				If sAllKeywordString <> "" Then sAllKeywordString = sAllKeywordString & " "
				'UPGRADE_WARNING: Couldn't resolve default property of object cAllKeywords.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sAllKeywordString = sAllKeywordString & cAllKeywords.Item(i)
			Next 
			'
			cnWriteDatabase.BeginTrans()
			
			If rstAllKeyword.EOF Then rstAllKeyword.AddNew()
			rstAllKeyword.Fields("RecordID").Value = iRecordID
			rstAllKeyword.Fields("AllKeywords").Value = sAllKeywordString
			rstAllKeyword.Update()
			cnWriteDatabase.CommitTrans()
			
			'Me.lblRecNum.Caption = "Record No. " & iRecordID & " processed."
			'Me.lblRecNum.Refresh
		Loop 
		If (Not rstAllKeyword.EOF) And cAllKeywords.Count() = 0 Then
			cnWriteDatabase.BeginTrans()
			rstAllKeyword.Delete()
			rstAllKeyword.Update()
			cnWriteDatabase.CommitTrans()
		End If
		'UPGRADE_NOTE: Object cAllKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cAllKeywords = Nothing
		'
		'cnDatabase.Close
		'Set cnDatabase = Nothing
		'Set rstRecordsAuthors = Nothing
		'UPGRADE_NOTE: Object rstRecordsKeywordsThesaurus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecordsKeywordsThesaurus = Nothing
		'Set rstAllAuthor = Nothing
		'UPGRADE_NOTE: Object rstAllKeyword may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAllKeyword = Nothing
		'end keyword part
		
		
		'author subpart
		sSQLString = "select * from qryAuthors where RecordID=" & iRecordID
		rstRecordsAuthors = New ADODB.Recordset
		rstRecordsAuthors.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstRecordsAuthors.Open(sSQLString, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		On Error GoTo AuthorSaveErr0
		
		sAuthorString = ""
		sAuthorLastString = ""
		If Not rstRecordsAuthors.EOF Then
			rstRecordsAuthors.MoveLast()
			iAuthorCount = rstRecordsAuthors.RecordCount
			rstRecordsAuthors.MoveFirst()
			sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
			Select Case iAuthorCount
				Case 1
					If rstRecordsAuthors.Fields("InstitutionalEntity").Value <> "" Then
						sAuthorString = rstRecordsAuthors.Fields("InstitutionalEntity").Value
						If sAETFMLS <> "" Then sAuthorString = sAuthorString & ", " & sAETFMLS
					Else
						If sAETFMLS <> "" Then sAuthorString = sAETFMLS
					End If
					If rstRecordsAuthors.Fields("LastName").Value <> "" Then sAuthorLastString = rstRecordsAuthors.Fields("LastName").Value
				Case 2
					If rstRecordsAuthors.Fields("InstitutionalEntity").Value <> "" Then sAuthorString = rstRecordsAuthors.Fields("InstitutionalEntity").Value & ","
					If sAETFMLS <> "" Then sAuthorString = sAETFMLS
					If rstRecordsAuthors.Fields("LastName").Value <> "" Then sAuthorLastString = rstRecordsAuthors.Fields("LastName").Value
					rstRecordsAuthors.MoveNext()
					sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
					sAuthorString = sAuthorString & " & "
					sAuthorLastString = sAuthorLastString & " " & rstRecordsAuthors.Fields("LastName").Value
					
					
					sAuthorString = sAuthorString & sAETFMLS
					
				Case Else
					If sAETFMLS <> "" Then sAuthorString = sAETFMLS & " et al."
					If rstRecordsAuthors.Fields("InstitutionalEntity").Value <> "" Then sAuthorString = rstRecordsAuthors.Fields("InstitutionalEntity").Value & ","
					If rstRecordsAuthors.Fields("LastName").Value <> "" Then sAuthorLastString = rstRecordsAuthors.Fields("LastName").Value
					rstRecordsAuthors.MoveNext()
					Do While Not rstRecordsAuthors.EOF
						sAuthorLastString = sAuthorLastString & " " & rstRecordsAuthors.Fields("LastName").Value
						rstRecordsAuthors.MoveNext()
						
					Loop 
			End Select
		End If
		
		''here we can perhaps check to see if published in a different source **republished
		
		''check to see if republished *************ADD SECTION LATER***************
		''look at qryCheckRepublished. Check to see if title is the same, and create a string variable of author
		''last names and check against the field in the query. Check to see if "republished" is checked before checking.
		''if there is a match in title and author from query, send user a message alert telling them to check the Record ID
		''and see if there is a match (because it might be a book edition or something else)
		''Even better: pop up a citation of the other record and ask if it is a republished source. If it is, mark both to republished
		
		'causing error now because Title can have a quote character in it and it messes up the SQL query. Commenting out for now.
		'If Me.tglNewRecords = True Then
		'     Set rstCheck = New ADODB.Recordset
		'     sCheckString = "SELECT * FROM qryCheckRepublished WHERE (Title='" & sTitle & "')"
		'     sCheckString = sCheckString & " AND (AllAuthorLastNameOnly = '" & sAuthorLastString & "')"
		'     rstCheck.CursorType = adOpenForwardOnly
		'     rstCheck.LockType = adLockReadOnly
		'     rstCheck.CursorLocation = adUseClient
		'     rstCheck.Open sCheckString, cnReadDatabase, adOpenForwardOnly, adLockReadOnly
		'     bDuplicate = False
		'     If Not rstCheck.EOF Then bDuplicate = True
		'     If bDuplicate Then MsgBox "This work might be republished. Please verify, go back and update records to check appropriate checkboxes if true.", vbOKOnly + vbInformation, "Alert"
		'
		'     rstCheck.Close
		'     Set rstCheck = Nothing
		'End If
		
		
		
		
AuthorEOFErr0: 
		Call report.Get_AET_String(iRecordID, (Me.cnReadDatabase), sAuthorCiteForm, sEditorCiteForm, cAuthors.Count(), cEditors.Count())
		
		rstAETCiteForm = New ADODB.Recordset
		
		With rstAETCiteForm
			.let_ActiveConnection(cnWriteDatabase)
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open("Select * from tblRecordsAETCiteForm WHERE RecordID=" & iRecordID)
		End With
		cnWriteDatabase.BeginTrans()
		
		If rstAETCiteForm.EOF Then rstAETCiteForm.AddNew()
		rstAETCiteForm.Fields("RecordID").Value = iRecordID
		rstAETCiteForm.Fields("authorciteform").Value = sAuthorCiteForm
		rstAETCiteForm.Fields("Editorciteform").Value = sEditorCiteForm
		
		rstAETCiteForm.Update()
		
		cnWriteDatabase.CommitTrans()
		
		
		rstAETCiteForm.Close()
		'UPGRADE_NOTE: Object rstAETCiteForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAETCiteForm = Nothing
		
		sSQLString = "select * from tblRecordsAuthorCiteForm where RecordID=" & iRecordID
		rstAuthorCiteForm = New ADODB.Recordset
		rstAuthorCiteForm.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstAuthorCiteForm.Open(sSQLString, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		Select Case sAuthorString
			Case ""
				If Not rstAuthorCiteForm.EOF Then
					cnWriteDatabase.BeginTrans()
					
					rstAuthorCiteForm.Delete()
					rstAuthorCiteForm.Update()
					
					cnWriteDatabase.CommitTrans()
				End If
			Case Else
				cnWriteDatabase.BeginTrans()
				
				If rstAuthorCiteForm.EOF Then rstAuthorCiteForm.AddNew()
				rstAuthorCiteForm.Fields("RecordID").Value = iRecordID
				rstAuthorCiteForm.Fields("authorciteform").Value = sAuthorString
				rstAuthorCiteForm.Update()
				
				cnWriteDatabase.CommitTrans()
				
		End Select
		rstAuthorCiteForm.Close()
		'UPGRADE_NOTE: Object rstAuthorCiteForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthorCiteForm = Nothing
		On Error GoTo AuthorSaveErr1
		
authorEOFErr1: 
		sSQLString = "select * from tblRecordsAllAuthorLastNameOnly where RecordID=" & iRecordID
		rstAuthorLast = New ADODB.Recordset
		rstAuthorLast.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstAuthorLast.Open(sSQLString, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		Select Case sAuthorLastString
			Case ""
				If Not rstAuthorLast.EOF Then
					cnWriteDatabase.BeginTrans()
					
					rstAuthorLast.Delete()
					rstAuthorLast.Update()
					
					cnWriteDatabase.CommitTrans()
					
				End If
				
			Case Else
				cnWriteDatabase.BeginTrans()
				
				If rstAuthorLast.EOF Then rstAuthorLast.AddNew()
				rstAuthorLast.Fields("RecordID").Value = iRecordID
				rstAuthorLast.Fields("AllAuthorLastNameOnly").Value = sAuthorLastString
				rstAuthorLast.Update()
				
				cnWriteDatabase.CommitTrans()
				
		End Select
		'rstAuthorLast.Update
		'cnDatabase.CommitTrans
		rstAuthorLast.Close()
		'UPGRADE_NOTE: Object rstAuthorLast may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthorLast = Nothing
		'If Not rstRecordsAuthors.EOF Then rstRecordsAuthors.MoveFirst
		If rstRecordsAuthors.RecordCount > 0 Then rstRecordsAuthors.MoveFirst()
		
		On Error GoTo AuthorSaveErr
		'Set rstRecordsAuthors = Nothing
		
		'iRecordID = rstRecordsAuthors!RecordID
		sFullAuthorString = ""
		Do While iRecordID = rstRecordsAuthors.Fields("RecordID").Value
			If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
			If rstRecordsAuthors.Fields("InstitutionalEntity").Value <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors.Fields("InstitutionalEntity").Value
			If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
			sAETFMLS = Full_AET(rstRecordsAuthors, "FMLS")
			If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
			'If rstRecordsAuthors!FMLS <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!FMLS
			sAETFMLS = Full_AET(rstRecordsAuthors, "FL")
			If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
			If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
			'If rstRecordsAuthors!FL <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!FL
			sAETFMLS = Full_AET(rstRecordsAuthors, "LFM")
			If sFullAuthorString <> "" Then sFullAuthorString = sFullAuthorString & " "
			If sAETFMLS <> "" Then sFullAuthorString = sFullAuthorString & sAETFMLS
			'If rstRecordsAuthors!LFM <> "" Then sFullAuthorString = sFullAuthorString & rstRecordsAuthors!LFM
			rstRecordsAuthors.MoveNext()
			
		Loop 
AuthorEOFErr: 
		sSQLString = "select * from tblRecordAllAuthor where RecordID=" & iRecordID
		rstAllAuthor = New ADODB.Recordset
		rstAllAuthor.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstAllAuthor.Open(sSQLString, cnWriteDatabase, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		Select Case sFullAuthorString
			Case ""
				If Not rstAllAuthor.EOF Then
					cnWriteDatabase.BeginTrans()
					
					rstAllAuthor.Delete()
					rstAllAuthor.Update()
					cnWriteDatabase.CommitTrans()
				End If
			Case Else
				cnWriteDatabase.BeginTrans()
				
				
				If rstAllAuthor.EOF Then rstAllAuthor.AddNew()
				rstAllAuthor.Fields("RecordID").Value = iRecordID
				rstAllAuthor.Fields("AllAuthors").Value = sFullAuthorString
				rstAllAuthor.Update()
				cnWriteDatabase.CommitTrans()
		End Select
		rstAllAuthor.Close()
		'UPGRADE_NOTE: Object rstAllAuthor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAllAuthor = Nothing
NoAuthor: 
		'end author part
		rstBigRecordIndex = New ADODB.Recordset
		With rstBigRecordIndex
			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.let_ActiveConnection(cnWriteDatabase)
			.Open("Select * from tblBigTextIndex WHERE RecordID=" & iRecordID)
		End With
		cnWriteDatabase.BeginTrans()
		If rstBigRecordIndex.EOF Then rstBigRecordIndex.AddNew()
		rstBigRecordIndex.Fields("RecordID").Value = iRecordID
		rstBigRecordIndex.Fields("Title").Value = sTitle
		rstBigRecordIndex.Fields("AllAuthors").Value = sFullAuthorString
		rstBigRecordIndex.Fields("AllAuthorLastNameOnly").Value = sAuthorLastString
		rstBigRecordIndex.Fields("AllKeywords").Value = sAllKeywordString
		rstBigRecordIndex.Fields("JournalTitle").Value = sJournalTitle
		rstBigRecordIndex.Fields("JournalTitleShortFOrm").Value = sJournalTitleShortForm
		rstBigRecordIndex.Update()
		cnWriteDatabase.CommitTrans()
		
		rstBigRecordIndex.Close()
		'UPGRADE_NOTE: Object rstBigRecordIndex may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstBigRecordIndex = Nothing
		If Me.tglNewRecords.get_Value() = True Then
			Me.cmbRecordNumber.Items.RemoveAt((Me.cmbRecordNumber.Items.Count - 1))
			Me.cmbRecordNumber.Items.Add(CStr(iRecordID))
			Me.cmbRecordNumber.Items.Add("New Record")
			Call Set_Entry_Form()
		End If
		Me.txtStatus.Text = "Saved"
		rstRecords.Requery()
		If Me.tglUpdateRecords.get_Value() = True Then
			If iRecordID <> rstRecords.Fields("RecordID").Value Then
				rstRecords.MoveFirst()
				Do Until rstRecords.Fields("RecordID").Value = iRecordID
					rstRecords.MoveNext()
				Loop 
			End If
		End If
		'cnDatabase.Close
		'sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\database\ncpl.mdb"
		'cnDatabase.Open (sConnectionString)
		'rstJournals.Requery
		'rstAuthors.Requery
		'rstEditors.Requery
		'rstTranslators.Requery
		'rstArticles.Requery
		'rstChapters.Requery
		'rstMisc.Requery
		'rstLegislativeMaterial.Requery
		'rstRecords.Requery
		'rstRecordsAET.Requery
		'rstRecordsKeywords.Requery
		'rstTreatises.Requery
		'rstUnpublishedWork.Requery
		
KeywordSaveErr: 
		Select Case Err.Number
			Case 3021
				Resume KeywordEOFErr
			Case 0
				
			Case Else
				cnWriteDatabase.RollbackTrans()
				MsgBox("Error#" & Err.Number & ": " & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
		End Select
AuthorSaveErr: 
		Select Case Err.Number
			Case 3021
				Resume AuthorEOFErr
				
			Case 0
				
			Case Else
				cnWriteDatabase.RollbackTrans()
				MsgBox("Error#" & Err.Number & ": " & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
		End Select
AuthorSaveErr0: 
		Select Case Err.Number
			Case 3021
				Resume AuthorEOFErr0
				
			Case 0
				
			Case Else
				cnWriteDatabase.RollbackTrans()
				MsgBox("Error#" & Err.Number & ": " & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
		End Select
AuthorSaveErr1: 
		Select Case Err.Number
			Case 3021
				Resume authorEOFErr1
			Case 0
			Case Else
				cnWriteDatabase.RollbackTrans()
				MsgBox("Error#" & Err.Number & ": " & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
		End Select
		
CancelErr: 
		Select Case Err.Number
			Case 0
			Case Else
				cnWriteDatabase.RollbackTrans()
				MsgBox("Error#" & Err.Number & ": " & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Saving Error")
		End Select
	End Sub
	
	
	
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim sConnectionString As String
		Dim dDate As Date
		'Set rstJournals = New ADODB.Recordset
		'Set rstAuthors = New ADODB.Recordset
		'Set rstEditors = New ADODB.Recordset
		'Set rstTranslators = New ADODB.Recordset
		'Set rstArticles = New ADODB.Recordset
		'Set rstChapters = New ADODB.Recordset
		'Set rstMisc = New ADODB.Recordset
		'Set rstLegislativeMaterial = New ADODB.Recordset
		'Set rstRecords = New ADODB.Recordset
		'Set rstRecordsAET = New ADODB.Recordset
		'Set rstRecordsKeywords = New ADODB.Recordset
		'Set rstTreatises = New ADODB.Recordset
		'Set rstUnpublishedWork = New ADODB.Recordset
		
		'Set rstLargerWorks = New ADODB.Recordset
		'Set rstKeywords = New ADODB.Recordset
		cnReadDatabase = New ADODB.Connection
		cnWriteDatabase = New ADODB.Connection
		
		cAuthors = New Collection
		cEditors = New Collection
		cTranslators = New Collection
		cKeywords = New Collection
		
        sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLCOPY;Data Source=NCPL" 'for local access
		'sConnectionString = "Provider=SQLOLEDB.1;Password=@boolean;Persist Security Info=True;User ID=dataentry;Initial Catalog=NCPLBETA;Data Source=128.122.192.28" 'for remote access
		
		
		cnReadDatabase.Open((sConnectionString))
		cnWriteDatabase.Open((sConnectionString))
		
		'cnDatabase.Open "Driver={SQL Server};" & _
		'"Server=NCPL;" & _
		'"Database=NCPL;" & _
		'"Uid=sa;" & _
		'"Pwd=autarchy"
		
		'frmNewJournal.Show
		'frmNewJournal.Hide
		
		Me.cmbSourceType.CausesValidation = False
		Me.tglUpdateRecords.set_Value(True)
		Me.cmbAETChoice.Text = "Author"
		'With rstJournals
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblJournals")
		'End With
		
		'With rstAuthors
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from qryAETRecords WHERE AETType='Author'")
		'End With
		
		'With rstEditors
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from qryAETRecords WHERE AETType='Editor'")
		'End With
		
		'With rstTranslators
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from qryAETRecords WHERE AETType='Translator'")
		'End With
		
		'With rstLargerWorks
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblLargerWorks")
		'End With
		
		'With rstKeywords
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblKeywords")
		'End With
		
		'With rstArticles
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblArticles")
		'End With
		
		'With rstChapters
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblChapters")
		'End With
		
		'With rstMisc
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblMisc")
		'End With
		
		'With rstLegislativeMaterial
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblLegislativeMaterial")
		'End With
		rstRecords = New ADODB.Recordset
		rstRecords.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		With rstRecords
			.let_ActiveConnection(cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from tblRecords"))
		End With
		'rstRecords.MoveFirst
		'With rstRecordsAET
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblRecordsAET")
		'End With
		
		'With rstRecordsKeywords
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblRecordsKeywords")
		'End With
		
		'With rstTreatises
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblTreatises")
		'End With
		
		'With rstUnpublishedWork
		'    .ActiveConnection = cnWriteDatabase
		'    .CursorType = adOpenKeyset
		'    .LockType = adLockOptimistic
		'    .Open ("SELECT * from tblUnpublishedWork")
		'End With
		dDate = Now
		'Me.txtDateAdded = Left(dDate, 7)
		
		Call Populate_Comboboxes()
		Call Erase_Form()
		If tglUpdateRecords.get_Value() = True Then
			Me.cmbRecordNumber.SelectedIndex = 0
			'rstRecords.MoveFirst
			'Call Fill_Form
		End If
	End Sub
	
	Private Sub Erase_Form()
		Dim sTempJournalTitle As String
		Dim iTempIndex As Short
		
		Erase_Object(lblLargerWorkID)
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Chapter in Treatise")) Then Erase_Object(txtLargerWorkID) Else Erase_Object(txtLargerWorkID, True)
		
		
		Erase_Object(lblArticleDesignation)
		Erase_Object(cmbArticleDesignation, True)
		
		'Erase_Object lblJournalID
		Erase_Object(lblJournalTitle)
		Erase_Object(lblPublicationDay)
		Erase_Object(lblVolume)
		Erase_Object(lblPublicationMonthOrSeason)
		
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object(txtJournalID) Else Erase_Object(txtJournalID, True)
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object(txtPublicationDay) Else Erase_Object(txtPublicationDay, True)
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object(txtVolume) Else Erase_Object(txtVolume, True)
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then Erase_Object(cmbPublicationMonthOrSeason) Else Erase_Object(cmbPublicationMonthOrSeason, True)
		
		
		Erase_Object((Me.lblCallNumber))
		Erase_Object((Me.txtCallNumber), True)
		
		Erase_Object(lblPage)
		Erase_Object(txtPage, True)
		
		Erase_Object(cmdEditJournal)
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
			'sTempJournalTitle = Me.cmbJournalTitle.Text
			iTempIndex = Me.cmbJournalTitle.SelectedIndex
			Erase_Object(cmbJournalTitle)
			'Me.cmbJournalTitle.Text = sTempJournalTitle
			Me.cmbJournalTitle.SelectedIndex = 0
			Me.cmbJournalTitle.SelectedIndex = iTempIndex
			Erase_Object((Me.txtJournaTitleShortForm))
			
		Else
			Erase_Object(cmbJournalTitle, True)
			Erase_Object((Me.txtJournaTitleShortForm), True)
			
		End If
		
		
		Erase_Object((Me.cmdNewLargerWork))
		
		Erase_Object((Me.chkKeepSelected))
		Erase_Object((Me.chkYear))
		'Erase_Object lblJournalTitleShortForm
		Erase_Object((Me.txtJournaTitleShortForm), True)
		
		'Erase_Object lblOrganizationIssuingNewsletter
		'Erase_Object txtOrganizationIssuingNewsletter, True
		
		'Erase_Object lblCallNumber
		'Erase_Object txtCallNumber, True
		
		'Erase_Object Me.cmdEditLargerWork
		
		'Erase_Object lblPagination
		Erase_Object(cmbPagination, True)
		
		'Erase_Object lblNotes
		'Erase_Object txtNotes, True
		
		Erase_Object(lblPage)
		Erase_Object(txtPage, True)
		
		'Erase_Object lblPlaceOfPublication
		'Erase_Object txtPlaceOfPublication, True
		
		Erase_Object(lblEditionandPrinting)
		Erase_Object(txtEditionandPrinting, True)
		
		Erase_Object(lblPublisher)
		Erase_Object(txtPublisher, True)
		
		Erase_Object(lblOriginalPublicationDate)
		Erase_Object(txtOriginalPublicationDate, True)
		
		Erase_Object(lblTitleOfSeriesIfNotIssuedByAuthor)
		Erase_Object(txtTitleOfSeriesIfNotIssuedByAuthor, True)
		
		Erase_Object(lblLocation)
		Erase_Object(txtLocation, True)
		
		Erase_Object(lblLegislativeHouse)
		Erase_Object(txtLegislativeHouse, True)
		
		Erase_Object(lblSeriesVolume)
		Erase_Object(txtSeriesVolume, True)
		
		Erase_Object(lblNumberOfCongress)
		Erase_Object(txtNumberOfCongress, True)
		
		Erase_Object(lblSessionOfCongress)
		Erase_Object(txtSessionOfCongress, True)
		
		Erase_Object(lblStateLegislativeSession)
		Erase_Object(txtStateLegislativeSession, True)
		
		Erase_Object(lblUSCCANCitation)
		Erase_Object(txtUSCCANCitation, True)
		
		Erase_Object(lblReportOrDocumentNumber)
		Erase_Object(txtReportOrDocumentNumber, True)
		
		Erase_Object(lblYear)
		If Me.chkYear.CheckState = False Then Erase_Object(txtYear, True) Else Erase_Object(txtYear)
		
		Erase_Object(lblSuDocNumber)
		Erase_Object(txtSuDocNumber, True)
		
		Erase_Object(lblUnpublishedType)
		Erase_Object(cmbUnpublishedType, True)
		
		Erase_Object(lblThesisDissertationType)
		Erase_Object(cmbThesisDissertationType, True)
		
		Erase_Object(lblMiscType)
		Erase_Object(cmbMiscType, True)
		
		Erase_Object(lblLegislativeType)
		Erase_Object(cmbLegislativeType, True)
		
		Erase_Object(lblLargerWorkTitle)
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Chapter in Treatise")) Then Erase_Object(cmbLargerWorkTitle) Else Erase_Object(cmbLargerWorkTitle, True)
		
		Erase_Object(chkAllChaptersBySameAuthor)
		
		Erase_Object(cmdNewJournal)
		
		
		
	End Sub
	
	
	
	
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'rstJournals.Close
		'rstAuthors.Close
		'rstEditors.Close
		'rstTranslators.Close
		'rstLargerWorks.Close
		'rstKeywords.Close
		'rstArticles.Close
		'rstChapters.Close
		'rstMisc.Close
		'rstLegislativeMaterial.Close
		rstRecords.Close()
		'rstRecordsAET.Close
		'rstRecordsKeywords.Close
		'rstTreatises.Close
		'rstUnpublishedWork.Close
		'
		cnReadDatabase.Close()
		cnWriteDatabase.Close()
		
		'UPGRADE_NOTE: Object rstJournals may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstJournals = Nothing
		'UPGRADE_NOTE: Object rstAuthors may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstAuthors = Nothing
		'UPGRADE_NOTE: Object rstEditors may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstEditors = Nothing
		'UPGRADE_NOTE: Object rstTranslators may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTranslators = Nothing
		'UPGRADE_NOTE: Object rstLargerWorks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstLargerWorks = Nothing
		'UPGRADE_NOTE: Object rstKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstKeywords = Nothing
		'UPGRADE_NOTE: Object cnWriteDatabase may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cnWriteDatabase = Nothing
		'UPGRADE_NOTE: Object cnReadDatabase may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cnReadDatabase = Nothing
		
		'UPGRADE_NOTE: Object rstArticles may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstArticles = Nothing
		'UPGRADE_NOTE: Object rstChapters may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstChapters = Nothing
		'UPGRADE_NOTE: Object rstMisc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstMisc = Nothing
		'UPGRADE_NOTE: Object rstLegislativeMaterial may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstLegislativeMaterial = Nothing
		'UPGRADE_NOTE: Object rstRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecords = Nothing
		'UPGRADE_NOTE: Object rstRecordsAET may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecordsAET = Nothing
		'UPGRADE_NOTE: Object rstRecordsKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstRecordsKeywords = Nothing
		'UPGRADE_NOTE: Object rstTreatises may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstTreatises = Nothing
		'UPGRADE_NOTE: Object rstUnpublishedWork may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rstUnpublishedWork = Nothing
	End Sub
	
	
	
	Private Sub lblRecordNumber_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblRecordNumber.Click
		frmJump.Show()
	End Sub
	
	Private Sub lstAuthors_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstAuthors.DoubleClick
		Call Manage_Lists(lstCurrentAuthors, lstAuthors, cAuthors)
		If cAuthors.Count() > 0 Then
			If cAuthors.Count() = 1 Then lblA.Text = "Author"
			If cAuthors.Count() > 1 Then lblA.Text = "Authors"
		Else
			lblA.Text = "No Author"
		End If
		
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub lstCurrentAuthors_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCurrentAuthors.DoubleClick
		Call Manage_Lists(lstAuthors, lstCurrentAuthors, cAuthors)
		Me.txtStatus.Text = "Not Saved"
		If cAuthors.Count() > 0 Then
			If cAuthors.Count() = 1 Then lblA.Text = "Author"
			If cAuthors.Count() > 1 Then lblA.Text = "Authors"
		Else
			lblA.Text = "No Author"
		End If
		
	End Sub
	Private Sub lstEditors_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstEditors.DoubleClick
		Call Manage_Lists(lstCurrentEditors, lstEditors, cEditors)
		If cEditors.Count() > 0 Then
			If cEditors.Count() = 1 Then lblE.Text = "Editor"
			If cEditors.Count() > 1 Then lblE.Text = "Editors"
		Else
			lblE.Text = "No Editor"
		End If
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub lstCurrentEditors_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCurrentEditors.DoubleClick
		Call Manage_Lists(lstEditors, lstCurrentEditors, cEditors)
		If cEditors.Count() > 0 Then
			If cEditors.Count() = 1 Then lblE.Text = "Editor"
			If cEditors.Count() > 1 Then lblE.Text = "Editors"
		Else
			lblE.Text = "No Editor"
		End If
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub lstNewKeywords_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstNewKeywords.DoubleClick
		Dim sSelText As String
		Dim iSelected As Short
		Dim i As Short
		
		sSelText = Me.lstNewKeywords.Text
		iSelected = Me.lstNewKeywords.SelectedIndex
		
		'For i = 0 To Me.lstKeywords.ListCount - 1
		
		'Next
		Me.lstKeywords.Text = sSelText
		Call Manage_Lists(lstCurrentKeywords, lstKeywords, cKeywords)
		Me.lstNewKeywords.Items.RemoveAt((iSelected))
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub lstTranslators_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstTranslators.DoubleClick
		Call Manage_Lists(lstCurrentTranslators, lstTranslators, cTranslators)
		If cTranslators.Count() > 0 Then
			If cTranslators.Count() = 1 Then lblT.Text = "Translator"
			If cTranslators.Count() > 1 Then lblT.Text = "Translators"
		Else
			lblT.Text = "No Translator"
		End If
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub lstCurrentTranslators_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCurrentTranslators.DoubleClick
		Call Manage_Lists(lstTranslators, lstCurrentTranslators, cTranslators)
		If cTranslators.Count() > 0 Then
			If cTranslators.Count() = 1 Then lblT.Text = "Translator"
			If cTranslators.Count() > 1 Then lblT.Text = "Translators"
		Else
			lblT.Text = "No Translator"
		End If
		Me.txtStatus.Text = "Not Saved"
	End Sub
	Private Sub lstKeywords_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstKeywords.DoubleClick
		Call Manage_Lists(lstCurrentKeywords, lstKeywords, cKeywords)
		Me.txtStatus.Text = "Not Saved"
	End Sub
	Private Sub lstCurrentKeywords_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCurrentKeywords.DoubleClick
		Call Manage_Lists(lstKeywords, lstCurrentKeywords, cKeywords)
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Public Sub Manage_Lists(ByRef oAdd As System.Windows.Forms.ListBox, ByRef oRemove As System.Windows.Forms.ListBox, ByRef cCollection As Collection, Optional ByRef iListindex As Integer = 999999)
		Dim sItem As String
		Dim iID As Short
		
		Dim iParenpos As Short
		'sItem = oRemove.Text
		If iListindex = 999999 Then iListindex = oRemove.SelectedIndex
		sItem = VB6.GetItemString(oRemove, iListindex)
		oAdd.Items.Add(sItem)
		oRemove.Items.RemoveAt((iListindex))
		If Mid(oAdd.Name, 4, 7) = "Current" Then
			iParenpos = InStr(1, sItem, " (ID: ")
			iID = Val(Mid(sItem, iParenpos + 6, Len(sItem) - (iParenpos + 6)))
		End If
		If Mid(oAdd.Name, 4, 7) = "Current" Then cCollection.Add(iID) Else cCollection.Remove((iListindex + 1))
	End Sub
	
	Private Sub Position_Object(ByRef oObject As Object, ByRef LeftPos As Short, ByRef TopPos As Short)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Left = LeftPos
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Top = TopPos
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Visible = True
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Enabled = True
	End Sub
	
	Private Sub Erase_Object(ByRef oObject As Object, Optional ByRef bErase As Boolean = False)
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Enabled = False
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oObject.Visible = False
		'UPGRADE_WARNING: Couldn't resolve default property of object oObject.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'If bErase Then oObject.Text = ""
	End Sub
	
	Private Sub Fill_Form()
		Dim rstArticlesJournals As ADODB.Recordset
		Dim rstQryKeywords As ADODB.Recordset
		Dim rstAETLMFRecords As ADODB.Recordset
		Dim rstLargerWorksChapters As ADODB.Recordset
		Dim rstLegislative As ADODB.Recordset
		Dim rstTreatise As ADODB.Recordset
		Dim rstUnpublished As ADODB.Recordset
		Dim rstOther As ADODB.Recordset
		Dim iAETID As Short
		Dim iKeywordID As Short
		Dim sSourceType As String
		Dim icounter As Short
		Dim iRecNum As Short
		Dim sCurrentAuthor As String
		Dim sCurrentKeyword As String
		
		Dim iListCount As Short
		Dim sAETType As String
		
		rstAETLMFRecords = New ADODB.Recordset
		rstQryKeywords = New ADODB.Recordset
		'If rstRecords.EOF Then rstRecords.MoveFirst
		If rstRecords.EOF Then rstRecords.MoveLast()
		
		'Me.cmbRecordNumber.Text = rstRecords!recordid
		Me.txtTitle.Text = rstRecords.Fields("Title").Value
        Me.cmbSourceType.Text = rstRecords.Fields("DocumentType").Value
        If Not IsDBNull(rstRecords("DateRecordAdded")) Then Me.txtDateAdded.Text = rstRecords.Fields("DateRecordAdded").Value
        If Not IsDBNull(rstRecords.Fields("dateRecordUpdated")) Then Me.txtDateUpdated.Text = rstRecords.Fields("dateRecordUpdated").Value
        If Not IsDBNull(rstRecords.Fields("InputInitials")) Then Me.txtInputInitials.Text = rstRecords.Fields("InputInitials").Value
        If Not IsDBNull(rstRecords.Fields("PageNumber")) Then Me.txtPage.Text = rstRecords.Fields("PageNumber").Value
        If Not IsDBNull(rstRecords.Fields("PublicationYear")) <> "" Then Me.txtYear.Text = rstRecords.Fields("PublicationYear").Value
        If Not IsDBNull(rstRecords.Fields("Notes")) <> "" Then Me.txtNotes.Text = rstRecords.Fields("Notes").Value Else Me.txtNotes.Text = ""
        If Not IsDBNull(rstRecords.Fields("LibraryCOllection")) = True Then Me.chkLibraryCollection.CheckState = System.Windows.Forms.CheckState.Checked Else Me.chkLibraryCollection.CheckState = System.Windows.Forms.CheckState.Unchecked
        If Not IsDBNull(rstRecords.Fields("Republished")) = True Then Me.chkRepublished.CheckState = System.Windows.Forms.CheckState.Checked Else Me.chkRepublished.CheckState = System.Windows.Forms.CheckState.Unchecked

        sSourceType = Me.cmbSourceType.Text
        'If Me.cmbRecordNumber.Text = "New Record" Then Me.cmbRecordNumber.Text = rtsrecords!recordid
        iRecNum = CShort(Me.cmbRecordNumber.Text)
        'iRecNum = rstRecords!recordid
        Select Case sSourceType
            Case "Chapter in Treatise"
                rstLargerWorksChapters = New ADODB.Recordset
                rstLargerWorksChapters.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstLargerWorksChapters.Open("Select * FROM qryLargerworksChapters WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstLargerWorksChapters.EOF Then
                    If rstLargerWorksChapters.Fields("LargerWorkTitle").Value <> "" Then Me.cmbLargerWorkTitle.Text = rstLargerWorksChapters.Fields("LargerWorkTitle").Value
                    Me.txtLargerWorkID.Text = rstLargerWorksChapters.Fields("LargerWorkID").Value
                    Me.txtChapterID.Text = rstLargerWorksChapters.Fields("chapterID").Value
                    If rstLargerWorksChapters.Fields("CallNumber").Value <> "" Then Me.txtCallNumber.Text = rstLargerWorksChapters.Fields("CallNumber").Value
                    If rstLargerWorksChapters.Fields("EditionAndPrinting").Value <> "" Then Me.txtEditionAndPrinting.Text = rstLargerWorksChapters.Fields("EditionAndPrinting").Value
                    If rstLargerWorksChapters.Fields("Publisher").Value <> "" Then Me.txtPublisher.Text = rstLargerWorksChapters.Fields("Publisher").Value
                    If rstLargerWorksChapters.Fields("OriginalPublicationDate").Value <> "" Then Me.txtOriginalPublicationDate.Text = rstLargerWorksChapters.Fields("OriginalPublicationDate").Value
                    If rstLargerWorksChapters.Fields("SeriesVolume").Value <> "" Then Me.txtSeriesVolume.Text = rstLargerWorksChapters.Fields("SeriesVolume").Value
                    If rstLargerWorksChapters.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = rstLargerWorksChapters.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value
                End If
                rstLargerWorksChapters.Close()
            Case "Journal Article"
                rstArticlesJournals = New ADODB.Recordset
                rstArticlesJournals.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstArticlesJournals.Open("Select * FROM qryarticlesjournals WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstArticlesJournals.EOF Then
                    Me.cmbJournalTitle.Text = rstArticlesJournals.Fields("JournalTitle").Value
                    Me.txtArticleID.Text = rstArticlesJournals.Fields("articleID").Value
                    'frmNewJournal.txtJournalID = rstArticlesJournals!JournalID
                    'frmNewJournal.txtNewJournal = rstArticlesJournals!JournalTitle
                    'frmNewJournal.txtNewJournalShortForm = rstArticlesJournals!JournalTitleShortFOrm
                    'frmNewJournal.cmbPagination.Text = rstArticlesJournals!Pagination
                    'If rstArticlesJournals!CallNumber <> Null Then frmNewJournal.txtCallNumber = rstArticlesJournals!CallNumber
                    'If rstArticlesJournals!PlaceOfPublication <> Null Then frmNewJournal.txtPlaceOfPublication = rstArticlesJournals!PlaceOfPublication

                    If rstArticlesJournals.Fields("Volume").Value <> "" Then Me.txtVolume.Text = rstArticlesJournals.Fields("Volume").Value
                    If rstArticlesJournals.Fields("PublicationMonthOrSeason").Value <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstArticlesJournals.Fields("PublicationMonthOrSeason").Value
                    If rstArticlesJournals.Fields("PublicationDay").Value <> "" Then Me.txtPublicationDay.Text = rstArticlesJournals.Fields("PublicationDay").Value
                    If rstArticlesJournals.Fields("ArticleDesignationForCitation").Value <> "" Then Me.cmbArticleDesignation.Text = rstArticlesJournals.Fields("ArticleDesignationForCitation").Value
                    Me.txtJournalID.Text = rstArticlesJournals.Fields("JournalID").Value
                    Me.txtJournaTitleShortForm.Text = rstArticlesJournals.Fields("JournalTitleShortFOrm").Value
                    'If rstArticlesJournals!JournalTitleShortForm <> "" Then Me.txtJournalTitleShortForm.Text = rstArticlesJournals!JournalTitleShortForm
                    If rstArticlesJournals.Fields("Pagination").Value <> "" Then Me.cmbPagination.Text = rstArticlesJournals.Fields("Pagination").Value
                    If rstArticlesJournals.Fields("CallNumber").Value <> "" Then Me.txtCallNumber.Text = rstArticlesJournals.Fields("CallNumber").Value
                    'If rstArticlesJournals!PLaceOfPublication <> "" Then Me.txtPlaceOfPublication = rstArticlesJournals!PLaceOfPublication
                    'Call Position_Article_Form
                End If
                rstArticlesJournals.Close()
            Case "Legislative Material"
                rstLegislative = New ADODB.Recordset
                rstLegislative.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstLegislative.Open("Select * FROM tblLegislativeMaterial WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstLegislative.EOF Then
                    Me.txtLegislativeID.Text = rstLegislative.Fields("LegislativeID").Value
                    If rstLegislative.Fields("materialtype").Value <> "" Then Me.cmbLegislativeType.Text = rstLegislative.Fields("materialtype").Value
                    If rstLegislative.Fields("NameOfHouse").Value <> "" Then Me.txtLegislativeHouse.Text = rstLegislative.Fields("NameOfHouse").Value
                    If rstLegislative.Fields("NumberOfCongress").Value <> "" Then Me.txtNumberOfCongress.Text = rstLegislative.Fields("NumberOfCongress").Value
                    If rstLegislative.Fields("SessionOfCongress").Value <> "" Then Me.txtSessionOfCongress.Text = rstLegislative.Fields("SessionOfCongress").Value
                    If rstLegislative.Fields("StateLegislativeSession").Value <> "" Then Me.txtStateLegislativeSession.Text = rstLegislative.Fields("StateLegislativeSession").Value
                    If rstLegislative.Fields("USCCANCitation").Value <> "" Then Me.txtUSCCANCitation.Text = rstLegislative.Fields("USCCANCitation").Value
                    If rstLegislative.Fields("ReportOrDocumentNumber").Value <> "" Then Me.txtReportOrDocumentNumber.Text = rstLegislative.Fields("ReportOrDocumentNumber").Value
                    If rstLegislative.Fields("SuDocNumber").Value <> "" Then Me.txtSuDocNumber.Text = rstLegislative.Fields("SuDocNumber").Value
                End If
                rstLegislative.Close()

            Case "Treatise"
                rstTreatise = New ADODB.Recordset
                rstTreatise.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstTreatise.Open("Select * FROM tblTreatises WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstTreatise.EOF Then
                    Me.txtTreatiseID.Text = rstTreatise.Fields("TreatiseID").Value
                    If rstTreatise.Fields("EditionAndPrinting").Value <> "" Then Me.txtEditionAndPrinting.Text = rstTreatise.Fields("EditionAndPrinting").Value
                    If rstTreatise.Fields("Publisher").Value <> "" Then Me.txtPublisher.Text = rstTreatise.Fields("Publisher").Value
                    If rstTreatise.Fields("OriginalPublicationDate").Value <> "" Then Me.txtOriginalPublicationDate.Text = rstTreatise.Fields("OriginalPublicationDate").Value
                    If rstTreatise.Fields("SeriesVolume").Value <> "" Then Me.txtSeriesVolume.Text = rstTreatise.Fields("SeriesVolume").Value
                    If rstTreatise.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value <> "" Then Me.txtTitleOfSeriesIfNotIssuedByAuthor.Text = rstTreatise.Fields("TitleOfSeriesIfNotIssuedByAuthor").Value
                    If rstTreatise.Fields("CallNumber").Value <> "" Then Me.txtCallNumber.Text = rstTreatise.Fields("CallNumber").Value
                End If
                rstTreatise.Close()

            Case "Unpublished Work"
                rstUnpublished = New ADODB.Recordset
                rstUnpublished.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstUnpublished.Open("Select * FROM tblUnpublishedWork WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstUnpublished.EOF Then
                    Me.txtUnpublishedID.Text = rstUnpublished.Fields("UnpublishedWorkID").Value
                    If rstUnpublished.Fields("Type").Value <> "" Then Me.cmbUnpublishedType.Text = rstUnpublished.Fields("Type").Value
                    If rstUnpublished.Fields("Thesis/Dissertation Type").Value <> "" Then Me.cmbThesisDissertationType.Text = rstUnpublished.Fields("Thesis/Dissertation Type").Value
                    If rstUnpublished.Fields("PublicationMonth").Value <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstUnpublished.Fields("PublicationMonth").Value
                    If rstUnpublished.Fields("PublicationDay").Value <> "" Then Me.txtPublicationDay.Text = rstUnpublished.Fields("PublicationDay").Value
                    If rstUnpublished.Fields("Location").Value <> "" Then Me.txtLocation.Text = rstUnpublished.Fields("Location").Value
                End If
                rstUnpublished.Close()

            Case "Nonprint Material"
                rstOther = New ADODB.Recordset
                rstOther.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                rstOther.Open("Select * FROM tblMisc WHERE RecordID = " & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                If Not rstOther.EOF Then
                    Me.txtMiscID.Text = rstOther.Fields("MiscID").Value
                    If rstOther.Fields("RecordType").Value <> "" Then Me.cmbMiscType.Text = rstOther.Fields("RecordType").Value
                    If rstOther.Fields("Location").Value <> "" Then Me.txtLocation.Text = rstOther.Fields("Location").Value
                    If rstOther.Fields("Month").Value <> "" Then Me.cmbPublicationMonthOrSeason.Text = rstOther.Fields("Month").Value
                    If rstOther.Fields("Day").Value <> "" Then Me.txtPublicationDay.Text = rstOther.Fields("Day").Value
                    'If rstOther!Location <> "" Then Me.txtLocation.Text = rstOther!Location
                End If
                rstOther.Close()
        End Select
        'rstRecordsAET.MoveFirst
        'rstAuthors.MoveFirst
        'rstAETLMFRecords.Open "SELECT * FROM qryAETLMFRecords WHERE RecordID=" & iRecNum, cnDatabase, adOpenStatic, adLockPessimistic
        rstAETLMFRecords.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rstAETLMFRecords.Open("SELECT * FROM qryAETRecords WHERE RecordID=" & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If rstAETLMFRecords.EOF Then
            lblA.Text = "No Author"
            lblE.Text = "No Editor"
            lblT.Text = "No Translator"

        End If

        Do While Not rstAETLMFRecords.EOF
            sCurrentAuthor = ""
            If rstAETLMFRecords.Fields("InstitutionalEntity").Value <> "" Then sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords.Fields("InstitutionalEntity").Value
            If rstAETLMFRecords.Fields("LastName").Value <> "" Then
                If sCurrentAuthor <> "" Then sCurrentAuthor = sCurrentAuthor & ", "
                sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords.Fields("LastName").Value
            End If
            If rstAETLMFRecords.Fields("FirstName").Value <> "" Then
                If sCurrentAuthor <> "" Then sCurrentAuthor = sCurrentAuthor & ", "
                sCurrentAuthor = sCurrentAuthor & rstAETLMFRecords.Fields("FirstName").Value
            End If
            If rstAETLMFRecords.Fields("MiddleName").Value <> "" Then sCurrentAuthor = sCurrentAuthor & " " & rstAETLMFRecords.Fields("MiddleName").Value
            If rstAETLMFRecords.Fields("Suffix").Value <> "" Then sCurrentAuthor = sCurrentAuthor & " " & rstAETLMFRecords.Fields("Suffix").Value
            sCurrentAuthor = sCurrentAuthor & " (ID: " & rstAETLMFRecords.Fields("AETID").Value & ")"

            'sCurrentAuthor = rstAETLMFRecords!FullName & " (ID: " & rstAETLMFRecords!AETID & ")"
            iAETID = rstAETLMFRecords.Fields("AETID").Value
            sAETType = rstAETLMFRecords.Fields("AETType").Value
            Select Case sAETType
                Case "Author"
                    cAuthors.Add(iAETID)
                    lstCurrentAuthors.Items.Add(sCurrentAuthor)

                    For iListCount = 0 To (lstAuthors.Items.Count - 1)
                        If VB6.GetItemString(lstAuthors, iListCount) = sCurrentAuthor Then
                            lstAuthors.Items.RemoveAt((iListCount))
                        End If
                    Next
                Case "Editor"
                    cEditors.Add(iAETID)
                    lstCurrentEditors.Items.Add(sCurrentAuthor)

                    For iListCount = 0 To (lstEditors.Items.Count - 1)
                        If VB6.GetItemString(lstEditors, iListCount) = sCurrentAuthor Then
                            lstEditors.Items.RemoveAt((iListCount))
                        End If
                    Next
                Case "Translator"
                    cTranslators.Add(iAETID)
                    lstCurrentTranslators.Items.Add(sCurrentAuthor)

                    For iListCount = 0 To (lstTranslators.Items.Count - 1)
                        If VB6.GetItemString(lstTranslators, iListCount) = sCurrentAuthor Then
                            lstTranslators.Items.RemoveAt((iListCount))
                        End If
                    Next
            End Select
            If cAuthors.Count() > 0 Then
                If cAuthors.Count() = 1 Then lblA.Text = "Author"
                If cAuthors.Count() > 1 Then lblA.Text = "Authors"
            Else
                lblA.Text = "No Author"
            End If
            If cEditors.Count() > 0 Then
                If cEditors.Count() = 1 Then lblE.Text = "Editor"
                If cEditors.Count() > 1 Then lblE.Text = "Editors"
            Else
                lblE.Text = "No Editor"
            End If
            If cTranslators.Count() > 0 Then
                If cTranslators.Count() = 1 Then lblT.Text = "Translator"
                If cTranslators.Count() > 1 Then lblT.Text = "Translators"
            Else
                lblT.Text = "No Translator"
            End If

            rstAETLMFRecords.MoveNext()
        Loop
        rstQryKeywords.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rstQryKeywords.Open("SELECT * FROM qryKeywords WHERE RecordID=" & iRecNum, cnReadDatabase, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        Do While Not rstQryKeywords.EOF
            sCurrentKeyword = rstQryKeywords.Fields("keywordorcodesection").Value & " (ID: " & rstQryKeywords.Fields("KeywordID").Value & ")"
            iKeywordID = rstQryKeywords.Fields("KeywordID").Value
            cKeywords.Add(iKeywordID)
            lstCurrentKeywords.Items.Add(sCurrentKeyword)

            For iListCount = 0 To (lstKeywords.Items.Count - 1)
                If VB6.GetItemString(lstKeywords, iListCount) = sCurrentKeyword Then
                    lstKeywords.Items.RemoveAt((iListCount))
                End If
            Next


            rstQryKeywords.MoveNext()
        Loop


        rstQryKeywords.Close()

        rstAETLMFRecords.Close()

        'Me.lstNewKeywords.Clear
        If Me.cmbRecordNumber.Text <> "" Then
            Call suggest_keywords()
        End If

        'UPGRADE_NOTE: Object rstQryKeywords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstQryKeywords = Nothing
        'UPGRADE_NOTE: Object rstArticlesJournals may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstArticlesJournals = Nothing
        'UPGRADE_NOTE: Object rstAETLMFRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstAETLMFRecords = Nothing
        'UPGRADE_NOTE: Object rstLegislative may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstLegislative = Nothing
        'UPGRADE_NOTE: Object rstOther may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstOther = Nothing
        'UPGRADE_NOTE: Object rstTreatise may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstTreatise = Nothing
        'UPGRADE_NOTE: Object rstUnpublished may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rstUnpublished = Nothing
        Me.txtStatus.Text = "Unchanged"
    End Sub
	
	Private Sub Change_Record_Lists()
		Dim icounter As Short
		For icounter = 1 To cAuthors.Count()
			'Manage_Lists lstAuthors, lstCurrentAuthors, cAuthors, (iCounter - 1)
			
			Manage_Lists(lstAuthors, lstCurrentAuthors, cAuthors, 0)
		Next 
		For icounter = 1 To cEditors.Count()
			Manage_Lists(lstEditors, lstCurrentEditors, cEditors, 0)
		Next 
		For icounter = 1 To cTranslators.Count()
			Manage_Lists(lstTranslators, lstCurrentTranslators, cTranslators, 0)
		Next 
		For icounter = 1 To cKeywords.Count()
			Manage_Lists(lstKeywords, lstCurrentKeywords, cKeywords, 0)
		Next 
		
	End Sub
	
	
	Public Sub mneNewAuthor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mneNewAuthor.Click
		Dim Index As Short = mneNewAuthor.GetIndex(eventSender)
		frmNewAuthor.Show()
	End Sub
	
	Public Sub mnuNewJournal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewJournal.Click
		Dim Index As Short = mnuNewJournal.GetIndex(eventSender)
		frmNewJournal.Show()
	End Sub
	
	Private Sub tglNewRecords_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tglNewRecords.ClickEvent
		If (tglNewRecords.get_Value() = False) And (tglUpdateRecords.get_Value() = False) And (tglImportRecords.get_Value() = False) Then tglNewRecords.set_Value(True)
		If tglNewRecords.get_Value() = True Then
			tglUpdateRecords.set_Value(False)
			tglImportRecords.set_Value(False)
			iSaveListIndex = Me.cmbRecordNumber.SelectedIndex
			Me.cmbRecordNumber.SelectedIndex = (Me.cmbRecordNumber.Items.Count - 1)
			Me.cmdSave.Text = "Save"
			Me.cmbRecordNumber.Enabled = False
		End If
		Call Set_Entry_Form()
		lblA.Text = "No Author"
		lblE.Text = "No Editor"
		lblT.Text = "No Translator"
		
	End Sub
	
	
	
	Private Sub tglUpdateRecords_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tglUpdateRecords.ClickEvent
		'If Me.tglImportRecords.Value = True Then Call Refresh_Record_List
		If (Me.tglImportRecords.get_Value() = False) And (Me.tglNewRecords.get_Value() = False) And (Me.tglUpdateRecords.get_Value() = True) Then GoTo Already_Update
		Me.cmbSourceType.CausesValidation = True
		
		If tglUpdateRecords.get_Value() = True Then
			tglNewRecords.set_Value(False)
			tglImportRecords.set_Value(False)
			Me.txtStatus.Enabled = True
		End If
		If Not (rstRecords.State = 0) And (tglUpdateRecords.get_Value() = True) Then
			'rstRecords.Requery
			Call Refresh_Record_List()
			rstRecords.MoveFirst()
			Me.cmbRecordNumber.Enabled = True
			Me.cmbRecordNumber.Text = rstRecords.Fields("RecordID").Value
		End If
		Call Change_Record_Lists()
		If (tglNewRecords.get_Value() = False) And (tglUpdateRecords.get_Value() = False) And (tglImportRecords.get_Value() = False) Then tglUpdateRecords.set_Value(True)
		'If tglNewRecords.Value = false Then Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
		If tglUpdateRecords.get_Value() = True Then
			'Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
			Me.cmdSave.Text = "Update"
			Me.chkKeepSelected.Enabled = False
			Me.chkKeepSelected.CheckState = False
			Me.chkSource.Enabled = False
			Me.chkSource.CheckState = False
			Me.chkYear.Enabled = False
			Me.chkYear.CheckState = False
			Me.cmdDelete.Enabled = True
			Me.cmbRecordNumber.Enabled = True
			Me.cmdNextRecord.Enabled = True
			Me.cmdPreviousRecord.Enabled = True
			Me.lblStatus.Visible = True
			Me.txtStatus.Visible = True
			'If Me.cmbRecordNumber.ListCount > 0 Then Me.cmbRecordNumber.ListIndex = iSaveListIndex
		End If
		
Already_Update: 
	End Sub
	
	
	Private Sub tglImportRecords_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tglImportRecords.ClickEvent
		Me.cmdSave.Text = "Save"
		If tglImportRecords.get_Value() = True Then
			tglUpdateRecords.set_Value(False)
			tglNewRecords.set_Value(False)
			'Me.cmbRecordNumber.ListIndex = (Me.cmbRecordNumber.ListCount - 1)
			Me.cmbRecordNumber.Enabled = True
			Me.cmdSave.Text = "Update"
			Me.chkKeepSelected.Enabled = False
			Me.chkKeepSelected.CheckState = False
			Me.chkSource.Enabled = False
			Me.chkSource.CheckState = False
			Me.chkYear.Enabled = False
			Me.chkYear.CheckState = False
			Me.cmdDelete.Enabled = True
			Me.cmbRecordNumber.Enabled = True
			Me.cmdNextRecord.Enabled = True
			Me.cmdPreviousRecord.Enabled = True
			Me.lblStatus.Visible = True
			Me.txtStatus.Visible = True
			'If Me.cmbRecordNumber.ListCount > 0 Then Me.cmbRecordNumber.ListIndex = iSaveListIndex
			frmFilter.Show()
			frmFilter.txtQuery.Focus()
		End If
		
		'Call Change_Record_Lists
		If (tglNewRecords.get_Value() = False) And (tglUpdateRecords.get_Value() = False) And (tglImportRecords.get_Value() = False) Then tglImportRecords.set_Value(True)
		
	End Sub
	
	Private Sub Clear_Form()
		'    Me.cmbSourceType.Text = ""
		Me.txtInputInitials.Text = ""
		Me.txtDateAdded.Text = ""
		Me.txtDateUpdated.Text = ""
		Me.txtInputInitials.Text = ""
		If Me.chkYear.CheckState = True Then Me.txtYear.Text = ""
		Me.txtTitle.Text = ""
		Me.txtArticleID.Text = ""
		Me.txtChapterID.Text = ""
		Me.txtUnpublishedID.Text = ""
		Me.txtLegislativeID.Text = ""
		Me.txtTreatiseID.Text = ""
		Me.txtMiscID.Text = ""
		Me.txtPublicationDay.Text = ""
		Me.txtNotes.Text = ""
		Me.lstNewKeywords.Items.Clear()
		Me.chkRepublished.CheckState = False
		'Me.lblA.Visible = False
		'Me.lblE.Visible = False
		'Me.lblT.Visible = False
		
	End Sub
	
	Private Sub Toolbar1_ButtonClick(ByVal Button As System.Windows.Forms.ToolStripButton)
		
	End Sub
	
	Private Sub Set_Entry_Form()
		Dim dDate As Date
		Dim iSaveCmbListIndex As Short
		Dim iTempIndex As Short
		
		iSaveListIndex = Me.cmbSourceType.SelectedIndex
		If Me.chkSource.CheckState = 0 Then iSaveListIndex = 0
		
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
			iTempIndex = Me.cmbJournalTitle.SelectedIndex
		End If
		
		Call Erase_Form()
		Call Clear_Form()
		Me.cmbSourceType.CausesValidation = True
		'Me.cmbSourceType.ListIndex = 0 'default to Journal Entry
		Me.cmbSourceType.Focus()
		dDate = Now
		
		If tglNewRecords.get_Value() = True Then
			Me.txtInputInitials.Text = "WLB"
			Me.txtDateAdded.Text = CStr(dDate)
			Me.txtStatus.Visible = False
			Me.lblStatus.Visible = False
			Me.chkKeepSelected.Enabled = True
			Me.chkSource.Enabled = True
			Me.chkYear.Enabled = True
			Me.cmbSourceType.SelectedIndex = -1 'this gets the next statement to effect a click procedure when it detects a change in value
			Me.cmbSourceType.SelectedIndex = iSaveListIndex
			'Me.cmbRecordNumber.Enabled = False
			Me.cmdNextRecord.Enabled = False
			Me.cmdPreviousRecord.Enabled = False
			Me.cmdDelete.Enabled = False
			
		End If
		
		Call Change_Record_Lists()
		If ((Me.chkKeepSelected.CheckState = 1) And (Me.cmbSourceType.Text = "Journal Article")) Then
			Me.cmbJournalTitle.SelectedIndex = -1
			Me.cmbJournalTitle.SelectedIndex = iTempIndex
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtCallNumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtCallNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCallNumber.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtEditionAndPrinting.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtEditionAndPrinting_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEditionAndPrinting.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	
	
	'UPGRADE_WARNING: Event txtLegislativeHouse.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtLegislativeHouse_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLegislativeHouse.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtLocation.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtLocation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocation.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtNotes.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtNotes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNotes.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtNumberOfCongress.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtNumberOfCongress_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberOfCongress.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtOriginalPublicationDate.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOriginalPublicationDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOriginalPublicationDate.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtPage.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPage_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPage.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtPublicationDay.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPublicationDay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPublicationDay.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtPublisher.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPublisher_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPublisher.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtReportOrDocumentNumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtReportOrDocumentNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReportOrDocumentNumber.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtSeriesVolume.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSeriesVolume_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSeriesVolume.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtSessionOfCongress.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSessionOfCongress_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSessionOfCongress.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtStateLegislativeSession.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtStateLegislativeSession_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStateLegislativeSession.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtSuDocNumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSuDocNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSuDocNumber.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtTitle.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTitle_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTitle.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtTitleOfSeriesIfNotIssuedByAuthor.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTitleOfSeriesIfNotIssuedByAuthor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTitleOfSeriesIfNotIssuedByAuthor.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtUSCCANCitation.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUSCCANCitation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUSCCANCitation.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtVolume.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtVolume_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVolume.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	'UPGRADE_WARNING: Event txtYear.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtYear_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtYear.TextChanged
		Me.txtStatus.Text = "Not Saved"
	End Sub
	
	Private Sub Refresh_Record_List()
		rstRecords.Close()
		With rstRecords
			.let_ActiveConnection(cnWriteDatabase)
			.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
			.LockType = ADODB.LockTypeEnum.adLockOptimistic
			.Open(("SELECT * from tblRecords"))
		End With
		Call Me.populate_RecordID_List()
		Me.cmbRecordNumber.SelectedIndex = 0
		
	End Sub
End Class