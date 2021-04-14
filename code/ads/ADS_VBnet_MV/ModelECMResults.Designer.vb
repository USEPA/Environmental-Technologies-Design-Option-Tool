<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmModelECMResults
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
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents cboGlob As System.Windows.Forms.ComboBox
    Public WithEvents grpGlob As AxGraphLib.AxGraph
    Public WithEvents CMDialog1 As AxMSComDlg.AxCommonDialog
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModelECMResults))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.cboGlob = New System.Windows.Forms.ComboBox()
        Me.grpGlob = New AxGraphLib.AxGraph()
        Me.CMDialog1 = New AxMSComDlg.AxCommonDialog()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblData5 = New System.Windows.Forms.Label()
        Me.lblCompo = New System.Windows.Forms.Label()
        Me.lblData4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblZone = New System.Windows.Forms.Label()
        Me.lblData2 = New System.Windows.Forms.Label()
        Me.lblData3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SSFrame1 = New AxThreed.AxSSFrame()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblData1 = New System.Windows.Forms.Label()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdFile = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpGlob, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CMDialog1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSFrame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(500, 399)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(113, 22)
        Me.Command4.TabIndex = 21
        Me.Command4.Text = "Print Screen"
        Me.ToolTip1.SetToolTip(Me.Command4, "Click here to print current screen to selected printer")
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(622, 399)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 22
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'cboGlob
        '
        Me.cboGlob.BackColor = System.Drawing.SystemColors.Window
        Me.cboGlob.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGlob.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGlob.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGlob.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGlob.Location = New System.Drawing.Point(516, 194)
        Me.cboGlob.Name = "cboGlob"
        Me.cboGlob.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGlob.Size = New System.Drawing.Size(121, 22)
        Me.cboGlob.TabIndex = 19
        '
        'grpGlob
        '
        Me.grpGlob.Location = New System.Drawing.Point(4, 194)
        Me.grpGlob.Name = "grpGlob"
        Me.grpGlob.OcxState = CType(resources.GetObject("grpGlob.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grpGlob.Size = New System.Drawing.Size(490, 310)
        Me.grpGlob.TabIndex = 15
        Me.grpGlob.TabStop = False
        '
        'CMDialog1
        '
        Me.CMDialog1.Enabled = True
        Me.CMDialog1.Location = New System.Drawing.Point(0, 0)
        Me.CMDialog1.Name = "CMDialog1"
        Me.CMDialog1.OcxState = CType(resources.GetObject("CMDialog1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.CMDialog1.Size = New System.Drawing.Size(32, 32)
        Me.CMDialog1.TabIndex = 23
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(536, 14)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(87, 33)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "Mass Bal. Err. (%)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblData5
        '
        Me.lblData5.BackColor = System.Drawing.SystemColors.Window
        Me.lblData5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData5.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblData5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData5.Location = New System.Drawing.Point(530, 47)
        Me.lblData5.Name = "lblData5"
        Me.lblData5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData5.Size = New System.Drawing.Size(93, 132)
        Me.lblData5.TabIndex = 2
        Me.lblData5.Text = "lblData5"
        Me.lblData5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCompo
        '
        Me.lblCompo.BackColor = System.Drawing.SystemColors.Window
        Me.lblCompo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCompo.Location = New System.Drawing.Point(34, 47)
        Me.lblCompo.Name = "lblCompo"
        Me.lblCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompo.Size = New System.Drawing.Size(128, 132)
        Me.lblCompo.TabIndex = 3
        Me.lblCompo.Text = "lblCompo"
        '
        'lblData4
        '
        Me.lblData4.BackColor = System.Drawing.SystemColors.Window
        Me.lblData4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData4.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblData4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData4.Location = New System.Drawing.Point(441, 47)
        Me.lblData4.Name = "lblData4"
        Me.lblData4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData4.Size = New System.Drawing.Size(89, 132)
        Me.lblData4.TabIndex = 4
        Me.lblData4.Text = "lblData4"
        Me.lblData4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(438, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(89, 33)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Time to break through (days)"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblZone
        '
        Me.lblZone.BackColor = System.Drawing.SystemColors.Window
        Me.lblZone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblZone.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblZone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblZone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblZone.Location = New System.Drawing.Point(7, 47)
        Me.lblZone.Name = "lblZone"
        Me.lblZone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblZone.Size = New System.Drawing.Size(27, 132)
        Me.lblZone.TabIndex = 6
        Me.lblZone.Text = "1"
        Me.lblZone.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblData2
        '
        Me.lblData2.BackColor = System.Drawing.SystemColors.Window
        Me.lblData2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblData2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData2.Location = New System.Drawing.Point(262, 47)
        Me.lblData2.Name = "lblData2"
        Me.lblData2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData2.Size = New System.Drawing.Size(88, 132)
        Me.lblData2.TabIndex = 8
        Me.lblData2.Text = "1E-11"
        Me.lblData2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblData3
        '
        Me.lblData3.BackColor = System.Drawing.SystemColors.Window
        Me.lblData3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData3.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblData3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData3.Location = New System.Drawing.Point(350, 47)
        Me.lblData3.Name = "lblData3"
        Me.lblData3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData3.Size = New System.Drawing.Size(93, 132)
        Me.lblData3.TabIndex = 9
        Me.lblData3.Text = "111E22"
        Me.lblData3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.Location = New System.Drawing.Point(6, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(50, 17)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Zone"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(49, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(114, 20)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Components"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(262, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(89, 33)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Wave velocity (cm/s)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(354, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(94, 33)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "VTM         (mg GAC/L)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'SSFrame1
        '
        Me.SSFrame1.Controls.Add(Me.Label1)
        Me.SSFrame1.Controls.Add(Me.Label2)
        Me.SSFrame1.Controls.Add(Me.Label3)
        Me.SSFrame1.Controls.Add(Me.Label4)
        Me.SSFrame1.Controls.Add(Me.Label5)
        Me.SSFrame1.Controls.Add(Me.lblData3)
        Me.SSFrame1.Controls.Add(Me.lblData2)
        Me.SSFrame1.Controls.Add(Me.lblData1)
        Me.SSFrame1.Controls.Add(Me.lblZone)
        Me.SSFrame1.Controls.Add(Me.Label6)
        Me.SSFrame1.Controls.Add(Me.lblData4)
        Me.SSFrame1.Controls.Add(Me.lblCompo)
        Me.SSFrame1.Controls.Add(Me.lblData5)
        Me.SSFrame1.Controls.Add(Me.Label7)
        Me.SSFrame1.Location = New System.Drawing.Point(46, 0)
        Me.SSFrame1.Name = "SSFrame1"
        Me.SSFrame1.OcxState = CType(resources.GetObject("SSFrame1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame1.Size = New System.Drawing.Size(634, 188)
        Me.SSFrame1.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(162, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(106, 33)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Bed Volume Fed"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblData1
        '
        Me.lblData1.BackColor = System.Drawing.SystemColors.Window
        Me.lblData1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblData1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblData1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData1.Location = New System.Drawing.Point(162, 47)
        Me.lblData1.Name = "lblData1"
        Me.lblData1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData1.Size = New System.Drawing.Size(100, 132)
        Me.lblData1.TabIndex = 29
        Me.lblData1.Text = " "
        Me.lblData1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdSelect
        '
        Me.cmdSelect.Location = New System.Drawing.Point(516, 234)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(89, 30)
        Me.cmdSelect.TabIndex = 24
        Me.cmdSelect.Text = "Select Printer"
        Me.cmdSelect.UseVisualStyleBackColor = True
        '
        'cmdPrint
        '
        Me.cmdPrint.Location = New System.Drawing.Point(516, 270)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(89, 30)
        Me.cmdPrint.TabIndex = 25
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.UseVisualStyleBackColor = True
        '
        'cmdFile
        '
        Me.cmdFile.Location = New System.Drawing.Point(516, 306)
        Me.cmdFile.Name = "cmdFile"
        Me.cmdFile.Size = New System.Drawing.Size(89, 30)
        Me.cmdFile.TabIndex = 26
        Me.cmdFile.Text = "Save"
        Me.cmdFile.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(516, 342)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(89, 30)
        Me.cmdClose.TabIndex = 27
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmModelECMResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(750, 516)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdFile)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.cmdSelect)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.cboGlob)
        Me.Controls.Add(Me.SSFrame1)
        Me.Controls.Add(Me.grpGlob)
        Me.Controls.Add(Me.CMDialog1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(64, 96)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmModelECMResults"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Results for the Equilibrium Column Model (ECM)"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpGlob, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CMDialog1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSFrame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents Label7 As Label
    Public WithEvents lblData5 As Label
    Public WithEvents lblCompo As Label
    Public WithEvents lblData4 As Label
    Public WithEvents Label6 As Label
    Public WithEvents lblZone As Label
    Public WithEvents lblData2 As Label
    Public WithEvents lblData3 As Label
    Public WithEvents Label5 As Label
    Public WithEvents Label4 As Label
    Public WithEvents Label2 As Label
    Public WithEvents Label1 As Label
    Public WithEvents SSFrame1 As AxThreed.AxSSFrame
    Public WithEvents Label3 As Label
    Public WithEvents lblData1 As Label
    Friend WithEvents cmdSelect As Button
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdFile As Button
    Friend WithEvents cmdClose As Button
#End Region
End Class