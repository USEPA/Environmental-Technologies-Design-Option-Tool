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
    Public WithEvents cboGlob As System.Windows.Forms.ComboBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.cboGlob = New System.Windows.Forms.ComboBox()
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblData1 = New System.Windows.Forms.Label()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.cmdFile = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(700, 417)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(95, 57)
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
        Me.cboGlob.Location = New System.Drawing.Point(594, 212)
        Me.cboGlob.Name = "cboGlob"
        Me.cboGlob.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGlob.Size = New System.Drawing.Size(95, 22)
        Me.cboGlob.TabIndex = 19
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(536, 13)
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
        Me.lblData5.Location = New System.Drawing.Point(530, 46)
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
        Me.lblCompo.Location = New System.Drawing.Point(34, 46)
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
        Me.lblData4.Location = New System.Drawing.Point(441, 46)
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
        Me.Label6.Location = New System.Drawing.Point(438, 13)
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
        Me.lblZone.Location = New System.Drawing.Point(7, 46)
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
        Me.lblData2.Location = New System.Drawing.Point(262, 46)
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
        Me.lblData3.Location = New System.Drawing.Point(350, 46)
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
        Me.Label5.Location = New System.Drawing.Point(6, 26)
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
        Me.Label4.Location = New System.Drawing.Point(49, 26)
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
        Me.Label2.Location = New System.Drawing.Point(262, 13)
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
        Me.Label1.Location = New System.Drawing.Point(354, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(94, 33)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "VTM         (mg GAC/L)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(162, 13)
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
        Me.lblData1.Location = New System.Drawing.Point(162, 46)
        Me.lblData1.Name = "lblData1"
        Me.lblData1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblData1.Size = New System.Drawing.Size(100, 132)
        Me.lblData1.TabIndex = 29
        Me.lblData1.Text = " "
        Me.lblData1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdPrint.Location = New System.Drawing.Point(594, 256)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(95, 30)
        Me.cmdPrint.TabIndex = 25
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.UseVisualStyleBackColor = False
        '
        'cmdFile
        '
        Me.cmdFile.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdFile.Location = New System.Drawing.Point(594, 309)
        Me.cmdFile.Name = "cmdFile"
        Me.cmdFile.Size = New System.Drawing.Size(95, 30)
        Me.cmdFile.TabIndex = 26
        Me.cmdFile.Text = "Save to File"
        Me.cmdFile.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdClose.Location = New System.Drawing.Point(594, 360)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(95, 30)
        Me.cmdClose.TabIndex = 27
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.lblCompo)
        Me.GroupBox1.Controls.Add(Me.lblData5)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblData4)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.lblZone)
        Me.GroupBox1.Controls.Add(Me.lblData1)
        Me.GroupBox1.Controls.Add(Me.lblData2)
        Me.GroupBox1.Controls.Add(Me.lblData3)
        Me.GroupBox1.Location = New System.Drawing.Point(38, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(641, 181)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Results:"
        '
        'Chart1
        '
        ChartArea1.Area3DStyle.Enable3D = True
        ChartArea1.Area3DStyle.IsClustered = True
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Font = New System.Drawing.Font("Agency FB", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Legend1.IsTextAutoFit = False
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(12, 195)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(560, 291)
        Me.Chart1.TabIndex = 31
        Me.Chart1.Text = "Chart1"
        '
        'frmModelECMResults
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(711, 539)
        Me.ControlBox = False
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdFile)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.cboGlob)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(64, 96)
        Me.MinimumSize = New System.Drawing.Size(727, 555)
        Me.Name = "frmModelECMResults"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Results for the Equilibrium Column Model (ECM)"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Public WithEvents Label3 As Label
    Public WithEvents lblData1 As Label
    Friend WithEvents cmdPrint As Button
    Friend WithEvents cmdFile As Button
    Friend WithEvents cmdClose As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
#End Region
End Class