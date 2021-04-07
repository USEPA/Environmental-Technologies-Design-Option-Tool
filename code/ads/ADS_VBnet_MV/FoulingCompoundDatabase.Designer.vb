<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFoulingCompoundDatabase
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
	Public WithEvents lstCorrelations As System.Windows.Forms.ListBox
    Public WithEvents SSFrame1 As AxThreed.AxSSFrame
    Public WithEvents txtName As System.Windows.Forms.TextBox
    Public WithEvents _txtCoeff_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtCoeff_1 As System.Windows.Forms.TextBox
    Public WithEvents lblName As System.Windows.Forms.Label
    Public WithEvents lblCoeff2 As System.Windows.Forms.Label
    Public WithEvents lblCoeff1 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_1 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_2 As System.Windows.Forms.Label
    Public WithEvents SSFrame2 As AxThreed.AxSSFrame
    Public cmdCancelOK(2) As AxThreed.AxSSCommand
    Public cmdRecord(5) As AxThreed.AxSSCommand
    ' Public WithEvents cmdRecord As SSCommandArray
    Public WithEvents lblDesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtCoeff As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFoulingCompoundDatabase))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me.txtCoeff = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtCoeff_2 = New System.Windows.Forms.TextBox()
        Me._txtCoeff_1 = New System.Windows.Forms.TextBox()
        Me.lstCorrelations = New System.Windows.Forms.ListBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblCoeff2 = New System.Windows.Forms.Label()
        Me.lblCoeff1 = New System.Windows.Forms.Label()
        Me.SSFrame1 = New AxThreed.AxSSFrame()
        Me.SSFrame2 = New AxThreed.AxSSFrame()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me._cmdRecord_0 = New System.Windows.Forms.Button()
        Me._cmdRecord_1 = New System.Windows.Forms.Button()
        Me._cmdRecord_2 = New System.Windows.Forms.Button()
        Me._cmdRecord_3 = New System.Windows.Forms.Button()
        Me._cmdRecord_4 = New System.Windows.Forms.Button()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSFrame1.SuspendLayout()
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSFrame2.SuspendLayout()
        Me.SuspendLayout()
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_1, CType(1, Short))
        Me._lblDesc_1.Location = New System.Drawing.Point(84, 48)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(65, 17)
        Me._lblDesc_1.TabIndex = 7
        Me._lblDesc_1.Text = "A1"
        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_2, CType(2, Short))
        Me._lblDesc_2.Location = New System.Drawing.Point(84, 76)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(65, 17)
        Me._lblDesc_2.TabIndex = 6
        Me._lblDesc_2.Text = "A2"
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtCoeff
        '
        '
        '_txtCoeff_2
        '
        Me._txtCoeff_2.AcceptsReturn = True
        Me._txtCoeff_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeff.SetIndex(Me._txtCoeff_2, CType(2, Short))
        Me._txtCoeff_2.Location = New System.Drawing.Point(156, 74)
        Me._txtCoeff_2.MaxLength = 0
        Me._txtCoeff_2.Name = "_txtCoeff_2"
        Me._txtCoeff_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_2.Size = New System.Drawing.Size(81, 20)
        Me._txtCoeff_2.TabIndex = 4
        Me._txtCoeff_2.Text = "txtCoeff(2)"
        Me._txtCoeff_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtCoeff_1
        '
        Me._txtCoeff_1.AcceptsReturn = True
        Me._txtCoeff_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCoeff_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtCoeff_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCoeff_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtCoeff_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeff.SetIndex(Me._txtCoeff_1, CType(1, Short))
        Me._txtCoeff_1.Location = New System.Drawing.Point(156, 46)
        Me._txtCoeff_1.MaxLength = 0
        Me._txtCoeff_1.Name = "_txtCoeff_1"
        Me._txtCoeff_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCoeff_1.Size = New System.Drawing.Size(81, 20)
        Me._txtCoeff_1.TabIndex = 3
        Me._txtCoeff_1.Text = "txtCoeff(1)"
        Me._txtCoeff_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lstCorrelations
        '
        Me.lstCorrelations.BackColor = System.Drawing.SystemColors.Window
        Me.lstCorrelations.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstCorrelations.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCorrelations.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCorrelations.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCorrelations.ItemHeight = 14
        Me.lstCorrelations.Location = New System.Drawing.Point(8, 20)
        Me.lstCorrelations.Name = "lstCorrelations"
        Me.lstCorrelations.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCorrelations.Size = New System.Drawing.Size(365, 142)
        Me.lstCorrelations.TabIndex = 2
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(10, 18)
        Me.txtName.MaxLength = 0
        Me.txtName.Name = "txtName"
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(347, 20)
        Me.txtName.TabIndex = 5
        Me.txtName.Text = "txtName"
        Me.txtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.Color.Transparent
        Me.lblName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.Location = New System.Drawing.Point(10, 18)
        Me.lblName.Name = "lblName"
        Me.lblName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblName.Size = New System.Drawing.Size(341, 17)
        Me.lblName.TabIndex = 10
        Me.lblName.Text = "lblName"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCoeff2
        '
        Me.lblCoeff2.BackColor = System.Drawing.SystemColors.Window
        Me.lblCoeff2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCoeff2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCoeff2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCoeff2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCoeff2.Location = New System.Drawing.Point(156, 74)
        Me.lblCoeff2.Name = "lblCoeff2"
        Me.lblCoeff2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCoeff2.Size = New System.Drawing.Size(81, 17)
        Me.lblCoeff2.TabIndex = 9
        Me.lblCoeff2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCoeff1
        '
        Me.lblCoeff1.BackColor = System.Drawing.SystemColors.Window
        Me.lblCoeff1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCoeff1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCoeff1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCoeff1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCoeff1.Location = New System.Drawing.Point(156, 46)
        Me.lblCoeff1.Name = "lblCoeff1"
        Me.lblCoeff1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCoeff1.Size = New System.Drawing.Size(81, 17)
        Me.lblCoeff1.TabIndex = 8
        Me.lblCoeff1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'SSFrame1
        '
        Me.SSFrame1.Controls.Add(Me.lstCorrelations)
        Me.SSFrame1.Location = New System.Drawing.Point(8, 6)
        Me.SSFrame1.Name = "SSFrame1"
        Me.SSFrame1.OcxState = CType(resources.GetObject("SSFrame1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame1.Size = New System.Drawing.Size(401, 184)
        Me.SSFrame1.TabIndex = 0
        '
        'SSFrame2
        '
        Me.SSFrame2.Controls.Add(Me.txtName)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_1)
        Me.SSFrame2.Controls.Add(Me._txtCoeff_2)
        Me.SSFrame2.Controls.Add(Me._lblDesc_1)
        Me.SSFrame2.Controls.Add(Me._lblDesc_2)
        Me.SSFrame2.Location = New System.Drawing.Point(12, 196)
        Me.SSFrame2.Name = "SSFrame2"
        Me.SSFrame2.OcxState = CType(resources.GetObject("SSFrame2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SSFrame2.Size = New System.Drawing.Size(397, 163)
        Me.SSFrame2.TabIndex = 1
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.Color.White
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(83, 363)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(104, 34)
        Me._cmdCancelOK_1.TabIndex = 35
        Me._cmdCancelOK_1.Text = "OK"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.Color.White
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(228, 363)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(104, 34)
        Me._cmdCancelOK_0.TabIndex = 36
        Me._cmdCancelOK_0.Text = "Cancel"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        '_cmdRecord_0
        '
        Me._cmdRecord_0.BackColor = System.Drawing.Color.White
        Me._cmdRecord_0.Location = New System.Drawing.Point(17, 309)
        Me._cmdRecord_0.Name = "_cmdRecord_0"
        Me._cmdRecord_0.Size = New System.Drawing.Size(67, 31)
        Me._cmdRecord_0.TabIndex = 37
        Me._cmdRecord_0.Text = "&New"
        Me._cmdRecord_0.UseVisualStyleBackColor = False
        '
        '_cmdRecord_1
        '
        Me._cmdRecord_1.BackColor = System.Drawing.Color.White
        Me._cmdRecord_1.Location = New System.Drawing.Point(90, 309)
        Me._cmdRecord_1.Name = "_cmdRecord_1"
        Me._cmdRecord_1.Size = New System.Drawing.Size(72, 31)
        Me._cmdRecord_1.TabIndex = 38
        Me._cmdRecord_1.Text = "&Edit"
        Me._cmdRecord_1.UseVisualStyleBackColor = False
        '
        '_cmdRecord_2
        '
        Me._cmdRecord_2.BackColor = System.Drawing.Color.White
        Me._cmdRecord_2.Location = New System.Drawing.Point(168, 309)
        Me._cmdRecord_2.Name = "_cmdRecord_2"
        Me._cmdRecord_2.Size = New System.Drawing.Size(72, 31)
        Me._cmdRecord_2.TabIndex = 39
        Me._cmdRecord_2.Text = "&Delete"
        Me._cmdRecord_2.UseVisualStyleBackColor = False
        '
        '_cmdRecord_3
        '
        Me._cmdRecord_3.BackColor = System.Drawing.Color.White
        Me._cmdRecord_3.Location = New System.Drawing.Point(246, 309)
        Me._cmdRecord_3.Name = "_cmdRecord_3"
        Me._cmdRecord_3.Size = New System.Drawing.Size(72, 31)
        Me._cmdRecord_3.TabIndex = 40
        Me._cmdRecord_3.Text = "&Save"
        Me._cmdRecord_3.UseVisualStyleBackColor = False
        '
        '_cmdRecord_4
        '
        Me._cmdRecord_4.BackColor = System.Drawing.Color.White
        Me._cmdRecord_4.Location = New System.Drawing.Point(324, 309)
        Me._cmdRecord_4.Name = "_cmdRecord_4"
        Me._cmdRecord_4.Size = New System.Drawing.Size(73, 31)
        Me._cmdRecord_4.TabIndex = 41
        Me._cmdRecord_4.Text = "Cancel Edit"
        Me._cmdRecord_4.UseVisualStyleBackColor = False
        '
        'frmFoulingCompoundDatabase
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(421, 409)
        Me.ControlBox = False
        Me.Controls.Add(Me._cmdRecord_4)
        Me.Controls.Add(Me._cmdRecord_3)
        Me.Controls.Add(Me._cmdRecord_2)
        Me.Controls.Add(Me._cmdRecord_1)
        Me.Controls.Add(Me._cmdRecord_0)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me.SSFrame1)
        Me.Controls.Add(Me.SSFrame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(186, 131)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFoulingCompoundDatabase"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Water Fouling Compound Correlation Database"
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCoeff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSFrame1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSFrame1.ResumeLayout(False)
        CType(Me.SSFrame2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSFrame2.ResumeLayout(False)
        Me.SSFrame2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents _cmdCancelOK_1 As Button
    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents _cmdRecord_0 As Button
    Friend WithEvents _cmdRecord_1 As Button
    Friend WithEvents _cmdRecord_2 As Button
    Friend WithEvents _cmdRecord_3 As Button
    Friend WithEvents _cmdRecord_4 As Button



#End Region
End Class