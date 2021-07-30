<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditIsothermCAS
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

    Public WithEvents _txtData_3 As System.Windows.Forms.TextBox
    Public WithEvents _txtData_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtData_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtData_0 As System.Windows.Forms.TextBox

    Public WithEvents _lblDesc_3 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_2 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_1 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_0 As System.Windows.Forms.Label
    '   Public WithEvents chkData As AxThreed.AxSSCheckArray
    '   Public WithEvents cmdSaveCancel As AxThreed.AxSSCommandArray
    Public WithEvents lblDesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtData As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._txtData_3 = New System.Windows.Forms.TextBox()
        Me._txtData_2 = New System.Windows.Forms.TextBox()
        Me._txtData_1 = New System.Windows.Forms.TextBox()
        Me._txtData_0 = New System.Windows.Forms.TextBox()
        Me._lblDesc_3 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_0 = New System.Windows.Forms.Label()
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtData = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._cmdSaveCancel_0 = New System.Windows.Forms.Button()
        Me._cmdSaveCancel_1 = New System.Windows.Forms.Button()
        Me._chkData_0 = New System.Windows.Forms.CheckBox()
        Me._chkData_1 = New System.Windows.Forms.CheckBox()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_txtData_3
        '
        Me._txtData_3.AcceptsReturn = True
        Me._txtData_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_3, CType(3, Short))
        Me._txtData_3.Location = New System.Drawing.Point(192, 78)
        Me._txtData_3.MaxLength = 0
        Me._txtData_3.Name = "_txtData_3"
        Me._txtData_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_3.Size = New System.Drawing.Size(300, 20)
        Me._txtData_3.TabIndex = 6
        Me._txtData_3.Text = "txtData(3)"
        '
        '_txtData_2
        '
        Me._txtData_2.AcceptsReturn = True
        Me._txtData_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_2, CType(2, Short))
        Me._txtData_2.Location = New System.Drawing.Point(192, 54)
        Me._txtData_2.MaxLength = 0
        Me._txtData_2.Name = "_txtData_2"
        Me._txtData_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_2.Size = New System.Drawing.Size(159, 20)
        Me._txtData_2.TabIndex = 4
        Me._txtData_2.Text = "txtData(2)"
        '
        '_txtData_1
        '
        Me._txtData_1.AcceptsReturn = True
        Me._txtData_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_1, CType(1, Short))
        Me._txtData_1.Location = New System.Drawing.Point(192, 30)
        Me._txtData_1.MaxLength = 0
        Me._txtData_1.Name = "_txtData_1"
        Me._txtData_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_1.Size = New System.Drawing.Size(300, 20)
        Me._txtData_1.TabIndex = 2
        Me._txtData_1.Text = "txtData(1)"
        '
        '_txtData_0
        '
        Me._txtData_0.AcceptsReturn = True
        Me._txtData_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtData_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtData_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtData_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtData_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.SetIndex(Me._txtData_0, CType(0, Short))
        Me._txtData_0.Location = New System.Drawing.Point(192, 6)
        Me._txtData_0.MaxLength = 0
        Me._txtData_0.Name = "_txtData_0"
        Me._txtData_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtData_0.Size = New System.Drawing.Size(159, 20)
        Me._txtData_0.TabIndex = 0
        Me._txtData_0.Text = "txtData(0)"
        '
        '_lblDesc_3
        '
        Me._lblDesc_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblDesc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesc.SetIndex(Me._lblDesc_3, CType(3, Short))
        Me._lblDesc_3.Location = New System.Drawing.Point(8, 80)
        Me._lblDesc_3.Name = "_lblDesc_3"
        Me._lblDesc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_3.Size = New System.Drawing.Size(181, 15)
        Me._lblDesc_3.TabIndex = 7
        Me._lblDesc_3.Text = "lblDesc(3)"
        Me._lblDesc_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesc.SetIndex(Me._lblDesc_2, CType(2, Short))
        Me._lblDesc_2.Location = New System.Drawing.Point(8, 56)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(181, 15)
        Me._lblDesc_2.TabIndex = 5
        Me._lblDesc_2.Text = "lblDesc(2)"
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesc.SetIndex(Me._lblDesc_1, CType(1, Short))
        Me._lblDesc_1.Location = New System.Drawing.Point(8, 32)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(181, 15)
        Me._lblDesc_1.TabIndex = 3
        Me._lblDesc_1.Text = "lblDesc(1)"
        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_0
        '
        Me._lblDesc_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblDesc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesc.SetIndex(Me._lblDesc_0, CType(0, Short))
        Me._lblDesc_0.Location = New System.Drawing.Point(8, 8)
        Me._lblDesc_0.Name = "_lblDesc_0"
        Me._lblDesc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_0.Size = New System.Drawing.Size(181, 15)
        Me._lblDesc_0.TabIndex = 1
        Me._lblDesc_0.Text = "lblDesc(0)"
        Me._lblDesc_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtData
        '
        '
        '_cmdSaveCancel_0
        '
        Me._cmdSaveCancel_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdSaveCancel_0.Location = New System.Drawing.Point(109, 155)
        Me._cmdSaveCancel_0.Name = "_cmdSaveCancel_0"
        Me._cmdSaveCancel_0.Size = New System.Drawing.Size(116, 46)
        Me._cmdSaveCancel_0.TabIndex = 12
        Me._cmdSaveCancel_0.Text = "&Save"
        Me._cmdSaveCancel_0.UseVisualStyleBackColor = False
        '
        '_cmdSaveCancel_1
        '
        Me._cmdSaveCancel_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdSaveCancel_1.Location = New System.Drawing.Point(298, 155)
        Me._cmdSaveCancel_1.Name = "_cmdSaveCancel_1"
        Me._cmdSaveCancel_1.Size = New System.Drawing.Size(116, 46)
        Me._cmdSaveCancel_1.TabIndex = 13
        Me._cmdSaveCancel_1.Text = "&Cancel"
        Me._cmdSaveCancel_1.UseVisualStyleBackColor = False
        '
        '_chkData_0
        '
        Me._chkData_0.AutoSize = True
        Me._chkData_0.Location = New System.Drawing.Point(22, 107)
        Me._chkData_0.Name = "_chkData_0"
        Me._chkData_0.Size = New System.Drawing.Size(81, 18)
        Me._chkData_0.TabIndex = 14
        Me._chkData_0.Text = "CheckBox1"
        Me._chkData_0.UseVisualStyleBackColor = True
        '
        '_chkData_1
        '
        Me._chkData_1.AutoSize = True
        Me._chkData_1.Location = New System.Drawing.Point(22, 131)
        Me._chkData_1.Name = "_chkData_1"
        Me._chkData_1.Size = New System.Drawing.Size(81, 18)
        Me._chkData_1.TabIndex = 15
        Me._chkData_1.Text = "CheckBox2"
        Me._chkData_1.UseVisualStyleBackColor = True
        '
        'frmEditIsothermCAS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(502, 224)
        Me.ControlBox = False
        Me.Controls.Add(Me._chkData_1)
        Me.Controls.Add(Me._chkData_0)
        Me.Controls.Add(Me._cmdSaveCancel_1)
        Me.Controls.Add(Me._cmdSaveCancel_0)
        Me.Controls.Add(Me._txtData_3)
        Me.Controls.Add(Me._txtData_2)
        Me.Controls.Add(Me._txtData_1)
        Me.Controls.Add(Me._txtData_0)
        Me.Controls.Add(Me._lblDesc_3)
        Me.Controls.Add(Me._lblDesc_2)
        Me.Controls.Add(Me._lblDesc_1)
        Me.Controls.Add(Me._lblDesc_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(139, 310)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(518, 263)
        Me.Name = "frmEditIsothermCAS"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "{me.caption}"
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents _cmdSaveCancel_0 As Button
    Friend WithEvents _cmdSaveCancel_1 As Button
    Friend WithEvents _chkData_0 As CheckBox
    Friend WithEvents _chkData_1 As CheckBox
#End Region
End Class