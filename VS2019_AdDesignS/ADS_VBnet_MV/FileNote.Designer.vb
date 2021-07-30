<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFileNote
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
    Public WithEvents txtData As System.Windows.Forms.TextBox
    Public cmdButton(6) As AxThreed.AxSSCommand    'Shang add control array
    Public WithEvents lblInstructions As System.Windows.Forms.Label
    '  Public WithEvents cmdButton As SSCommandArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.txtData = New System.Windows.Forms.TextBox()
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me._cmdButton_0 = New System.Windows.Forms.Button()
        Me._cmdButton_1 = New System.Windows.Forms.Button()
        Me._cmdButton_2 = New System.Windows.Forms.Button()
        Me._cmdButton_5 = New System.Windows.Forms.Button()
        Me._cmdButton_3 = New System.Windows.Forms.Button()
        Me._cmdButton_4 = New System.Windows.Forms.Button()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(604, 211)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 9
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'txtData
        '
        Me.txtData.AcceptsReturn = True
        Me.txtData.BackColor = System.Drawing.SystemColors.Window
        Me.txtData.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtData.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtData.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtData.Location = New System.Drawing.Point(6, 34)
        Me.txtData.MaxLength = 500
        Me.txtData.Multiline = True
        Me.txtData.Name = "txtData"
        Me.txtData.ReadOnly = True
        Me.txtData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtData.Size = New System.Drawing.Size(572, 207)
        Me.txtData.TabIndex = 0
        Me.txtData.Text = "txtData"
        '
        'lblInstructions
        '
        Me.lblInstructions.BackColor = System.Drawing.SystemColors.Control
        Me.lblInstructions.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInstructions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInstructions.Location = New System.Drawing.Point(6, 2)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInstructions.Size = New System.Drawing.Size(397, 27)
        Me.lblInstructions.TabIndex = 1
        Me.lblInstructions.Text = "You may enter up to 500 characters of text.  Line breaks are acceptable."
        '
        '_cmdButton_0
        '
        Me._cmdButton_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_0.Location = New System.Drawing.Point(8, 278)
        Me._cmdButton_0.Name = "_cmdButton_0"
        Me._cmdButton_0.Size = New System.Drawing.Size(80, 39)
        Me._cmdButton_0.TabIndex = 10
        Me._cmdButton_0.Text = "&Delete"
        Me._cmdButton_0.UseVisualStyleBackColor = False
        '
        '_cmdButton_1
        '
        Me._cmdButton_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_1.Location = New System.Drawing.Point(87, 278)
        Me._cmdButton_1.Name = "_cmdButton_1"
        Me._cmdButton_1.Size = New System.Drawing.Size(82, 39)
        Me._cmdButton_1.TabIndex = 11
        Me._cmdButton_1.Text = "&Edit"
        Me._cmdButton_1.UseVisualStyleBackColor = False
        '
        '_cmdButton_2
        '
        Me._cmdButton_2.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_2.Location = New System.Drawing.Point(169, 278)
        Me._cmdButton_2.Name = "_cmdButton_2"
        Me._cmdButton_2.Size = New System.Drawing.Size(81, 39)
        Me._cmdButton_2.TabIndex = 12
        Me._cmdButton_2.Text = "&Close"
        Me._cmdButton_2.UseVisualStyleBackColor = False
        '
        '_cmdButton_5
        '
        Me._cmdButton_5.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_5.Location = New System.Drawing.Point(275, 278)
        Me._cmdButton_5.Name = "_cmdButton_5"
        Me._cmdButton_5.Size = New System.Drawing.Size(133, 40)
        Me._cmdButton_5.TabIndex = 13
        Me._cmdButton_5.Text = "Insert Date/Time"
        Me._cmdButton_5.UseVisualStyleBackColor = False
        '
        '_cmdButton_3
        '
        Me._cmdButton_3.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_3.Location = New System.Drawing.Point(408, 278)
        Me._cmdButton_3.Name = "_cmdButton_3"
        Me._cmdButton_3.Size = New System.Drawing.Size(76, 40)
        Me._cmdButton_3.TabIndex = 14
        Me._cmdButton_3.Text = "&Save"
        Me._cmdButton_3.UseVisualStyleBackColor = False
        '
        '_cmdButton_4
        '
        Me._cmdButton_4.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdButton_4.Location = New System.Drawing.Point(484, 278)
        Me._cmdButton_4.Name = "_cmdButton_4"
        Me._cmdButton_4.Size = New System.Drawing.Size(96, 40)
        Me._cmdButton_4.TabIndex = 15
        Me._cmdButton_4.Text = "C&ancel Edit"
        Me._cmdButton_4.UseVisualStyleBackColor = False
        '
        'frmFileNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(598, 331)
        Me.ControlBox = False
        Me.Controls.Add(Me._cmdButton_4)
        Me.Controls.Add(Me._cmdButton_3)
        Me.Controls.Add(Me._cmdButton_5)
        Me.Controls.Add(Me._cmdButton_2)
        Me.Controls.Add(Me._cmdButton_1)
        Me.Controls.Add(Me._cmdButton_0)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.txtData)
        Me.Controls.Add(Me.lblInstructions)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(152, 304)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(614, 370)
        Me.Name = "frmFileNote"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "File Note"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents _cmdButton_0 As Button
    Friend WithEvents _cmdButton_1 As Button
    Friend WithEvents _cmdButton_2 As Button
    Friend WithEvents _cmdButton_5 As Button
    Friend WithEvents _cmdButton_3 As Button
    Friend WithEvents _cmdButton_4 As Button
#End Region
End Class