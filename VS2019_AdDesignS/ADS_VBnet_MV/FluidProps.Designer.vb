<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFluidProps
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




    Friend WithEvents _lblunit_0 As System.Windows.Forms.Label
    Friend WithEvents _lblunit_1 As System.Windows.Forms.Label
    Friend WithEvents _txtWater_0 As System.Windows.Forms.TextBox
    Friend WithEvents _txtWater_1 As System.Windows.Forms.TextBox
    ' Friend WithEvents Label1 As Label
    ' Friend WithEvents Label2 As Label
    ' Friend WithEvents TextBox1 As TextBox
    ' Friend WithEvents TextBox2 As TextBox
    '   Public WithEvents chkCorr As SSCheckArray
    '   Public WithEvents cmdCancelOK As SSCommandArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtWater As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblunit_0 = New System.Windows.Forms.Label()
        Me._lblunit_1 = New System.Windows.Forms.Label()
        Me.txtWater = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtWater_0 = New System.Windows.Forms.TextBox()
        Me._txtWater_1 = New System.Windows.Forms.TextBox()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me._chkCorr_0 = New System.Windows.Forms.CheckBox()
        Me._chkCorr_1 = New System.Windows.Forms.CheckBox()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWater, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(390, 30)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 12
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        '_lblunit_0
        '
        Me._lblunit_0.AutoSize = True
        Me.lblUnit.SetIndex(Me._lblunit_0, CType(0, Short))
        Me._lblunit_0.Location = New System.Drawing.Point(113, 27)
        Me._lblunit_0.Name = "_lblunit_0"
        Me._lblunit_0.Size = New System.Drawing.Size(49, 14)
        Me._lblunit_0.TabIndex = 13
        Me._lblunit_0.Text = "lblUnit(0)"
        '
        '_lblunit_1
        '
        Me._lblunit_1.AutoSize = True
        Me.lblUnit.SetIndex(Me._lblunit_1, CType(1, Short))
        Me._lblunit_1.Location = New System.Drawing.Point(113, 73)
        Me._lblunit_1.Name = "_lblunit_1"
        Me._lblunit_1.Size = New System.Drawing.Size(49, 14)
        Me._lblunit_1.TabIndex = 14
        Me._lblunit_1.Text = "lblUnit(1)"
        '
        'txtWater
        '
        '
        '_txtWater_0
        '
        Me.txtWater.SetIndex(Me._txtWater_0, CType(0, Short))
        Me._txtWater_0.Location = New System.Drawing.Point(189, 24)
        Me._txtWater_0.Name = "_txtWater_0"
        Me._txtWater_0.Size = New System.Drawing.Size(100, 20)
        Me._txtWater_0.TabIndex = 15
        Me._txtWater_0.Text = "txtWater(0)"
        '
        '_txtWater_1
        '
        Me.txtWater.SetIndex(Me._txtWater_1, CType(1, Short))
        Me._txtWater_1.Location = New System.Drawing.Point(189, 70)
        Me._txtWater_1.Name = "_txtWater_1"
        Me._txtWater_1.Size = New System.Drawing.Size(100, 20)
        Me._txtWater_1.TabIndex = 16
        Me._txtWater_1.Text = "txtWater(1)"
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(48, 148)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(100, 32)
        Me._cmdCancelOK_1.TabIndex = 17
        Me._cmdCancelOK_1.Text = "Cancel"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(189, 148)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(100, 32)
        Me._cmdCancelOK_0.TabIndex = 18
        Me._cmdCancelOK_0.Text = "OK"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        '_chkCorr_0
        '
        Me._chkCorr_0.AutoSize = True
        Me._chkCorr_0.Location = New System.Drawing.Point(12, 36)
        Me._chkCorr_0.Name = "_chkCorr_0"
        Me._chkCorr_0.Size = New System.Drawing.Size(62, 18)
        Me._chkCorr_0.TabIndex = 19
        Me._chkCorr_0.Text = "Density"
        Me._chkCorr_0.UseVisualStyleBackColor = True
        '
        '_chkCorr_1
        '
        Me._chkCorr_1.AutoSize = True
        Me._chkCorr_1.Location = New System.Drawing.Point(12, 69)
        Me._chkCorr_1.Name = "_chkCorr_1"
        Me._chkCorr_1.Size = New System.Drawing.Size(71, 18)
        Me._chkCorr_1.TabIndex = 20
        Me._chkCorr_1.Text = "Viscosity"
        Me._chkCorr_1.UseVisualStyleBackColor = True
        '
        'frmFluidProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(356, 201)
        Me.ControlBox = False
        Me.Controls.Add(Me._chkCorr_1)
        Me.Controls.Add(Me._chkCorr_0)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._txtWater_0)
        Me.Controls.Add(Me._txtWater_1)
        Me.Controls.Add(Me._lblunit_0)
        Me.Controls.Add(Me._lblunit_1)
        Me.Controls.Add(Me.Picture1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(282, 232)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(372, 240)
        Me.Name = "frmFluidProps"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "{Fluid} Properties"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWater, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Friend WithEvents _cmdCancelOK_1 As Button
    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents _chkCorr_0 As CheckBox
    Friend WithEvents _chkCorr_1 As CheckBox




#End Region
End Class