<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout2
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
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents _pnl_titleX_5 As System.Windows.Forms.PictureBox
    '   Public WithEvents pnl_title As SSPanelArray
    Public WithEvents pnl_titleX As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOK = New System.Windows.Forms.Button()
        Me._pnl_titleX_5 = New System.Windows.Forms.PictureBox()
        Me.pnl_titleX = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(Me.components)
        Me._pnl_title_3 = New System.Windows.Forms.Label()
        CType(Me._pnl_titleX_5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnl_titleX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(260, 170)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(91, 23)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        '_pnl_titleX_5
        '
        Me._pnl_titleX_5.BackColor = System.Drawing.SystemColors.Window
        Me._pnl_titleX_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._pnl_titleX_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._pnl_titleX_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._pnl_titleX_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._pnl_titleX_5.Location = New System.Drawing.Point(68, 220)
        Me._pnl_titleX_5.Name = "_pnl_titleX_5"
        Me._pnl_titleX_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._pnl_titleX_5.Size = New System.Drawing.Size(347, 49)
        Me._pnl_titleX_5.TabIndex = 0
        Me._pnl_titleX_5.TabStop = False
        Me._pnl_titleX_5.Visible = False
        '
        '_pnl_title_3
        '
        Me._pnl_title_3.Location = New System.Drawing.Point(12, 16)
        Me._pnl_title_3.Name = "_pnl_title_3"
        Me._pnl_title_3.Size = New System.Drawing.Size(331, 140)
        Me._pnl_title_3.TabIndex = 2
        Me._pnl_title_3.Text = "Label1"
        Me._pnl_title_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmAbout2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(355, 198)
        Me.Controls.Add(Me._pnl_title_3)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me._pnl_titleX_5)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(153, 453)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout2"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Technical Support Provided By:"
        CType(Me._pnl_titleX_5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnl_titleX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents _pnl_title_3 As Label
#End Region
End Class