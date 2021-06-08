<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPrintInputs
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
    Public chkSelect(7) As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_6 As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_4 As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_1 As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_2 As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_0 As AxThreed.AxSSCheck
    Public WithEvents _chkSelect_3 As AxThreed.AxSSCheck
    '   Public WithEvents chkSelect As AxThreed.SSCheckArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrintInputs))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkSelect_6 = New AxThreed.AxSSCheck()
        Me._chkSelect_4 = New AxThreed.AxSSCheck()
        Me._chkSelect_1 = New AxThreed.AxSSCheck()
        Me._chkSelect_2 = New AxThreed.AxSSCheck()
        Me._chkSelect_0 = New AxThreed.AxSSCheck()
        Me._chkSelect_3 = New AxThreed.AxSSCheck()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me._chkSelect_6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkSelect_4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkSelect_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkSelect_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkSelect_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkSelect_3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_chkSelect_6
        '
        Me._chkSelect_6.Location = New System.Drawing.Point(8, 138)
        Me._chkSelect_6.Name = "_chkSelect_6"
        Me._chkSelect_6.OcxState = CType(resources.GetObject("_chkSelect_6.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_6.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_6.TabIndex = 1
        '
        '_chkSelect_4
        '
        Me._chkSelect_4.Location = New System.Drawing.Point(8, 98)
        Me._chkSelect_4.Name = "_chkSelect_4"
        Me._chkSelect_4.OcxState = CType(resources.GetObject("_chkSelect_4.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_4.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_4.TabIndex = 3
        '
        '_chkSelect_1
        '
        Me._chkSelect_1.Location = New System.Drawing.Point(8, 38)
        Me._chkSelect_1.Name = "_chkSelect_1"
        Me._chkSelect_1.OcxState = CType(resources.GetObject("_chkSelect_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_1.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_1.TabIndex = 4
        '
        '_chkSelect_2
        '
        Me._chkSelect_2.Location = New System.Drawing.Point(8, 58)
        Me._chkSelect_2.Name = "_chkSelect_2"
        Me._chkSelect_2.OcxState = CType(resources.GetObject("_chkSelect_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_2.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_2.TabIndex = 5
        '
        '_chkSelect_0
        '
        Me._chkSelect_0.Location = New System.Drawing.Point(8, 18)
        Me._chkSelect_0.Name = "_chkSelect_0"
        Me._chkSelect_0.OcxState = CType(resources.GetObject("_chkSelect_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_0.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_0.TabIndex = 6
        '
        '_chkSelect_3
        '
        Me._chkSelect_3.Location = New System.Drawing.Point(8, 78)
        Me._chkSelect_3.Name = "_chkSelect_3"
        Me._chkSelect_3.OcxState = CType(resources.GetObject("_chkSelect_3.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkSelect_3.Size = New System.Drawing.Size(100, 50)
        Me._chkSelect_3.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Screen No Longer Active"
        '
        'frmPrintInputs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(489, 309)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(241, 388)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPrintInputs"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Print Model Inputs"
        CType(Me._chkSelect_6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkSelect_4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkSelect_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkSelect_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkSelect_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkSelect_3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
#End Region
End Class