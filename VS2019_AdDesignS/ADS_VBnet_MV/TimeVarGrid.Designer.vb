<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTimeVarGrid
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
	Public WithEvents _mnuFileItem_10 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_49 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _mnuFileItem_50 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_60 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_70 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_99 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _mnuFileItem_100 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuEditItem_10 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuEditItem_20 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents _cboUnits_1 As System.Windows.Forms.ComboBox
	Public WithEvents _cboUnits_0 As System.Windows.Forms.ComboBox
    Public WithEvents _lblData_1 As System.Windows.Forms.Label
    Public WithEvents _lblData_0 As System.Windows.Forms.Label
    Public WithEvents cboUnits As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
    '   Public WithEvents cmdCancelOK As SSCommandArray
    Public WithEvents lblData As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents mnuEditItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuFileItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_49 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFileItem_50 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_60 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_70 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_99 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFileItem_100 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_20 = New System.Windows.Forms.ToolStripMenuItem()
        Me._cboUnits_1 = New System.Windows.Forms.ComboBox()
        Me._cboUnits_0 = New System.Windows.Forms.ComboBox()
        Me._lblData_1 = New System.Windows.Forms.Label()
        Me._lblData_0 = New System.Windows.Forms.Label()
        Me.cboUnits = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(Me.components)
        Me.lblData = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.mnuEditItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuFileItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me.MainMenu1.SuspendLayout()
        CType(Me.cboUnits, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuEditItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFileItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(496, 24)
        Me.MainMenu1.TabIndex = 11
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFileItem_10, Me._mnuFileItem_49, Me._mnuFileItem_50, Me._mnuFileItem_60, Me._mnuFileItem_70, Me._mnuFileItem_99, Me._mnuFileItem_100})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        '_mnuFileItem_10
        '
        Me._mnuFileItem_10.Enabled = False
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_10, CType(10, Short))
        Me._mnuFileItem_10.Name = "_mnuFileItem_10"
        Me._mnuFileItem_10.Size = New System.Drawing.Size(154, 22)
        Me._mnuFileItem_10.Text = "Save &As ..."
        '
        '_mnuFileItem_49
        '
        Me._mnuFileItem_49.Name = "_mnuFileItem_49"
        Me._mnuFileItem_49.Size = New System.Drawing.Size(151, 6)
        '
        '_mnuFileItem_50
        '
        Me._mnuFileItem_50.Enabled = False
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_50, CType(50, Short))
        Me._mnuFileItem_50.Name = "_mnuFileItem_50"
        Me._mnuFileItem_50.Size = New System.Drawing.Size(154, 22)
        Me._mnuFileItem_50.Text = "Page Setup ..."
        '
        '_mnuFileItem_60
        '
        Me._mnuFileItem_60.Enabled = False
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_60, CType(60, Short))
        Me._mnuFileItem_60.Name = "_mnuFileItem_60"
        Me._mnuFileItem_60.Size = New System.Drawing.Size(154, 22)
        Me._mnuFileItem_60.Text = "Printer Setup ..."
        '
        '_mnuFileItem_70
        '
        Me._mnuFileItem_70.Enabled = False
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_70, CType(70, Short))
        Me._mnuFileItem_70.Name = "_mnuFileItem_70"
        Me._mnuFileItem_70.Size = New System.Drawing.Size(154, 22)
        Me._mnuFileItem_70.Text = "&Print ..."
        '
        '_mnuFileItem_99
        '
        Me._mnuFileItem_99.Name = "_mnuFileItem_99"
        Me._mnuFileItem_99.Size = New System.Drawing.Size(151, 6)
        Me._mnuFileItem_99.Visible = False
        '
        '_mnuFileItem_100
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_100, CType(100, Short))
        Me._mnuFileItem_100.Name = "_mnuFileItem_100"
        Me._mnuFileItem_100.Size = New System.Drawing.Size(154, 22)
        Me._mnuFileItem_100.Text = "&Close"
        Me._mnuFileItem_100.Visible = False
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuEditItem_10, Me._mnuEditItem_20})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(39, 20)
        Me.mnuEdit.Text = "&Edit"
        '
        '_mnuEditItem_10
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_10, CType(10, Short))
        Me._mnuEditItem_10.Name = "_mnuEditItem_10"
        Me._mnuEditItem_10.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me._mnuEditItem_10.Size = New System.Drawing.Size(144, 22)
        Me._mnuEditItem_10.Text = "&Copy"
        '
        '_mnuEditItem_20
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_20, CType(20, Short))
        Me._mnuEditItem_20.Name = "_mnuEditItem_20"
        Me._mnuEditItem_20.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me._mnuEditItem_20.Size = New System.Drawing.Size(144, 22)
        Me._mnuEditItem_20.Text = "&Paste"
        '
        '_cboUnits_1
        '
        Me._cboUnits_1.BackColor = System.Drawing.SystemColors.Window
        Me._cboUnits_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboUnits_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboUnits_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboUnits_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboUnits.SetIndex(Me._cboUnits_1, CType(1, Short))
        Me._cboUnits_1.Location = New System.Drawing.Point(250, 56)
        Me._cboUnits_1.Name = "_cboUnits_1"
        Me._cboUnits_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboUnits_1.Size = New System.Drawing.Size(103, 24)
        Me._cboUnits_1.TabIndex = 3
        Me._cboUnits_1.TabStop = False
        '
        '_cboUnits_0
        '
        Me._cboUnits_0.BackColor = System.Drawing.SystemColors.Window
        Me._cboUnits_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboUnits_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboUnits_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboUnits_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboUnits.SetIndex(Me._cboUnits_0, CType(0, Short))
        Me._cboUnits_0.Location = New System.Drawing.Point(250, 28)
        Me._cboUnits_0.Name = "_cboUnits_0"
        Me._cboUnits_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboUnits_0.Size = New System.Drawing.Size(103, 24)
        Me._cboUnits_0.TabIndex = 1
        Me._cboUnits_0.TabStop = False
        '
        '_lblData_1
        '
        Me._lblData_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblData_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblData.SetIndex(Me._lblData_1, CType(1, Short))
        Me._lblData_1.Location = New System.Drawing.Point(98, 60)
        Me._lblData_1.Name = "_lblData_1"
        Me._lblData_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_1.Size = New System.Drawing.Size(147, 19)
        Me._lblData_1.TabIndex = 4
        Me._lblData_1.Text = "{Whatever} Units:"
        Me._lblData_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblData_0
        '
        Me._lblData_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblData_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblData.SetIndex(Me._lblData_0, CType(0, Short))
        Me._lblData_0.Location = New System.Drawing.Point(98, 32)
        Me._lblData_0.Name = "_lblData_0"
        Me._lblData_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_0.Size = New System.Drawing.Size(147, 19)
        Me._lblData_0.TabIndex = 2
        Me._lblData_0.Text = "Time Units:"
        Me._lblData_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cboUnits
        '
        '
        'mnuEditItem
        '
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(12, 32)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(72, 49)
        Me._cmdCancelOK_0.TabIndex = 12
        Me._cmdCancelOK_0.Text = "OK"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(12, 98)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(72, 49)
        Me._cmdCancelOK_1.TabIndex = 13
        Me._cmdCancelOK_1.Text = "Cancel"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        'frmTimeVarGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(496, 329)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Controls.Add(Me._cboUnits_1)
        Me.Controls.Add(Me._cboUnits_0)
        Me.Controls.Add(Me._lblData_1)
        Me.Controls.Add(Me._lblData_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(64, 132)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(512, 368)
        Me.Name = "frmTimeVarGrid"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "{Caption}"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.cboUnits, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuEditItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFileItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents _cmdCancelOK_1 As Button
#End Region
End Class