<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmVarConcentrations
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
	Public WithEvents _mnuFileItem_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_190 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _mnuFileItem_191 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_192 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_193 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuFileItem_194 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuEditItem_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuEditItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuEditItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox

    '    Public WithEvents Sheet1 As VCIF1Lib.F1Book
    Public WithEvents Sheet1DataGrid As DataGridView

    Public WithEvents excelsheet1 As Microsoft.Office.Interop.Excel.Worksheet
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_8 As System.Windows.Forms.Label
    Public WithEvents _Label1_9 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_2 As System.Windows.Forms.Label
    Public WithEvents _Label2_3 As System.Windows.Forms.Label
    Public WithEvents _Label2_4 As System.Windows.Forms.Label
    Public WithEvents _Label2_5 As System.Windows.Forms.Label
    Public WithEvents _Label2_6 As System.Windows.Forms.Label
    Public WithEvents _Label2_7 As System.Windows.Forms.Label
    Public WithEvents _Label2_8 As System.Windows.Forms.Label
    Public WithEvents _Label2_9 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
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
        Me._mnuFileItem_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_190 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFileItem_191 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_192 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_193 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_194 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Sheet1DataGrid = New System.Windows.Forms.DataGridView()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_8 = New System.Windows.Forms.Label()
        Me._Label1_9 = New System.Windows.Forms.Label()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_2 = New System.Windows.Forms.Label()
        Me._Label2_3 = New System.Windows.Forms.Label()
        Me._Label2_4 = New System.Windows.Forms.Label()
        Me._Label2_5 = New System.Windows.Forms.Label()
        Me._Label2_6 = New System.Windows.Forms.Label()
        Me._Label2_7 = New System.Windows.Forms.Label()
        Me._Label2_8 = New System.Windows.Forms.Label()
        Me._Label2_9 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.mnuEditItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuFileItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.MainMenu1.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Sheet1DataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuEditItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFileItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(537, 24)
        Me.MainMenu1.TabIndex = 27
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFileItem_0, Me._mnuFileItem_1, Me._mnuFileItem_2, Me._mnuFileItem_3, Me._mnuFileItem_190, Me._mnuFileItem_191, Me._mnuFileItem_192, Me._mnuFileItem_193, Me._mnuFileItem_194})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        '_mnuFileItem_0
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_0, CType(0, Short))
        Me._mnuFileItem_0.Name = "_mnuFileItem_0"
        Me._mnuFileItem_0.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_0.Text = "&New"
        '
        '_mnuFileItem_1
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_1, CType(1, Short))
        Me._mnuFileItem_1.Name = "_mnuFileItem_1"
        Me._mnuFileItem_1.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_1.Text = "&Open ..."
        '
        '_mnuFileItem_2
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_2, CType(2, Short))
        Me._mnuFileItem_2.Name = "_mnuFileItem_2"
        Me._mnuFileItem_2.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_2.Text = "&Save"
        '
        '_mnuFileItem_3
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_3, CType(3, Short))
        Me._mnuFileItem_3.Name = "_mnuFileItem_3"
        Me._mnuFileItem_3.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_3.Text = "Save &As ..."
        '
        '_mnuFileItem_190
        '
        Me._mnuFileItem_190.Name = "_mnuFileItem_190"
        Me._mnuFileItem_190.Size = New System.Drawing.Size(136, 6)
        '
        '_mnuFileItem_191
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_191, CType(191, Short))
        Me._mnuFileItem_191.Name = "_mnuFileItem_191"
        Me._mnuFileItem_191.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_191.Text = "&1 Old File #1"
        '
        '_mnuFileItem_192
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_192, CType(192, Short))
        Me._mnuFileItem_192.Name = "_mnuFileItem_192"
        Me._mnuFileItem_192.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_192.Text = "&2 Old File #2"
        '
        '_mnuFileItem_193
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_193, CType(193, Short))
        Me._mnuFileItem_193.Name = "_mnuFileItem_193"
        Me._mnuFileItem_193.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_193.Text = "&3 Old File #3"
        '
        '_mnuFileItem_194
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_194, CType(194, Short))
        Me._mnuFileItem_194.Name = "_mnuFileItem_194"
        Me._mnuFileItem_194.Size = New System.Drawing.Size(139, 22)
        Me._mnuFileItem_194.Text = "&4 Old File #4"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuEditItem_0, Me._mnuEditItem_1, Me._mnuEditItem_2})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(39, 20)
        Me.mnuEdit.Text = "&Edit"
        '
        '_mnuEditItem_0
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_0, CType(0, Short))
        Me._mnuEditItem_0.Name = "_mnuEditItem_0"
        Me._mnuEditItem_0.Size = New System.Drawing.Size(102, 22)
        Me._mnuEditItem_0.Text = "Cu&t"
        '
        '_mnuEditItem_1
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_1, CType(1, Short))
        Me._mnuEditItem_1.Name = "_mnuEditItem_1"
        Me._mnuEditItem_1.Size = New System.Drawing.Size(102, 22)
        Me._mnuEditItem_1.Text = "&Copy"
        '
        '_mnuEditItem_2
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_2, CType(2, Short))
        Me._mnuEditItem_2.Name = "_mnuEditItem_2"
        Me._mnuEditItem_2.Size = New System.Drawing.Size(102, 22)
        Me._mnuEditItem_2.Text = "&Paste"
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(504, 0)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 25
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'Sheet1DataGrid
        '
        Me.Sheet1DataGrid.ColumnHeadersHeight = 29
        Me.Sheet1DataGrid.Location = New System.Drawing.Point(9, 173)
        Me.Sheet1DataGrid.Name = "Sheet1DataGrid"
        Me.Sheet1DataGrid.RowHeadersWidth = 51
        Me.Sheet1DataGrid.Size = New System.Drawing.Size(324, 287)
        Me.Sheet1DataGrid.TabIndex = 26
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(6, 65)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(37, 17)
        Me._Label1_0.TabIndex = 21
        Me._Label1_0.Text = "A="
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_0.Visible = False
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_0, CType(0, Short))
        Me._Label2_0.Location = New System.Drawing.Point(46, 66)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(217, 17)
        Me._Label2_0.TabIndex = 20
        Me._Label2_0.Text = "Label2"
        Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_0.Visible = False
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(6, 81)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(37, 17)
        Me._Label1_1.TabIndex = 19
        Me._Label1_1.Text = "B="
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_1.Visible = False
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(6, 97)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(37, 17)
        Me._Label1_2.TabIndex = 18
        Me._Label1_2.Text = "C="
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_2.Visible = False
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(6, 113)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(37, 17)
        Me._Label1_3.TabIndex = 17
        Me._Label1_3.Text = "D="
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_3.Visible = False
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_4, CType(4, Short))
        Me._Label1_4.Location = New System.Drawing.Point(6, 129)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(37, 17)
        Me._Label1_4.TabIndex = 16
        Me._Label1_4.Text = "E="
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_4.Visible = False
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_5, CType(5, Short))
        Me._Label1_5.Location = New System.Drawing.Point(270, 65)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(37, 17)
        Me._Label1_5.TabIndex = 15
        Me._Label1_5.Text = "F="
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_5.Visible = False
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_6, CType(6, Short))
        Me._Label1_6.Location = New System.Drawing.Point(270, 81)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(37, 17)
        Me._Label1_6.TabIndex = 14
        Me._Label1_6.Text = "G="
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_6.Visible = False
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_7, CType(7, Short))
        Me._Label1_7.Location = New System.Drawing.Point(270, 97)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(37, 17)
        Me._Label1_7.TabIndex = 13
        Me._Label1_7.Text = "H="
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_7.Visible = False
        '
        '_Label1_8
        '
        Me._Label1_8.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_8, CType(8, Short))
        Me._Label1_8.Location = New System.Drawing.Point(270, 113)
        Me._Label1_8.Name = "_Label1_8"
        Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_8.Size = New System.Drawing.Size(37, 17)
        Me._Label1_8.TabIndex = 12
        Me._Label1_8.Text = "I="
        Me._Label1_8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_8.Visible = False
        '
        '_Label1_9
        '
        Me._Label1_9.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.SetIndex(Me._Label1_9, CType(9, Short))
        Me._Label1_9.Location = New System.Drawing.Point(270, 129)
        Me._Label1_9.Name = "_Label1_9"
        Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_9.Size = New System.Drawing.Size(37, 17)
        Me._Label1_9.TabIndex = 11
        Me._Label1_9.Text = "J="
        Me._Label1_9.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_9.Visible = False
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_1, CType(1, Short))
        Me._Label2_1.Location = New System.Drawing.Point(46, 82)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(217, 17)
        Me._Label2_1.TabIndex = 10
        Me._Label2_1.Text = "Label2"
        Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_1.Visible = False
        '
        '_Label2_2
        '
        Me._Label2_2.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_2, CType(2, Short))
        Me._Label2_2.Location = New System.Drawing.Point(46, 98)
        Me._Label2_2.Name = "_Label2_2"
        Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_2.Size = New System.Drawing.Size(217, 17)
        Me._Label2_2.TabIndex = 9
        Me._Label2_2.Text = "Label2"
        Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_2.Visible = False
        '
        '_Label2_3
        '
        Me._Label2_3.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_3, CType(3, Short))
        Me._Label2_3.Location = New System.Drawing.Point(46, 114)
        Me._Label2_3.Name = "_Label2_3"
        Me._Label2_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_3.Size = New System.Drawing.Size(217, 17)
        Me._Label2_3.TabIndex = 8
        Me._Label2_3.Text = "Label2"
        Me._Label2_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_3.Visible = False
        '
        '_Label2_4
        '
        Me._Label2_4.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_4, CType(4, Short))
        Me._Label2_4.Location = New System.Drawing.Point(46, 130)
        Me._Label2_4.Name = "_Label2_4"
        Me._Label2_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_4.Size = New System.Drawing.Size(217, 17)
        Me._Label2_4.TabIndex = 7
        Me._Label2_4.Text = "Label2"
        Me._Label2_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_4.Visible = False
        '
        '_Label2_5
        '
        Me._Label2_5.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_5, CType(5, Short))
        Me._Label2_5.Location = New System.Drawing.Point(310, 66)
        Me._Label2_5.Name = "_Label2_5"
        Me._Label2_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_5.Size = New System.Drawing.Size(217, 17)
        Me._Label2_5.TabIndex = 6
        Me._Label2_5.Text = "Label2"
        Me._Label2_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_5.Visible = False
        '
        '_Label2_6
        '
        Me._Label2_6.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_6, CType(6, Short))
        Me._Label2_6.Location = New System.Drawing.Point(310, 82)
        Me._Label2_6.Name = "_Label2_6"
        Me._Label2_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_6.Size = New System.Drawing.Size(217, 17)
        Me._Label2_6.TabIndex = 5
        Me._Label2_6.Text = "Label2"
        Me._Label2_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_6.Visible = False
        '
        '_Label2_7
        '
        Me._Label2_7.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_7, CType(7, Short))
        Me._Label2_7.Location = New System.Drawing.Point(310, 98)
        Me._Label2_7.Name = "_Label2_7"
        Me._Label2_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_7.Size = New System.Drawing.Size(217, 17)
        Me._Label2_7.TabIndex = 4
        Me._Label2_7.Text = "Label2"
        Me._Label2_7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_7.Visible = False
        '
        '_Label2_8
        '
        Me._Label2_8.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_8, CType(8, Short))
        Me._Label2_8.Location = New System.Drawing.Point(310, 114)
        Me._Label2_8.Name = "_Label2_8"
        Me._Label2_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_8.Size = New System.Drawing.Size(217, 17)
        Me._Label2_8.TabIndex = 3
        Me._Label2_8.Text = "Label2"
        Me._Label2_8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_8.Visible = False
        '
        '_Label2_9
        '
        Me._Label2_9.BackColor = System.Drawing.SystemColors.Window
        Me._Label2_9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._Label2_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.SetIndex(Me._Label2_9, CType(9, Short))
        Me._Label2_9.Location = New System.Drawing.Point(310, 130)
        Me._Label2_9.Name = "_Label2_9"
        Me._Label2_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_9.Size = New System.Drawing.Size(217, 17)
        Me._Label2_9.TabIndex = 2
        Me._Label2_9.Text = "Label2"
        Me._Label2_9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me._Label2_9.Visible = False
        '
        'mnuEditItem
        '
        '
        'mnuFileItem
        '
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdCancel.Location = New System.Drawing.Point(24, 27)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(87, 32)
        Me.cmdCancel.TabIndex = 28
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdOK.Location = New System.Drawing.Point(114, 27)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(86, 32)
        Me.cmdOK.TabIndex = 29
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmVarConcentrations
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(537, 472)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Sheet1DataGrid)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_2)
        Me.Controls.Add(Me._Label1_3)
        Me.Controls.Add(Me._Label1_4)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_6)
        Me.Controls.Add(Me._Label1_7)
        Me.Controls.Add(Me._Label1_8)
        Me.Controls.Add(Me._Label1_9)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me._Label2_2)
        Me.Controls.Add(Me._Label2_3)
        Me.Controls.Add(Me._Label2_4)
        Me.Controls.Add(Me._Label2_5)
        Me.Controls.Add(Me._Label2_6)
        Me.Controls.Add(Me._Label2_7)
        Me.Controls.Add(Me._Label2_8)
        Me.Controls.Add(Me._Label2_9)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(129, 208)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(553, 511)
        Me.Name = "frmVarConcentrations"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Influent/Effluent Concentrations"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Sheet1DataGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuEditItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFileItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
#End Region
End Class