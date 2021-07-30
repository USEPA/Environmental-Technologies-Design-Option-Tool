<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditIsotherm
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
	Public WithEvents _mnuChemicalItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuChemicalItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuChemicalItem_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuChemical As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuIsothermItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuIsothermItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuIsothermItem_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuIsothermItem_4 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuIsotherm As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
    Public WithEvents lstCompo As System.Windows.Forms.ListBox
    Public WithEvents lblEmpty_lstCompo As System.Windows.Forms.Label
    Public WithEvents _lstRange_1 As System.Windows.Forms.ListBox
    Public WithEvents _lstRange_0 As System.Windows.Forms.ListBox
    Public WithEvents lblPhase As System.Windows.Forms.Label
    Public WithEvents _Label5_0 As System.Windows.Forms.Label
    Public WithEvents _lblValue_3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents _lblValue_2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblText_1 As System.Windows.Forms.Label
    Public WithEvents _lblValue_1 As System.Windows.Forms.Label
    Public WithEvents _lblValue_0 As System.Windows.Forms.Label
    Public WithEvents _lblText_0 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents lblTemperature As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents lblComments As System.Windows.Forms.Label
    Public WithEvents Label5 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    '  Public WithEvents cmdFind As SSCommandArray
    Public WithEvents lblText As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblValue As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lstRange As Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray
    Public WithEvents mnuChemicalItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuIsothermItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    '  Public WithEvents optSort As SSOptionArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuChemical = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuChemicalItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuChemicalItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuChemicalItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuIsotherm = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuIsothermItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuIsothermItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuIsothermItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuIsothermItem_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Label5 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._Label5_0 = New System.Windows.Forms.Label()
        Me.lblText = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblText_1 = New System.Windows.Forms.Label()
        Me._lblText_0 = New System.Windows.Forms.Label()
        Me.lblValue = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblValue_3 = New System.Windows.Forms.Label()
        Me._lblValue_2 = New System.Windows.Forms.Label()
        Me._lblValue_1 = New System.Windows.Forms.Label()
        Me._lblValue_0 = New System.Windows.Forms.Label()
        Me.lstRange = New Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray(Me.components)
        Me._lstRange_1 = New System.Windows.Forms.ListBox()
        Me._lstRange_0 = New System.Windows.Forms.ListBox()
        Me.mnuChemicalItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuIsothermItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.lstCompo = New System.Windows.Forms.ListBox()
        Me.lblEmpty_lstCompo = New System.Windows.Forms.Label()
        Me.lblPhase = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblTemperature = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me._optSort_1 = New System.Windows.Forms.RadioButton()
        Me._cmdFind_1 = New System.Windows.Forms.Button()
        Me._cmdFind_0 = New System.Windows.Forms.Button()
        Me._optSort_0 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.MainMenu1.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblText, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lstRange, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuChemicalItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuIsothermItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuChemical, Me.mnuIsotherm, Me.mnuExit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(731, 24)
        Me.MainMenu1.TabIndex = 30
        '
        'mnuChemical
        '
        Me.mnuChemical.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuChemicalItem_1, Me._mnuChemicalItem_2, Me._mnuChemicalItem_3})
        Me.mnuChemical.Name = "mnuChemical"
        Me.mnuChemical.Size = New System.Drawing.Size(69, 20)
        Me.mnuChemical.Text = "C&hemical"
        '
        '_mnuChemicalItem_1
        '
        Me.mnuChemicalItem.SetIndex(Me._mnuChemicalItem_1, CType(1, Short))
        Me._mnuChemicalItem_1.Name = "_mnuChemicalItem_1"
        Me._mnuChemicalItem_1.Size = New System.Drawing.Size(150, 22)
        Me._mnuChemicalItem_1.Text = "&New"
        '
        '_mnuChemicalItem_2
        '
        Me.mnuChemicalItem.SetIndex(Me._mnuChemicalItem_2, CType(2, Short))
        Me._mnuChemicalItem_2.Name = "_mnuChemicalItem_2"
        Me._mnuChemicalItem_2.Size = New System.Drawing.Size(150, 22)
        Me._mnuChemicalItem_2.Text = "&Edit Current"
        '
        '_mnuChemicalItem_3
        '
        Me.mnuChemicalItem.SetIndex(Me._mnuChemicalItem_3, CType(3, Short))
        Me._mnuChemicalItem_3.Name = "_mnuChemicalItem_3"
        Me._mnuChemicalItem_3.Size = New System.Drawing.Size(150, 22)
        Me._mnuChemicalItem_3.Text = "&Delete Current"
        '
        'mnuIsotherm
        '
        Me.mnuIsotherm.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuIsothermItem_1, Me._mnuIsothermItem_2, Me._mnuIsothermItem_3, Me._mnuIsothermItem_4})
        Me.mnuIsotherm.Name = "mnuIsotherm"
        Me.mnuIsotherm.Size = New System.Drawing.Size(66, 20)
        Me.mnuIsotherm.Text = "&Isotherm"
        '
        '_mnuIsothermItem_1
        '
        Me.mnuIsothermItem.SetIndex(Me._mnuIsothermItem_1, CType(1, Short))
        Me._mnuIsothermItem_1.Name = "_mnuIsothermItem_1"
        Me._mnuIsothermItem_1.Size = New System.Drawing.Size(150, 22)
        Me._mnuIsothermItem_1.Text = "&New"
        '
        '_mnuIsothermItem_2
        '
        Me.mnuIsothermItem.SetIndex(Me._mnuIsothermItem_2, CType(2, Short))
        Me._mnuIsothermItem_2.Name = "_mnuIsothermItem_2"
        Me._mnuIsothermItem_2.Size = New System.Drawing.Size(150, 22)
        Me._mnuIsothermItem_2.Text = "&Edit Current"
        '
        '_mnuIsothermItem_3
        '
        Me.mnuIsothermItem.SetIndex(Me._mnuIsothermItem_3, CType(3, Short))
        Me._mnuIsothermItem_3.Name = "_mnuIsothermItem_3"
        Me._mnuIsothermItem_3.Size = New System.Drawing.Size(150, 22)
        Me._mnuIsothermItem_3.Text = "&Delete Current"
        '
        '_mnuIsothermItem_4
        '
        Me.mnuIsothermItem.SetIndex(Me._mnuIsothermItem_4, CType(4, Short))
        Me._mnuIsothermItem_4.Name = "_mnuIsothermItem_4"
        Me._mnuIsothermItem_4.Size = New System.Drawing.Size(150, 22)
        Me._mnuIsothermItem_4.Text = "Delete &All"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(38, 20)
        Me.mnuExit.Text = "E&xit"
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(699, 430)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 29
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        '_Label5_0
        '
        Me._Label5_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._Label5_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label5_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label5_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.SetIndex(Me._Label5_0, CType(0, Short))
        Me._Label5_0.Location = New System.Drawing.Point(181, 22)
        Me._Label5_0.Name = "_Label5_0"
        Me._Label5_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label5_0.Size = New System.Drawing.Size(156, 17)
        Me._Label5_0.TabIndex = 22
        Me._Label5_0.Text = "(mg/g)*(L/mg)^(1/n)"
        '
        '_lblText_1
        '
        Me._lblText_1.BackColor = System.Drawing.Color.Transparent
        Me._lblText_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_1, CType(1, Short))
        Me._lblText_1.Location = New System.Drawing.Point(24, 42)
        Me._lblText_1.Name = "_lblText_1"
        Me._lblText_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_1.Size = New System.Drawing.Size(65, 17)
        Me._lblText_1.TabIndex = 15
        Me._lblText_1.Text = "1/n"
        Me._lblText_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblText_0
        '
        Me._lblText_0.BackColor = System.Drawing.Color.Transparent
        Me._lblText_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_0, CType(0, Short))
        Me._lblText_0.Location = New System.Drawing.Point(24, 22)
        Me._lblText_0.Name = "_lblText_0"
        Me._lblText_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_0.Size = New System.Drawing.Size(65, 17)
        Me._lblText_0.TabIndex = 12
        Me._lblText_0.Text = "K"
        Me._lblText_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_3
        '
        Me._lblValue_3.BackColor = System.Drawing.Color.Transparent
        Me._lblValue_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_3, CType(3, Short))
        Me._lblValue_3.Location = New System.Drawing.Point(8, 347)
        Me._lblValue_3.Name = "_lblValue_3"
        Me._lblValue_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_3.Size = New System.Drawing.Size(342, 24)
        Me._lblValue_3.TabIndex = 21
        Me._lblValue_3.Text = "lblValue(3)"
        Me._lblValue_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_2
        '
        Me._lblValue_2.BackColor = System.Drawing.Color.Transparent
        Me._lblValue_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_2, CType(2, Short))
        Me._lblValue_2.Location = New System.Drawing.Point(145, 263)
        Me._lblValue_2.Name = "_lblValue_2"
        Me._lblValue_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_2.Size = New System.Drawing.Size(205, 17)
        Me._lblValue_2.TabIndex = 19
        Me._lblValue_2.Text = "lblValue(2)"
        Me._lblValue_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_1
        '
        Me._lblValue_1.BackColor = System.Drawing.Color.Transparent
        Me._lblValue_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_1, CType(1, Short))
        Me._lblValue_1.Location = New System.Drawing.Point(92, 41)
        Me._lblValue_1.Name = "_lblValue_1"
        Me._lblValue_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_1.Size = New System.Drawing.Size(83, 18)
        Me._lblValue_1.TabIndex = 14
        Me._lblValue_1.Text = "lblValue(1)"
        Me._lblValue_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_0
        '
        Me._lblValue_0.BackColor = System.Drawing.Color.Transparent
        Me._lblValue_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_0, CType(0, Short))
        Me._lblValue_0.Location = New System.Drawing.Point(92, 21)
        Me._lblValue_0.Name = "_lblValue_0"
        Me._lblValue_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_0.Size = New System.Drawing.Size(83, 20)
        Me._lblValue_0.TabIndex = 13
        Me._lblValue_0.Text = "lblValue(0)"
        Me._lblValue_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lstRange
        '
        '
        '_lstRange_1
        '
        Me._lstRange_1.BackColor = System.Drawing.SystemColors.Window
        Me._lstRange_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lstRange_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstRange_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lstRange_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRange.SetIndex(Me._lstRange_1, CType(1, Short))
        Me._lstRange_1.ItemHeight = 14
        Me._lstRange_1.Location = New System.Drawing.Point(139, 82)
        Me._lstRange_1.Name = "_lstRange_1"
        Me._lstRange_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRange_1.Size = New System.Drawing.Size(213, 156)
        Me._lstRange_1.TabIndex = 6
        '
        '_lstRange_0
        '
        Me._lstRange_0.BackColor = System.Drawing.SystemColors.Window
        Me._lstRange_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lstRange_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstRange_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lstRange_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRange.SetIndex(Me._lstRange_0, CType(0, Short))
        Me._lstRange_0.ItemHeight = 14
        Me._lstRange_0.Location = New System.Drawing.Point(8, 82)
        Me._lstRange_0.Name = "_lstRange_0"
        Me._lstRange_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRange_0.Size = New System.Drawing.Size(125, 156)
        Me._lstRange_0.TabIndex = 5
        '
        'mnuChemicalItem
        '
        '
        'mnuIsothermItem
        '
        '
        'lstCompo
        '
        Me.lstCompo.BackColor = System.Drawing.SystemColors.Window
        Me.lstCompo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCompo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCompo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCompo.ItemHeight = 14
        Me.lstCompo.Location = New System.Drawing.Point(7, 76)
        Me.lstCompo.Name = "lstCompo"
        Me.lstCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCompo.Size = New System.Drawing.Size(270, 254)
        Me.lstCompo.TabIndex = 4
        '
        'lblEmpty_lstCompo
        '
        Me.lblEmpty_lstCompo.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmpty_lstCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmpty_lstCompo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpty_lstCompo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmpty_lstCompo.Location = New System.Drawing.Point(21, 50)
        Me.lblEmpty_lstCompo.Name = "lblEmpty_lstCompo"
        Me.lblEmpty_lstCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmpty_lstCompo.Size = New System.Drawing.Size(241, 19)
        Me.lblEmpty_lstCompo.TabIndex = 24
        Me.lblEmpty_lstCompo.Text = "No Chemicals Available"
        Me.lblEmpty_lstCompo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEmpty_lstCompo.Visible = False
        '
        'lblPhase
        '
        Me.lblPhase.BackColor = System.Drawing.Color.Transparent
        Me.lblPhase.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPhase.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPhase.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhase.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblPhase.Location = New System.Drawing.Point(145, 303)
        Me.lblPhase.Name = "lblPhase"
        Me.lblPhase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPhase.Size = New System.Drawing.Size(205, 17)
        Me.lblPhase.TabIndex = 23
        Me.lblPhase.Text = "lblPhase"
        Me.lblPhase.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(8, 330)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(81, 17)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Source:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.DarkGray
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(12, 264)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(121, 20)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Carbon Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.DarkGray
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(8, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(89, 17)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "pH Range:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(145, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(205, 17)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Concentration Range (mg/L):"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.DarkGray
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(28, 304)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(105, 26)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Phase:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTemperature
        '
        Me.lblTemperature.BackColor = System.Drawing.Color.Transparent
        Me.lblTemperature.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTemperature.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTemperature.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemperature.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTemperature.Location = New System.Drawing.Point(145, 283)
        Me.lblTemperature.Name = "lblTemperature"
        Me.lblTemperature.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTemperature.Size = New System.Drawing.Size(205, 17)
        Me.lblTemperature.TabIndex = 10
        Me.lblTemperature.Text = "lblTemperature"
        Me.lblTemperature.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DarkGray
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(6, 284)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(127, 20)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Temperature (C):"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label8.Location = New System.Drawing.Point(8, 376)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(89, 16)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Comments:"
        '
        'lblComments
        '
        Me.lblComments.BackColor = System.Drawing.Color.Transparent
        Me.lblComments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblComments.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblComments.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblComments.Location = New System.Drawing.Point(7, 395)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblComments.Size = New System.Drawing.Size(342, 28)
        Me.lblComments.TabIndex = 7
        Me.lblComments.Text = "lblComments"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdSelect)
        Me.GroupBox1.Controls.Add(Me._optSort_1)
        Me.GroupBox1.Controls.Add(Me._cmdFind_1)
        Me.GroupBox1.Controls.Add(Me._cmdFind_0)
        Me.GroupBox1.Controls.Add(Me._optSort_0)
        Me.GroupBox1.Controls.Add(Me.lblEmpty_lstCompo)
        Me.GroupBox1.Controls.Add(Me.lstCompo)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 44)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(289, 443)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "fraOne"
        '
        'cmdSelect
        '
        Me.cmdSelect.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdSelect.Location = New System.Drawing.Point(12, 383)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(249, 32)
        Me.cmdSelect.TabIndex = 36
        Me.cmdSelect.Text = "Select Chemic&al"
        Me.cmdSelect.UseVisualStyleBackColor = False
        '
        '_optSort_1
        '
        Me._optSort_1.AutoSize = True
        Me._optSort_1.Location = New System.Drawing.Point(123, 19)
        Me._optSort_1.Name = "_optSort_1"
        Me._optSort_1.Size = New System.Drawing.Size(125, 18)
        Me._optSort_1.TabIndex = 33
        Me._optSort_1.Text = "Sort by CAS Number"
        Me._optSort_1.UseVisualStyleBackColor = True
        '
        '_cmdFind_1
        '
        Me._cmdFind_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdFind_1.Location = New System.Drawing.Point(150, 347)
        Me._cmdFind_1.Name = "_cmdFind_1"
        Me._cmdFind_1.Size = New System.Drawing.Size(98, 30)
        Me._cmdFind_1.TabIndex = 35
        Me._cmdFind_1.Text = "Find &Again"
        Me._cmdFind_1.UseVisualStyleBackColor = False
        '
        '_cmdFind_0
        '
        Me._cmdFind_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdFind_0.Location = New System.Drawing.Point(24, 347)
        Me._cmdFind_0.Name = "_cmdFind_0"
        Me._cmdFind_0.Size = New System.Drawing.Size(103, 30)
        Me._cmdFind_0.TabIndex = 34
        Me._cmdFind_0.Text = "&Find"
        Me._cmdFind_0.UseVisualStyleBackColor = False
        '
        '_optSort_0
        '
        Me._optSort_0.AutoSize = True
        Me._optSort_0.Checked = True
        Me._optSort_0.Location = New System.Drawing.Point(5, 19)
        Me._optSort_0.Name = "_optSort_0"
        Me._optSort_0.Size = New System.Drawing.Size(90, 18)
        Me._optSort_0.TabIndex = 32
        Me._optSort_0.TabStop = True
        Me._optSort_0.Text = "Sort by Name"
        Me._optSort_0.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me._lblText_0)
        Me.GroupBox2.Controls.Add(Me._lblText_1)
        Me.GroupBox2.Controls.Add(Me._lblValue_0)
        Me.GroupBox2.Controls.Add(Me._lblValue_1)
        Me.GroupBox2.Controls.Add(Me._Label5_0)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me._lstRange_0)
        Me.GroupBox2.Controls.Add(Me._lstRange_1)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me._lblValue_2)
        Me.GroupBox2.Controls.Add(Me.lblTemperature)
        Me.GroupBox2.Controls.Add(Me.lblPhase)
        Me.GroupBox2.Controls.Add(Me._lblValue_3)
        Me.GroupBox2.Controls.Add(Me.lblComments)
        Me.GroupBox2.Location = New System.Drawing.Point(307, 44)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(366, 443)
        Me.GroupBox2.TabIndex = 33
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "fraTwo"
        '
        'frmEditIsotherm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(731, 576)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(129, 129)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(747, 615)
        Me.Name = "frmEditIsotherm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Isotherm Database"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblText, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lstRange, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuChemicalItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuIsothermItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents _optSort_1 As RadioButton
    Friend WithEvents _optSort_0 As RadioButton
    Friend WithEvents _cmdFind_0 As Button
    Friend WithEvents _cmdFind_1 As Button
    Friend WithEvents cmdSelect As Button
#End Region
End Class