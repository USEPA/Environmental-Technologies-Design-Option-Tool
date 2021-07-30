<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditCarbon
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
	Public WithEvents _mnuManufacturerItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuManufacturerItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuManufacturerItem_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuManufacturer As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuAdsorbentItem_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuAdsorbentItem_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuAdsorbentItem_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAdsorbent As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
    Public WithEvents lstManu As System.Windows.Forms.ListBox
    Public WithEvents lblEmpty_lstManu As System.Windows.Forms.Label
    Public WithEvents _lblDesc_0 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_1 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_2 As System.Windows.Forms.Label
    Public WithEvents _lblData_0 As System.Windows.Forms.Label
    Public WithEvents _lblData_1 As System.Windows.Forms.Label
    Public WithEvents _lblData_2 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_0 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_1 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_2 As System.Windows.Forms.Label
    Public WithEvents _lblData_3 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_3 As System.Windows.Forms.Label
    Public WithEvents _lblData_4 As System.Windows.Forms.Label
    Public WithEvents _lblData_5 As System.Windows.Forms.Label
    Public WithEvents _lblData_6 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_4 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_5 As System.Windows.Forms.Label
    Public WithEvents _lblDesc_6 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_4 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_5 As System.Windows.Forms.Label
    Public WithEvents _lblUnit_6 As System.Windows.Forms.Label
    Public WithEvents lblUnitB As System.Windows.Forms.Label
    Public WithEvents lstName As System.Windows.Forms.ListBox
    Public WithEvents lblEmpty_lstName As System.Windows.Forms.Label
    Public WithEvents lblData As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblDesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents mnuAdsorbentItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuManufacturerItem As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    '  Public WithEvents optPhase As SSOptionArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuManufacturer = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuManufacturerItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuManufacturerItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuManufacturerItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAdsorbent = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuAdsorbentItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuAdsorbentItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuAdsorbentItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.lblData = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblData_0 = New System.Windows.Forms.Label()
        Me._lblData_3 = New System.Windows.Forms.Label()
        Me._lblData_4 = New System.Windows.Forms.Label()
        Me._lblData_5 = New System.Windows.Forms.Label()
        Me._lblData_6 = New System.Windows.Forms.Label()
        Me._lblData_1 = New System.Windows.Forms.Label()
        Me._lblData_2 = New System.Windows.Forms.Label()
        Me._lblDesc_1 = New System.Windows.Forms.Label()
        Me._lblDesc_2 = New System.Windows.Forms.Label()
        Me.lblDesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblDesc_0 = New System.Windows.Forms.Label()
        Me._lblDesc_3 = New System.Windows.Forms.Label()
        Me._lblDesc_4 = New System.Windows.Forms.Label()
        Me._lblDesc_5 = New System.Windows.Forms.Label()
        Me._lblDesc_6 = New System.Windows.Forms.Label()
        Me.lblUnit = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblUnit_0 = New System.Windows.Forms.Label()
        Me._lblUnit_1 = New System.Windows.Forms.Label()
        Me._lblUnit_2 = New System.Windows.Forms.Label()
        Me._lblUnit_4 = New System.Windows.Forms.Label()
        Me._lblUnit_5 = New System.Windows.Forms.Label()
        Me._lblUnit_6 = New System.Windows.Forms.Label()
        Me.mnuAdsorbentItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuManufacturerItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.lstManu = New System.Windows.Forms.ListBox()
        Me.lblEmpty_lstManu = New System.Windows.Forms.Label()
        Me.lblUnitB = New System.Windows.Forms.Label()
        Me.lstName = New System.Windows.Forms.ListBox()
        Me.lblEmpty_lstName = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me._optPhase_1 = New System.Windows.Forms.RadioButton()
        Me._optPhase_0 = New System.Windows.Forms.RadioButton()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.MainMenu1.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuAdsorbentItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuManufacturerItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuManufacturer, Me.mnuAdsorbent})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(666, 24)
        Me.MainMenu1.TabIndex = 34
        '
        'mnuManufacturer
        '
        Me.mnuManufacturer.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuManufacturerItem_1, Me._mnuManufacturerItem_2, Me._mnuManufacturerItem_3})
        Me.mnuManufacturer.Name = "mnuManufacturer"
        Me.mnuManufacturer.Size = New System.Drawing.Size(91, 20)
        Me.mnuManufacturer.Text = "&Manufacturer"
        '
        '_mnuManufacturerItem_1
        '
        Me.mnuManufacturerItem.SetIndex(Me._mnuManufacturerItem_1, CType(1, Short))
        Me._mnuManufacturerItem_1.Name = "_mnuManufacturerItem_1"
        Me._mnuManufacturerItem_1.Size = New System.Drawing.Size(150, 22)
        Me._mnuManufacturerItem_1.Text = "&New"
        '
        '_mnuManufacturerItem_2
        '
        Me.mnuManufacturerItem.SetIndex(Me._mnuManufacturerItem_2, CType(2, Short))
        Me._mnuManufacturerItem_2.Name = "_mnuManufacturerItem_2"
        Me._mnuManufacturerItem_2.Size = New System.Drawing.Size(150, 22)
        Me._mnuManufacturerItem_2.Text = "&Edit Current"
        '
        '_mnuManufacturerItem_3
        '
        Me.mnuManufacturerItem.SetIndex(Me._mnuManufacturerItem_3, CType(3, Short))
        Me._mnuManufacturerItem_3.Name = "_mnuManufacturerItem_3"
        Me._mnuManufacturerItem_3.Size = New System.Drawing.Size(150, 22)
        Me._mnuManufacturerItem_3.Text = "&Delete Current"
        '
        'mnuAdsorbent
        '
        Me.mnuAdsorbent.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuAdsorbentItem_1, Me._mnuAdsorbentItem_2, Me._mnuAdsorbentItem_3})
        Me.mnuAdsorbent.Name = "mnuAdsorbent"
        Me.mnuAdsorbent.Size = New System.Drawing.Size(74, 20)
        Me.mnuAdsorbent.Text = "&Adsorbent"
        '
        '_mnuAdsorbentItem_1
        '
        Me.mnuAdsorbentItem.SetIndex(Me._mnuAdsorbentItem_1, CType(1, Short))
        Me._mnuAdsorbentItem_1.Name = "_mnuAdsorbentItem_1"
        Me._mnuAdsorbentItem_1.Size = New System.Drawing.Size(150, 22)
        Me._mnuAdsorbentItem_1.Text = "&New"
        '
        '_mnuAdsorbentItem_2
        '
        Me.mnuAdsorbentItem.SetIndex(Me._mnuAdsorbentItem_2, CType(2, Short))
        Me._mnuAdsorbentItem_2.Name = "_mnuAdsorbentItem_2"
        Me._mnuAdsorbentItem_2.Size = New System.Drawing.Size(150, 22)
        Me._mnuAdsorbentItem_2.Text = "&Edit Current"
        '
        '_mnuAdsorbentItem_3
        '
        Me.mnuAdsorbentItem.SetIndex(Me._mnuAdsorbentItem_3, CType(3, Short))
        Me._mnuAdsorbentItem_3.Name = "_mnuAdsorbentItem_3"
        Me._mnuAdsorbentItem_3.Size = New System.Drawing.Size(150, 22)
        Me._mnuAdsorbentItem_3.Text = "&Delete Current"
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(624, 304)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 33
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        '_lblData_0
        '
        Me._lblData_0.BackColor = System.Drawing.Color.Transparent
        Me._lblData_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_0, CType(0, Short))
        Me._lblData_0.Location = New System.Drawing.Point(158, 35)
        Me._lblData_0.Name = "_lblData_0"
        Me._lblData_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_0.Size = New System.Drawing.Size(73, 17)
        Me._lblData_0.TabIndex = 26
        Me._lblData_0.Text = "lblData(0)"
        Me._lblData_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_3
        '
        Me._lblData_3.BackColor = System.Drawing.Color.Transparent
        Me._lblData_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_3, CType(3, Short))
        Me._lblData_3.Location = New System.Drawing.Point(158, 108)
        Me._lblData_3.Name = "_lblData_3"
        Me._lblData_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_3.Size = New System.Drawing.Size(143, 16)
        Me._lblData_3.TabIndex = 20
        Me._lblData_3.Text = "lblData(3)"
        Me._lblData_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_4
        '
        Me._lblData_4.BackColor = System.Drawing.Color.Transparent
        Me._lblData_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_4, CType(4, Short))
        Me._lblData_4.Location = New System.Drawing.Point(158, 131)
        Me._lblData_4.Name = "_lblData_4"
        Me._lblData_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_4.Size = New System.Drawing.Size(73, 17)
        Me._lblData_4.TabIndex = 18
        Me._lblData_4.Text = "lblData(4)"
        Me._lblData_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_5
        '
        Me._lblData_5.BackColor = System.Drawing.Color.Transparent
        Me._lblData_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_5, CType(5, Short))
        Me._lblData_5.Location = New System.Drawing.Point(158, 155)
        Me._lblData_5.Name = "_lblData_5"
        Me._lblData_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_5.Size = New System.Drawing.Size(73, 17)
        Me._lblData_5.TabIndex = 17
        Me._lblData_5.Text = "lblData(5)"
        Me._lblData_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_6
        '
        Me._lblData_6.BackColor = System.Drawing.Color.Transparent
        Me._lblData_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_6, CType(6, Short))
        Me._lblData_6.Location = New System.Drawing.Point(158, 179)
        Me._lblData_6.Name = "_lblData_6"
        Me._lblData_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_6.Size = New System.Drawing.Size(73, 17)
        Me._lblData_6.TabIndex = 16
        Me._lblData_6.Text = "lblData(6)"
        Me._lblData_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_1
        '
        Me._lblData_1.BackColor = System.Drawing.Color.Transparent
        Me._lblData_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_1, CType(1, Short))
        Me._lblData_1.Location = New System.Drawing.Point(158, 59)
        Me._lblData_1.Name = "_lblData_1"
        Me._lblData_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_1.Size = New System.Drawing.Size(73, 17)
        Me._lblData_1.TabIndex = 25
        Me._lblData_1.Text = "lblData(1)"
        Me._lblData_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblData_2
        '
        Me._lblData_2.BackColor = System.Drawing.Color.Transparent
        Me._lblData_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblData_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblData_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblData_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblData.SetIndex(Me._lblData_2, CType(2, Short))
        Me._lblData_2.Location = New System.Drawing.Point(158, 83)
        Me._lblData_2.Name = "_lblData_2"
        Me._lblData_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblData_2.Size = New System.Drawing.Size(73, 17)
        Me._lblData_2.TabIndex = 24
        Me._lblData_2.Text = "lblData(2)"
        Me._lblData_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblDesc_1
        '
        Me._lblDesc_1.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_1, CType(1, Short))
        Me._lblDesc_1.Location = New System.Drawing.Point(30, 60)
        Me._lblDesc_1.Name = "_lblDesc_1"
        Me._lblDesc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_1.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_1.TabIndex = 28
        Me._lblDesc_1.Text = "Particle Radius"
        Me._lblDesc_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_2
        '
        Me._lblDesc_2.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_2, CType(2, Short))
        Me._lblDesc_2.Location = New System.Drawing.Point(30, 84)
        Me._lblDesc_2.Name = "_lblDesc_2"
        Me._lblDesc_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_2.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_2.TabIndex = 27
        Me._lblDesc_2.Text = "Particle Porosity"
        Me._lblDesc_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_0
        '
        Me._lblDesc_0.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_0, CType(0, Short))
        Me._lblDesc_0.Location = New System.Drawing.Point(30, 36)
        Me._lblDesc_0.Name = "_lblDesc_0"
        Me._lblDesc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_0.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_0.TabIndex = 29
        Me._lblDesc_0.Text = "Apparent Density"
        Me._lblDesc_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_3
        '
        Me._lblDesc_3.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_3, CType(3, Short))
        Me._lblDesc_3.Location = New System.Drawing.Point(30, 108)
        Me._lblDesc_3.Name = "_lblDesc_3"
        Me._lblDesc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_3.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_3.TabIndex = 19
        Me._lblDesc_3.Text = "Adsorbent Type"
        Me._lblDesc_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_4
        '
        Me._lblDesc_4.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_4, CType(4, Short))
        Me._lblDesc_4.Location = New System.Drawing.Point(30, 132)
        Me._lblDesc_4.Name = "_lblDesc_4"
        Me._lblDesc_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_4.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_4.TabIndex = 15
        Me._lblDesc_4.Text = "Polanyi W0"
        Me._lblDesc_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_5
        '
        Me._lblDesc_5.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_5, CType(5, Short))
        Me._lblDesc_5.Location = New System.Drawing.Point(30, 156)
        Me._lblDesc_5.Name = "_lblDesc_5"
        Me._lblDesc_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_5.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_5.TabIndex = 14
        Me._lblDesc_5.Text = "Polanyi B"
        Me._lblDesc_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblDesc_6
        '
        Me._lblDesc_6.BackColor = System.Drawing.Color.Transparent
        Me._lblDesc_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDesc_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblDesc_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDesc.SetIndex(Me._lblDesc_6, CType(6, Short))
        Me._lblDesc_6.Location = New System.Drawing.Point(30, 180)
        Me._lblDesc_6.Name = "_lblDesc_6"
        Me._lblDesc_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDesc_6.Size = New System.Drawing.Size(117, 17)
        Me._lblDesc_6.TabIndex = 13
        Me._lblDesc_6.Text = "Polanyi Exponent"
        Me._lblDesc_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblUnit_0
        '
        Me._lblUnit_0.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_0, CType(0, Short))
        Me._lblUnit_0.Location = New System.Drawing.Point(238, 36)
        Me._lblUnit_0.Name = "_lblUnit_0"
        Me._lblUnit_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_0.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_0.TabIndex = 23
        Me._lblUnit_0.Text = "g/cm3"
        Me._lblUnit_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_1
        '
        Me._lblUnit_1.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_1, CType(1, Short))
        Me._lblUnit_1.Location = New System.Drawing.Point(238, 60)
        Me._lblUnit_1.Name = "_lblUnit_1"
        Me._lblUnit_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_1.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_1.TabIndex = 22
        Me._lblUnit_1.Text = "cm"
        Me._lblUnit_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_2
        '
        Me._lblUnit_2.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_2, CType(2, Short))
        Me._lblUnit_2.Location = New System.Drawing.Point(238, 84)
        Me._lblUnit_2.Name = "_lblUnit_2"
        Me._lblUnit_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_2.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_2.TabIndex = 21
        Me._lblUnit_2.Text = "-"
        Me._lblUnit_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_4
        '
        Me._lblUnit_4.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_4, CType(4, Short))
        Me._lblUnit_4.Location = New System.Drawing.Point(238, 132)
        Me._lblUnit_4.Name = "_lblUnit_4"
        Me._lblUnit_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_4.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_4.TabIndex = 12
        Me._lblUnit_4.Text = "cm3/g"
        Me._lblUnit_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_5
        '
        Me._lblUnit_5.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_5, CType(5, Short))
        Me._lblUnit_5.Location = New System.Drawing.Point(238, 156)
        Me._lblUnit_5.Name = "_lblUnit_5"
        Me._lblUnit_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_5.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_5.TabIndex = 11
        Me._lblUnit_5.Text = "*"
        Me._lblUnit_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblUnit_6
        '
        Me._lblUnit_6.BackColor = System.Drawing.Color.Transparent
        Me._lblUnit_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblUnit_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblUnit_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnit.SetIndex(Me._lblUnit_6, CType(6, Short))
        Me._lblUnit_6.Location = New System.Drawing.Point(238, 180)
        Me._lblUnit_6.Name = "_lblUnit_6"
        Me._lblUnit_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblUnit_6.Size = New System.Drawing.Size(65, 17)
        Me._lblUnit_6.TabIndex = 10
        Me._lblUnit_6.Text = "-"
        Me._lblUnit_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'mnuAdsorbentItem
        '
        '
        'mnuManufacturerItem
        '
        '
        'lstManu
        '
        Me.lstManu.BackColor = System.Drawing.SystemColors.Window
        Me.lstManu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstManu.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstManu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstManu.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstManu.ItemHeight = 14
        Me.lstManu.Location = New System.Drawing.Point(8, 26)
        Me.lstManu.Name = "lstManu"
        Me.lstManu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstManu.Size = New System.Drawing.Size(233, 142)
        Me.lstManu.TabIndex = 3
        '
        'lblEmpty_lstManu
        '
        Me.lblEmpty_lstManu.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmpty_lstManu.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmpty_lstManu.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpty_lstManu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmpty_lstManu.Location = New System.Drawing.Point(8, 26)
        Me.lblEmpty_lstManu.Name = "lblEmpty_lstManu"
        Me.lblEmpty_lstManu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmpty_lstManu.Size = New System.Drawing.Size(233, 25)
        Me.lblEmpty_lstManu.TabIndex = 31
        Me.lblEmpty_lstManu.Text = "No Manufacturers Available"
        Me.lblEmpty_lstManu.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEmpty_lstManu.Visible = False
        '
        'lblUnitB
        '
        Me.lblUnitB.BackColor = System.Drawing.Color.Transparent
        Me.lblUnitB.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUnitB.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnitB.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUnitB.Location = New System.Drawing.Point(64, 212)
        Me.lblUnitB.Name = "lblUnitB"
        Me.lblUnitB.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUnitB.Size = New System.Drawing.Size(242, 25)
        Me.lblUnitB.TabIndex = 9
        Me.lblUnitB.Text = "* in (mol/cal) ^(Polanyi Exponent)"
        '
        'lstName
        '
        Me.lstName.BackColor = System.Drawing.SystemColors.Window
        Me.lstName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstName.ItemHeight = 14
        Me.lstName.Location = New System.Drawing.Point(0, 80)
        Me.lstName.Name = "lstName"
        Me.lstName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstName.Size = New System.Drawing.Size(233, 128)
        Me.lstName.TabIndex = 6
        '
        'lblEmpty_lstName
        '
        Me.lblEmpty_lstName.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmpty_lstName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmpty_lstName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpty_lstName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmpty_lstName.Location = New System.Drawing.Point(6, 57)
        Me.lblEmpty_lstName.Name = "lblEmpty_lstName"
        Me.lblEmpty_lstName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmpty_lstName.Size = New System.Drawing.Size(219, 20)
        Me.lblEmpty_lstName.TabIndex = 30
        Me.lblEmpty_lstName.Text = "No Adsorbents Available"
        Me.lblEmpty_lstName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEmpty_lstName.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblEmpty_lstManu)
        Me.GroupBox1.Controls.Add(Me.lstManu)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(249, 180)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select a Manufacturer"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me._lblData_0)
        Me.GroupBox2.Controls.Add(Me._lblDesc_0)
        Me.GroupBox2.Controls.Add(Me._lblUnit_0)
        Me.GroupBox2.Controls.Add(Me._lblData_1)
        Me.GroupBox2.Controls.Add(Me._lblDesc_1)
        Me.GroupBox2.Controls.Add(Me._lblUnit_1)
        Me.GroupBox2.Controls.Add(Me._lblData_2)
        Me.GroupBox2.Controls.Add(Me._lblDesc_2)
        Me.GroupBox2.Controls.Add(Me._lblUnit_2)
        Me.GroupBox2.Controls.Add(Me._lblData_3)
        Me.GroupBox2.Controls.Add(Me._lblDesc_3)
        Me.GroupBox2.Controls.Add(Me._lblData_4)
        Me.GroupBox2.Controls.Add(Me._lblDesc_4)
        Me.GroupBox2.Controls.Add(Me._lblUnit_4)
        Me.GroupBox2.Controls.Add(Me._lblData_5)
        Me.GroupBox2.Controls.Add(Me._lblDesc_5)
        Me.GroupBox2.Controls.Add(Me._lblUnit_5)
        Me.GroupBox2.Controls.Add(Me._lblData_6)
        Me.GroupBox2.Controls.Add(Me._lblDesc_6)
        Me.GroupBox2.Controls.Add(Me._lblUnit_6)
        Me.GroupBox2.Controls.Add(Me.lblUnitB)
        Me.GroupBox2.Location = New System.Drawing.Point(284, 57)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(323, 270)
        Me.GroupBox2.TabIndex = 36
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Adsorbent Properties:"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Panel1)
        Me.GroupBox3.Controls.Add(Me.lblEmpty_lstName)
        Me.GroupBox3.Controls.Add(Me.lstName)
        Me.GroupBox3.Location = New System.Drawing.Point(20, 226)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(233, 209)
        Me.GroupBox3.TabIndex = 37
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Select an Adsorbent:"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me._optPhase_1)
        Me.Panel1.Controls.Add(Me._optPhase_0)
        Me.Panel1.Location = New System.Drawing.Point(3, 19)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(222, 35)
        Me.Panel1.TabIndex = 31
        '
        '_optPhase_1
        '
        Me._optPhase_1.AutoSize = True
        Me._optPhase_1.Location = New System.Drawing.Point(118, 7)
        Me._optPhase_1.Name = "_optPhase_1"
        Me._optPhase_1.Size = New System.Drawing.Size(78, 18)
        Me._optPhase_1.TabIndex = 39
        Me._optPhase_1.Text = "Gas Phase"
        Me._optPhase_1.UseVisualStyleBackColor = True
        '
        '_optPhase_0
        '
        Me._optPhase_0.AutoSize = True
        Me._optPhase_0.Checked = True
        Me._optPhase_0.Location = New System.Drawing.Point(6, 7)
        Me._optPhase_0.Name = "_optPhase_0"
        Me._optPhase_0.Size = New System.Drawing.Size(86, 18)
        Me._optPhase_0.TabIndex = 38
        Me._optPhase_0.TabStop = True
        Me._optPhase_0.Text = "Liquid Phase"
        Me._optPhase_0.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Location = New System.Drawing.Point(542, 411)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 41)
        Me.cmdCancel.TabIndex = 38
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Location = New System.Drawing.Point(284, 411)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(238, 41)
        Me.cmdOK.TabIndex = 39
        Me.cmdOK.Text = "&Use these Adsorbent Specifications"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'frmEditCarbon
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(666, 464)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(28, 83)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(682, 503)
        Me.Name = "frmEditCarbon"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Adsorbent Database"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuAdsorbentItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuManufacturerItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents _optPhase_1 As RadioButton
    Friend WithEvents _optPhase_0 As RadioButton
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdOK As Button
#End Region
End Class