<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFreundlich
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
    Public WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
    Public WithEvents Option1 As System.Windows.Forms.RadioButton
    Public WithEvents UserK As System.Windows.Forms.TextBox
    Public WithEvents UserOneOverN As System.Windows.Forms.TextBox
    Public WithEvents _lblText_4 As System.Windows.Forms.Label
    Public WithEvents _lblText_5 As System.Windows.Forms.Label
    Public WithEvents _Label5_2 As System.Windows.Forms.Label
    Public WithEvents lblWarning As System.Windows.Forms.Label
    Public WithEvents cboMethod As System.Windows.Forms.ComboBox
    Public WithEvents _txtInput_13 As System.Windows.Forms.TextBox
    Public WithEvents _txtInput_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtInput_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtInput_10 As System.Windows.Forms.TextBox
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents _lblInput_0 As System.Windows.Forms.Label
    Public WithEvents _lblInput_1 As System.Windows.Forms.Label
    Public WithEvents _lblInput_4 As System.Windows.Forms.Label
    Public WithEvents _txtInput_12 As System.Windows.Forms.TextBox
    Public WithEvents _txtInput_11 As System.Windows.Forms.TextBox
    Public WithEvents _lblInput_6 As System.Windows.Forms.Label
    Public WithEvents _lblInput_5 As System.Windows.Forms.Label
    Public WithEvents _lblText_3 As System.Windows.Forms.Label
    Public WithEvents _lblValue_5 As System.Windows.Forms.Label
    Public WithEvents _lblValue_4 As System.Windows.Forms.Label
    Public WithEvents _lblText_2 As System.Windows.Forms.Label
    Public WithEvents _Label5_1 As System.Windows.Forms.Label
    Public WithEvents lblEstimationMethod As System.Windows.Forms.Label
    Public WithEvents Label5 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    '   Public WithEvents cmdCancelOK As SSCommandArray
    '  Public WithEvents cmdFind As SSCommandArray
    Public WithEvents lblInput As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblText As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblValue As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lstRange As Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray
    '    Public WithEvents optFreundlichSource As SSOptionArray
    Public WithEvents txtInput As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Option1 = New System.Windows.Forms.RadioButton()
        Me.Label5 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblValue_2 = New System.Windows.Forms.Label()
        Me._Label5_1 = New System.Windows.Forms.Label()
        Me._Label5_0 = New System.Windows.Forms.Label()
        Me.lblInput = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblInput_0 = New System.Windows.Forms.Label()
        Me._lblInput_1 = New System.Windows.Forms.Label()
        Me._lblInput_4 = New System.Windows.Forms.Label()
        Me._lblInput_6 = New System.Windows.Forms.Label()
        Me._lblInput_5 = New System.Windows.Forms.Label()
        Me._lblInput_8 = New System.Windows.Forms.Label()
        Me.lblText = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblText_4 = New System.Windows.Forms.Label()
        Me._lblText_5 = New System.Windows.Forms.Label()
        Me._lblText_3 = New System.Windows.Forms.Label()
        Me._lblText_2 = New System.Windows.Forms.Label()
        Me._lblText_1 = New System.Windows.Forms.Label()
        Me._lblText_0 = New System.Windows.Forms.Label()
        Me.lblValue = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblValue_5 = New System.Windows.Forms.Label()
        Me._lblValue_4 = New System.Windows.Forms.Label()
        Me._lblValue_1 = New System.Windows.Forms.Label()
        Me._lblValue_0 = New System.Windows.Forms.Label()
        Me._lblValue_3 = New System.Windows.Forms.Label()
        Me.lstRange = New Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray(Me.components)
        Me._lstRange_0 = New System.Windows.Forms.ListBox()
        Me._lstRange_1 = New System.Windows.Forms.ListBox()
        Me.txtInput = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me._txtInput_13 = New System.Windows.Forms.TextBox()
        Me._txtInput_0 = New System.Windows.Forms.TextBox()
        Me._txtInput_1 = New System.Windows.Forms.TextBox()
        Me._txtInput_10 = New System.Windows.Forms.TextBox()
        Me._txtInput_12 = New System.Windows.Forms.TextBox()
        Me._txtInput_11 = New System.Windows.Forms.TextBox()
        Me.UserK = New System.Windows.Forms.TextBox()
        Me.UserOneOverN = New System.Windows.Forms.TextBox()
        Me._Label5_2 = New System.Windows.Forms.Label()
        Me.cboMethod = New System.Windows.Forms.ComboBox()
        Me.lblEstimationMethod = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.Line2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.grpSource = New System.Windows.Forms.GroupBox()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me._cmdCancelOK_0 = New System.Windows.Forms.Button()
        Me._cmdCancelOK_1 = New System.Windows.Forms.Button()
        Me.grpIPES = New System.Windows.Forms.GroupBox()
        Me.cmdCalculate = New System.Windows.Forms.Button()
        Me.grpAdditional = New System.Windows.Forms.GroupBox()
        Me.grpPolanyi = New System.Windows.Forms.GroupBox()
        Me.cmdEditPolanyi = New System.Windows.Forms.Button()
        Me.grpUserInput = New System.Windows.Forms.GroupBox()
        Me.grpOne = New System.Windows.Forms.GroupBox()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me.lblEmpty_lstCompo = New System.Windows.Forms.Label()
        Me._cmdFind_1 = New System.Windows.Forms.Button()
        Me.lstCompo = New System.Windows.Forms.ListBox()
        Me._cmdFind_0 = New System.Windows.Forms.Button()
        Me.cboSortMethod = New System.Windows.Forms.ComboBox()
        Me.grpTwo = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblTemperature = New System.Windows.Forms.Label()
        Me.lblPhase = New System.Windows.Forms.Label()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.grpIsothermDB = New System.Windows.Forms.GroupBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripDirty = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatus = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblInput, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblText, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lstRange, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtInput, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSource.SuspendLayout()
        Me.grpIPES.SuspendLayout()
        Me.grpAdditional.SuspendLayout()
        Me.grpPolanyi.SuspendLayout()
        Me.grpUserInput.SuspendLayout()
        Me.grpOne.SuspendLayout()
        Me.grpTwo.SuspendLayout()
        Me.grpIsothermDB.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(818, 148)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 71
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'Option1
        '
        Me.Option1.BackColor = System.Drawing.SystemColors.Control
        Me.Option1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Option1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option1.Location = New System.Drawing.Point(528, 488)
        Me.Option1.Name = "Option1"
        Me.Option1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Option1.Size = New System.Drawing.Size(83, 35)
        Me.Option1.TabIndex = 70
        Me.Option1.TabStop = True
        Me.Option1.Text = "Option1"
        Me.Option1.UseVisualStyleBackColor = False
        Me.Option1.Visible = False
        '
        '_lblValue_2
        '
        Me._lblValue_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_2, CType(2, Short))
        Me.Label5.SetIndex(Me._lblValue_2, CType(2, Short))
        Me._lblValue_2.Location = New System.Drawing.Point(157, 180)
        Me._lblValue_2.Name = "_lblValue_2"
        Me._lblValue_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_2.Size = New System.Drawing.Size(207, 25)
        Me._lblValue_2.TabIndex = 47
        Me._lblValue_2.Text = "lblValue(2)"
        Me._lblValue_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label5_1
        '
        Me._Label5_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._Label5_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label5_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label5_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.SetIndex(Me._Label5_1, CType(1, Short))
        Me._Label5_1.Location = New System.Drawing.Point(498, 134)
        Me._Label5_1.Name = "_Label5_1"
        Me._Label5_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label5_1.Size = New System.Drawing.Size(129, 26)
        Me._Label5_1.TabIndex = 8
        Me._Label5_1.Text = "(mg/g)*(L/mg)^(1/n)"
        '
        '_Label5_0
        '
        Me._Label5_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._Label5_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label5_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label5_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.SetIndex(Me._Label5_0, CType(0, Short))
        Me._Label5_0.Location = New System.Drawing.Point(230, 19)
        Me._Label5_0.Name = "_Label5_0"
        Me._Label5_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label5_0.Size = New System.Drawing.Size(134, 27)
        Me._Label5_0.TabIndex = 44
        Me._Label5_0.Text = "(mg/g)*(L/mg)^(1/n)"
        '
        '_lblInput_0
        '
        Me._lblInput_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_0, CType(0, Short))
        Me._lblInput_0.Location = New System.Drawing.Point(127, 34)
        Me._lblInput_0.Name = "_lblInput_0"
        Me._lblInput_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_0.Size = New System.Drawing.Size(101, 17)
        Me._lblInput_0.TabIndex = 20
        Me._lblInput_0.Text = "W0 (cm3/g)"
        Me._lblInput_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblInput_1
        '
        Me._lblInput_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_1, CType(1, Short))
        Me._lblInput_1.Location = New System.Drawing.Point(102, 51)
        Me._lblInput_1.Name = "_lblInput_1"
        Me._lblInput_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_1.Size = New System.Drawing.Size(126, 19)
        Me._lblInput_1.TabIndex = 19
        Me._lblInput_1.Text = "BB (mol/cal)^GM"
        Me._lblInput_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblInput_4
        '
        Me._lblInput_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_4, CType(4, Short))
        Me._lblInput_4.Location = New System.Drawing.Point(122, 70)
        Me._lblInput_4.Name = "_lblInput_4"
        Me._lblInput_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_4.Size = New System.Drawing.Size(105, 13)
        Me._lblInput_4.TabIndex = 18
        Me._lblInput_4.Text = "GM"
        Me._lblInput_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblInput_6
        '
        Me._lblInput_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_6, CType(6, Short))
        Me._lblInput_6.Location = New System.Drawing.Point(19, 41)
        Me._lblInput_6.Name = "_lblInput_6"
        Me._lblInput_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_6.Size = New System.Drawing.Size(153, 21)
        Me._lblInput_6.TabIndex = 26
        Me._lblInput_6.Text = "No. of regression points:"
        Me._lblInput_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblInput_5
        '
        Me._lblInput_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_5, CType(5, Short))
        Me._lblInput_5.Location = New System.Drawing.Point(19, 20)
        Me._lblInput_5.Name = "_lblInput_5"
        Me._lblInput_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_5.Size = New System.Drawing.Size(153, 21)
        Me._lblInput_5.TabIndex = 25
        Me._lblInput_5.Text = "Order of magnitude:"
        Me._lblInput_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblInput_8
        '
        Me._lblInput_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._lblInput_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblInput_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblInput_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblInput.SetIndex(Me._lblInput_8, CType(8, Short))
        Me._lblInput_8.Location = New System.Drawing.Point(17, 259)
        Me._lblInput_8.Name = "_lblInput_8"
        Me._lblInput_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblInput_8.Size = New System.Drawing.Size(151, 24)
        Me._lblInput_8.TabIndex = 34
        Me._lblInput_8.Text = "Sorting Method:"
        Me._lblInput_8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_4
        '
        Me._lblText_4.BackColor = System.Drawing.Color.Transparent
        Me._lblText_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_4, CType(4, Short))
        Me._lblText_4.Location = New System.Drawing.Point(115, 20)
        Me._lblText_4.Name = "_lblText_4"
        Me._lblText_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_4.Size = New System.Drawing.Size(21, 17)
        Me._lblText_4.TabIndex = 65
        Me._lblText_4.Text = "K"
        Me._lblText_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_5
        '
        Me._lblText_5.BackColor = System.Drawing.Color.Transparent
        Me._lblText_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_5, CType(5, Short))
        Me._lblText_5.Location = New System.Drawing.Point(8, 20)
        Me._lblText_5.Name = "_lblText_5"
        Me._lblText_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_5.Size = New System.Drawing.Size(25, 17)
        Me._lblText_5.TabIndex = 64
        Me._lblText_5.Text = "1/n"
        Me._lblText_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_3
        '
        Me._lblText_3.BackColor = System.Drawing.Color.Transparent
        Me._lblText_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_3, CType(3, Short))
        Me._lblText_3.Location = New System.Drawing.Point(388, 135)
        Me._lblText_3.Name = "_lblText_3"
        Me._lblText_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_3.Size = New System.Drawing.Size(26, 26)
        Me._lblText_3.TabIndex = 12
        Me._lblText_3.Text = "K"
        Me._lblText_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_2
        '
        Me._lblText_2.BackColor = System.Drawing.Color.Transparent
        Me._lblText_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_2, CType(2, Short))
        Me._lblText_2.Location = New System.Drawing.Point(272, 134)
        Me._lblText_2.Name = "_lblText_2"
        Me._lblText_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_2.Size = New System.Drawing.Size(40, 29)
        Me._lblText_2.TabIndex = 9
        Me._lblText_2.Text = "1/n"
        Me._lblText_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_1
        '
        Me._lblText_1.BackColor = System.Drawing.Color.Transparent
        Me._lblText_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_1, CType(1, Short))
        Me._lblText_1.Location = New System.Drawing.Point(9, 22)
        Me._lblText_1.Name = "_lblText_1"
        Me._lblText_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_1.Size = New System.Drawing.Size(38, 16)
        Me._lblText_1.TabIndex = 51
        Me._lblText_1.Text = "1/n"
        Me._lblText_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblText_0
        '
        Me._lblText_0.BackColor = System.Drawing.Color.Transparent
        Me._lblText_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblText_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblText_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblText.SetIndex(Me._lblText_0, CType(0, Short))
        Me._lblText_0.Location = New System.Drawing.Point(128, 22)
        Me._lblText_0.Name = "_lblText_0"
        Me._lblText_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblText_0.Size = New System.Drawing.Size(14, 16)
        Me._lblText_0.TabIndex = 54
        Me._lblText_0.Text = "K"
        Me._lblText_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblValue_5
        '
        Me._lblValue_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_5, CType(5, Short))
        Me._lblValue_5.Location = New System.Drawing.Point(419, 134)
        Me._lblValue_5.Name = "_lblValue_5"
        Me._lblValue_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_5.Size = New System.Drawing.Size(78, 26)
        Me._lblValue_5.TabIndex = 11
        Me._lblValue_5.Text = "lblValue(5)"
        Me._lblValue_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_4
        '
        Me._lblValue_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_4, CType(4, Short))
        Me._lblValue_4.Location = New System.Drawing.Point(312, 133)
        Me._lblValue_4.Name = "_lblValue_4"
        Me._lblValue_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_4.Size = New System.Drawing.Size(81, 26)
        Me._lblValue_4.TabIndex = 10
        Me._lblValue_4.Text = "lblValue(4)"
        Me._lblValue_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_1
        '
        Me._lblValue_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_1, CType(1, Short))
        Me._lblValue_1.Location = New System.Drawing.Point(48, 21)
        Me._lblValue_1.Name = "_lblValue_1"
        Me._lblValue_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_1.Size = New System.Drawing.Size(79, 25)
        Me._lblValue_1.TabIndex = 52
        Me._lblValue_1.Text = "lblValue(1)"
        Me._lblValue_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_0
        '
        Me._lblValue_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_0, CType(0, Short))
        Me._lblValue_0.Location = New System.Drawing.Point(145, 21)
        Me._lblValue_0.Name = "_lblValue_0"
        Me._lblValue_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_0.Size = New System.Drawing.Size(79, 25)
        Me._lblValue_0.TabIndex = 53
        Me._lblValue_0.Text = "lblValue(0)"
        Me._lblValue_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblValue_3
        '
        Me._lblValue_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblValue_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblValue_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblValue_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblValue_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblValue.SetIndex(Me._lblValue_3, CType(3, Short))
        Me._lblValue_3.Location = New System.Drawing.Point(98, 239)
        Me._lblValue_3.Name = "_lblValue_3"
        Me._lblValue_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblValue_3.Size = New System.Drawing.Size(266, 18)
        Me._lblValue_3.TabIndex = 45
        Me._lblValue_3.Text = "lblValue(3)"
        Me._lblValue_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lstRange
        '
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
        Me._lstRange_0.Location = New System.Drawing.Point(14, 76)
        Me._lstRange_0.Name = "_lstRange_0"
        Me._lstRange_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRange_0.Size = New System.Drawing.Size(128, 86)
        Me._lstRange_0.TabIndex = 37
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
        Me._lstRange_1.Location = New System.Drawing.Point(168, 76)
        Me._lstRange_1.Name = "_lstRange_1"
        Me._lstRange_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstRange_1.Size = New System.Drawing.Size(196, 86)
        Me._lstRange_1.TabIndex = 36
        '
        'txtInput
        '
        '
        '_txtInput_13
        '
        Me._txtInput_13.AcceptsReturn = True
        Me._txtInput_13.BackColor = System.Drawing.SystemColors.Control
        Me._txtInput_13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_13.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_13, CType(13, Short))
        Me._txtInput_13.Location = New System.Drawing.Point(135, 10)
        Me._txtInput_13.MaxLength = 0
        Me._txtInput_13.Name = "_txtInput_13"
        Me._txtInput_13.ReadOnly = True
        Me._txtInput_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_13.Size = New System.Drawing.Size(182, 20)
        Me._txtInput_13.TabIndex = 17
        Me._txtInput_13.Text = "txtInput(13)"
        Me._txtInput_13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtInput_0
        '
        Me._txtInput_0.AcceptsReturn = True
        Me._txtInput_0.BackColor = System.Drawing.SystemColors.Control
        Me._txtInput_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_0, CType(0, Short))
        Me._txtInput_0.Location = New System.Drawing.Point(227, 28)
        Me._txtInput_0.MaxLength = 0
        Me._txtInput_0.Name = "_txtInput_0"
        Me._txtInput_0.ReadOnly = True
        Me._txtInput_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_0.Size = New System.Drawing.Size(90, 20)
        Me._txtInput_0.TabIndex = 16
        Me._txtInput_0.Text = "txtInput(0)"
        Me._txtInput_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtInput_1
        '
        Me._txtInput_1.AcceptsReturn = True
        Me._txtInput_1.BackColor = System.Drawing.SystemColors.Control
        Me._txtInput_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_1, CType(1, Short))
        Me._txtInput_1.Location = New System.Drawing.Point(227, 48)
        Me._txtInput_1.MaxLength = 0
        Me._txtInput_1.Name = "_txtInput_1"
        Me._txtInput_1.ReadOnly = True
        Me._txtInput_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_1.Size = New System.Drawing.Size(90, 20)
        Me._txtInput_1.TabIndex = 15
        Me._txtInput_1.Text = "txtInput(1)"
        Me._txtInput_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtInput_10
        '
        Me._txtInput_10.AcceptsReturn = True
        Me._txtInput_10.BackColor = System.Drawing.SystemColors.Control
        Me._txtInput_10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_10.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_10, CType(10, Short))
        Me._txtInput_10.Location = New System.Drawing.Point(227, 68)
        Me._txtInput_10.MaxLength = 0
        Me._txtInput_10.Name = "_txtInput_10"
        Me._txtInput_10.ReadOnly = True
        Me._txtInput_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_10.Size = New System.Drawing.Size(90, 20)
        Me._txtInput_10.TabIndex = 14
        Me._txtInput_10.Text = "txtInput(10)"
        Me._txtInput_10.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtInput_12
        '
        Me._txtInput_12.AcceptsReturn = True
        Me._txtInput_12.BackColor = System.Drawing.SystemColors.Window
        Me._txtInput_12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_12.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_12, CType(12, Short))
        Me._txtInput_12.Location = New System.Drawing.Point(183, 39)
        Me._txtInput_12.MaxLength = 0
        Me._txtInput_12.Name = "_txtInput_12"
        Me._txtInput_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_12.Size = New System.Drawing.Size(87, 20)
        Me._txtInput_12.TabIndex = 24
        Me._txtInput_12.Text = "txtInput(12)"
        Me._txtInput_12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_txtInput_11
        '
        Me._txtInput_11.AcceptsReturn = True
        Me._txtInput_11.BackColor = System.Drawing.SystemColors.Window
        Me._txtInput_11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._txtInput_11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtInput_11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtInput_11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInput.SetIndex(Me._txtInput_11, CType(11, Short))
        Me._txtInput_11.Location = New System.Drawing.Point(183, 17)
        Me._txtInput_11.MaxLength = 0
        Me._txtInput_11.Name = "_txtInput_11"
        Me._txtInput_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtInput_11.Size = New System.Drawing.Size(87, 20)
        Me._txtInput_11.TabIndex = 23
        Me._txtInput_11.Text = "txtInput(11)"
        Me._txtInput_11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'UserK
        '
        Me.UserK.AcceptsReturn = True
        Me.UserK.BackColor = System.Drawing.SystemColors.Window
        Me.UserK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.UserK.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.UserK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserK.ForeColor = System.Drawing.SystemColors.WindowText
        Me.UserK.Location = New System.Drawing.Point(139, 18)
        Me.UserK.MaxLength = 0
        Me.UserK.Name = "UserK"
        Me.UserK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.UserK.Size = New System.Drawing.Size(65, 20)
        Me.UserK.TabIndex = 62
        Me.UserK.Text = "UserK"
        Me.UserK.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'UserOneOverN
        '
        Me.UserOneOverN.AcceptsReturn = True
        Me.UserOneOverN.BackColor = System.Drawing.SystemColors.Window
        Me.UserOneOverN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.UserOneOverN.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.UserOneOverN.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserOneOverN.ForeColor = System.Drawing.SystemColors.WindowText
        Me.UserOneOverN.Location = New System.Drawing.Point(40, 18)
        Me.UserOneOverN.MaxLength = 0
        Me.UserOneOverN.Name = "UserOneOverN"
        Me.UserOneOverN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.UserOneOverN.Size = New System.Drawing.Size(78, 20)
        Me.UserOneOverN.TabIndex = 61
        Me.UserOneOverN.Text = "UserOneOverN"
        Me.UserOneOverN.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '_Label5_2
        '
        Me._Label5_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me._Label5_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label5_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label5_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._Label5_2.Location = New System.Drawing.Point(207, 20)
        Me._Label5_2.Name = "_Label5_2"
        Me._Label5_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label5_2.Size = New System.Drawing.Size(139, 21)
        Me._Label5_2.TabIndex = 63
        Me._Label5_2.Text = "(mg/g)*(L/mg)^(1/n)"
        '
        'cboMethod
        '
        Me.cboMethod.BackColor = System.Drawing.SystemColors.Window
        Me.cboMethod.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMethod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMethod.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboMethod.Location = New System.Drawing.Point(6, 52)
        Me.cboMethod.Name = "cboMethod"
        Me.cboMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboMethod.Size = New System.Drawing.Size(273, 22)
        Me.cboMethod.TabIndex = 6
        '
        'lblEstimationMethod
        '
        Me.lblEstimationMethod.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstimationMethod.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstimationMethod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEstimationMethod.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblEstimationMethod.Location = New System.Drawing.Point(6, 21)
        Me.lblEstimationMethod.Name = "lblEstimationMethod"
        Me.lblEstimationMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstimationMethod.Size = New System.Drawing.Size(273, 31)
        Me.lblEstimationMethod.TabIndex = 7
        Me.lblEstimationMethod.Text = "Estimation Method:"
        Me.lblEstimationMethod.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.Line2})
        Me.ShapeContainer1.Size = New System.Drawing.Size(309, 255)
        Me.ShapeContainer1.TabIndex = 55
        Me.ShapeContainer1.TabStop = False
        '
        'Line2
        '
        Me.Line2.BorderColor = System.Drawing.SystemColors.WindowText
        Me.Line2.Name = "Line2"
        Me.Line2.X1 = 2
        Me.Line2.X2 = 306
        Me.Line2.Y1 = 42
        Me.Line2.Y2 = 42
        '
        'lblWarning
        '
        Me.lblWarning.BackColor = System.Drawing.SystemColors.Control
        Me.lblWarning.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWarning.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarning.ForeColor = System.Drawing.Color.Red
        Me.lblWarning.Location = New System.Drawing.Point(223, 14)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWarning.Size = New System.Drawing.Size(271, 76)
        Me.lblWarning.TabIndex = 67
        Me.lblWarning.Text = "lblWarning"
        Me.lblWarning.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label10.Location = New System.Drawing.Point(33, 18)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(92, 20)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Adsorbent:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'grpSource
        '
        Me.grpSource.Controls.Add(Me.RadioButton3)
        Me.grpSource.Controls.Add(Me.RadioButton2)
        Me.grpSource.Controls.Add(Me.RadioButton1)
        Me.grpSource.Controls.Add(Me._cmdCancelOK_0)
        Me.grpSource.Controls.Add(Me._cmdCancelOK_1)
        Me.grpSource.Controls.Add(Me.lblWarning)
        Me.grpSource.Location = New System.Drawing.Point(3, 4)
        Me.grpSource.Name = "grpSource"
        Me.grpSource.Size = New System.Drawing.Size(645, 110)
        Me.grpSource.TabIndex = 72
        Me.grpSource.TabStop = False
        Me.grpSource.Text = "Source of K and 1/n"
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(23, 78)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(77, 18)
        Me.RadioButton3.TabIndex = 87
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "U&ser Input "
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(23, 53)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(177, 18)
        Me.RadioButton2.TabIndex = 86
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "(Isotherm Parameter &Estimation)"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(23, 28)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(115, 18)
        Me.RadioButton1.TabIndex = 85
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Isotherm &Database"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(513, 14)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(101, 22)
        Me._cmdCancelOK_0.TabIndex = 83
        Me._cmdCancelOK_0.Text = "&Cancel"
        Me._cmdCancelOK_0.UseVisualStyleBackColor = False
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(513, 40)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(101, 22)
        Me._cmdCancelOK_1.TabIndex = 84
        Me._cmdCancelOK_1.Text = "&Ok"
        Me._cmdCancelOK_1.UseVisualStyleBackColor = False
        '
        'grpIPES
        '
        Me.grpIPES.Controls.Add(Me.cmdCalculate)
        Me.grpIPES.Controls.Add(Me.grpAdditional)
        Me.grpIPES.Controls.Add(Me.grpPolanyi)
        Me.grpIPES.Controls.Add(Me.lblEstimationMethod)
        Me.grpIPES.Controls.Add(Me.cboMethod)
        Me.grpIPES.Controls.Add(Me._lblText_2)
        Me.grpIPES.Controls.Add(Me._lblValue_4)
        Me.grpIPES.Controls.Add(Me._lblText_3)
        Me.grpIPES.Controls.Add(Me._lblValue_5)
        Me.grpIPES.Controls.Add(Me._Label5_1)
        Me.grpIPES.Location = New System.Drawing.Point(669, 44)
        Me.grpIPES.Name = "grpIPES"
        Me.grpIPES.Size = New System.Drawing.Size(634, 182)
        Me.grpIPES.TabIndex = 73
        Me.grpIPES.TabStop = False
        Me.grpIPES.Text = "Isotherm Parameter Estimation (IPE)"
        '
        'cmdCalculate
        '
        Me.cmdCalculate.Location = New System.Drawing.Point(297, 108)
        Me.cmdCalculate.Name = "cmdCalculate"
        Me.cmdCalculate.Size = New System.Drawing.Size(239, 21)
        Me.cmdCalculate.TabIndex = 84
        Me.cmdCalculate.Text = "&Perform IPE Calculations"
        Me.cmdCalculate.UseVisualStyleBackColor = True
        '
        'grpAdditional
        '
        Me.grpAdditional.Controls.Add(Me._lblInput_5)
        Me.grpAdditional.Controls.Add(Me._lblInput_6)
        Me.grpAdditional.Controls.Add(Me._txtInput_11)
        Me.grpAdditional.Controls.Add(Me._txtInput_12)
        Me.grpAdditional.Location = New System.Drawing.Point(3, 85)
        Me.grpAdditional.Name = "grpAdditional"
        Me.grpAdditional.Size = New System.Drawing.Size(276, 73)
        Me.grpAdditional.TabIndex = 14
        Me.grpAdditional.TabStop = False
        Me.grpAdditional.Text = "Additional Parameters"
        '
        'grpPolanyi
        '
        Me.grpPolanyi.Controls.Add(Me.Label10)
        Me.grpPolanyi.Controls.Add(Me._txtInput_13)
        Me.grpPolanyi.Controls.Add(Me._lblInput_0)
        Me.grpPolanyi.Controls.Add(Me.cmdEditPolanyi)
        Me.grpPolanyi.Controls.Add(Me._lblInput_1)
        Me.grpPolanyi.Controls.Add(Me._lblInput_4)
        Me.grpPolanyi.Controls.Add(Me._txtInput_0)
        Me.grpPolanyi.Controls.Add(Me._txtInput_1)
        Me.grpPolanyi.Controls.Add(Me._txtInput_10)
        Me.grpPolanyi.Location = New System.Drawing.Point(291, 14)
        Me.grpPolanyi.Name = "grpPolanyi"
        Me.grpPolanyi.Size = New System.Drawing.Size(337, 93)
        Me.grpPolanyi.TabIndex = 13
        Me.grpPolanyi.TabStop = False
        Me.grpPolanyi.Text = "Polanyi Parameters"
        '
        'cmdEditPolanyi
        '
        Me.cmdEditPolanyi.Location = New System.Drawing.Point(6, 51)
        Me.cmdEditPolanyi.Name = "cmdEditPolanyi"
        Me.cmdEditPolanyi.Size = New System.Drawing.Size(90, 21)
        Me.cmdEditPolanyi.TabIndex = 85
        Me.cmdEditPolanyi.Text = "Edi&t Parameters"
        Me.cmdEditPolanyi.UseVisualStyleBackColor = True
        '
        'grpUserInput
        '
        Me.grpUserInput.Controls.Add(Me._lblText_5)
        Me.grpUserInput.Controls.Add(Me.UserOneOverN)
        Me.grpUserInput.Controls.Add(Me._lblText_4)
        Me.grpUserInput.Controls.Add(Me.UserK)
        Me.grpUserInput.Controls.Add(Me._Label5_2)
        Me.grpUserInput.Location = New System.Drawing.Point(187, 524)
        Me.grpUserInput.Name = "grpUserInput"
        Me.grpUserInput.Size = New System.Drawing.Size(376, 50)
        Me.grpUserInput.TabIndex = 81
        Me.grpUserInput.TabStop = False
        Me.grpUserInput.Text = "User Input:"
        '
        'grpOne
        '
        Me.grpOne.Controls.Add(Me.cmdSelect)
        Me.grpOne.Controls.Add(Me.lblEmpty_lstCompo)
        Me.grpOne.Controls.Add(Me._cmdFind_1)
        Me.grpOne.Controls.Add(Me.lstCompo)
        Me.grpOne.Controls.Add(Me._cmdFind_0)
        Me.grpOne.Controls.Add(Me._lblInput_8)
        Me.grpOne.Controls.Add(Me.cboSortMethod)
        Me.grpOne.Location = New System.Drawing.Point(11, 22)
        Me.grpOne.Name = "grpOne"
        Me.grpOne.Size = New System.Drawing.Size(313, 299)
        Me.grpOne.TabIndex = 80
        Me.grpOne.TabStop = False
        Me.grpOne.Text = "Select a Component:"
        '
        'cmdSelect
        '
        Me.cmdSelect.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmdSelect.Location = New System.Drawing.Point(30, 221)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(206, 23)
        Me.cmdSelect.TabIndex = 88
        Me.cmdSelect.Text = "Select Chemic&al"
        Me.cmdSelect.UseVisualStyleBackColor = False
        '
        'lblEmpty_lstCompo
        '
        Me.lblEmpty_lstCompo.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmpty_lstCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEmpty_lstCompo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpty_lstCompo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmpty_lstCompo.Location = New System.Drawing.Point(8, 24)
        Me.lblEmpty_lstCompo.Name = "lblEmpty_lstCompo"
        Me.lblEmpty_lstCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEmpty_lstCompo.Size = New System.Drawing.Size(273, 25)
        Me.lblEmpty_lstCompo.TabIndex = 69
        Me.lblEmpty_lstCompo.Text = "No Components Available"
        Me.lblEmpty_lstCompo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblEmpty_lstCompo.Visible = False
        '
        '_cmdFind_1
        '
        Me._cmdFind_1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdFind_1.Location = New System.Drawing.Point(136, 191)
        Me._cmdFind_1.Name = "_cmdFind_1"
        Me._cmdFind_1.Size = New System.Drawing.Size(100, 24)
        Me._cmdFind_1.TabIndex = 87
        Me._cmdFind_1.Text = "Find A&gain"
        Me._cmdFind_1.UseVisualStyleBackColor = False
        '
        'lstCompo
        '
        Me.lstCompo.BackColor = System.Drawing.SystemColors.Window
        Me.lstCompo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstCompo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCompo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCompo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCompo.ItemHeight = 14
        Me.lstCompo.Location = New System.Drawing.Point(8, 45)
        Me.lstCompo.Name = "lstCompo"
        Me.lstCompo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCompo.Size = New System.Drawing.Size(277, 128)
        Me.lstCompo.TabIndex = 33
        '
        '_cmdFind_0
        '
        Me._cmdFind_0.BackColor = System.Drawing.SystemColors.ButtonFace
        Me._cmdFind_0.Location = New System.Drawing.Point(30, 191)
        Me._cmdFind_0.Name = "_cmdFind_0"
        Me._cmdFind_0.Size = New System.Drawing.Size(100, 24)
        Me._cmdFind_0.TabIndex = 86
        Me._cmdFind_0.Text = "&Find"
        Me._cmdFind_0.UseVisualStyleBackColor = False
        '
        'cboSortMethod
        '
        Me.cboSortMethod.BackColor = System.Drawing.SystemColors.Window
        Me.cboSortMethod.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSortMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSortMethod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSortMethod.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSortMethod.Location = New System.Drawing.Point(182, 256)
        Me.cboSortMethod.Name = "cboSortMethod"
        Me.cboSortMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSortMethod.Size = New System.Drawing.Size(89, 22)
        Me.cboSortMethod.TabIndex = 29
        '
        'grpTwo
        '
        Me.grpTwo.Controls.Add(Me._lblValue_2)
        Me.grpTwo.Controls.Add(Me._lblText_1)
        Me.grpTwo.Controls.Add(Me._lblValue_1)
        Me.grpTwo.Controls.Add(Me._lblText_0)
        Me.grpTwo.Controls.Add(Me._lblValue_0)
        Me.grpTwo.Controls.Add(Me._Label5_0)
        Me.grpTwo.Controls.Add(Me.Label1)
        Me.grpTwo.Controls.Add(Me.Label2)
        Me.grpTwo.Controls.Add(Me._lstRange_0)
        Me.grpTwo.Controls.Add(Me._lstRange_1)
        Me.grpTwo.Controls.Add(Me.Label3)
        Me.grpTwo.Controls.Add(Me.Label4)
        Me.grpTwo.Controls.Add(Me.Label7)
        Me.grpTwo.Controls.Add(Me.Label8)
        Me.grpTwo.Controls.Add(Me.Label9)
        Me.grpTwo.Controls.Add(Me.lblTemperature)
        Me.grpTwo.Controls.Add(Me.lblPhase)
        Me.grpTwo.Controls.Add(Me.lblComments)
        Me.grpTwo.Controls.Add(Me._lblValue_3)
        Me.grpTwo.Location = New System.Drawing.Point(330, 22)
        Me.grpTwo.Name = "grpTwo"
        Me.grpTwo.Size = New System.Drawing.Size(409, 299)
        Me.grpTwo.TabIndex = 81
        Me.grpTwo.TabStop = False
        Me.grpTwo.Text = "{X} {What}-Phase isotherms(s) for {chemical name}"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(168, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(196, 22)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "Conc. Range (mg/L):"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label2.Location = New System.Drawing.Point(14, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(128, 17)
        Me.Label2.TabIndex = 49
        Me.Label2.Text = "pH Range:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(12, 181)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(130, 24)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "Adsorbent Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(14, 239)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(66, 18)
        Me.Label4.TabIndex = 46
        Me.Label4.Text = "Source:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(6, 213)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(132, 22)
        Me.Label7.TabIndex = 42
        Me.Label7.Text = "Temperature (C):"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label8.Location = New System.Drawing.Point(227, 209)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(52, 23)
        Me.Label8.TabIndex = 40
        Me.Label8.Text = "Phase:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label9.Location = New System.Drawing.Point(6, 254)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(89, 16)
        Me.Label9.TabIndex = 38
        Me.Label9.Text = "Comment:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTemperature
        '
        Me.lblTemperature.BackColor = System.Drawing.SystemColors.Control
        Me.lblTemperature.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTemperature.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTemperature.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTemperature.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblTemperature.Location = New System.Drawing.Point(144, 209)
        Me.lblTemperature.Name = "lblTemperature"
        Me.lblTemperature.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTemperature.Size = New System.Drawing.Size(80, 23)
        Me.lblTemperature.TabIndex = 41
        Me.lblTemperature.Text = "lblTemperature"
        Me.lblTemperature.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblPhase
        '
        Me.lblPhase.BackColor = System.Drawing.SystemColors.Control
        Me.lblPhase.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPhase.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPhase.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhase.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblPhase.Location = New System.Drawing.Point(284, 209)
        Me.lblPhase.Name = "lblPhase"
        Me.lblPhase.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPhase.Size = New System.Drawing.Size(83, 26)
        Me.lblPhase.TabIndex = 43
        Me.lblPhase.Text = "lblPhase"
        Me.lblPhase.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblComments
        '
        Me.lblComments.BackColor = System.Drawing.SystemColors.Control
        Me.lblComments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblComments.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblComments.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblComments.Location = New System.Drawing.Point(98, 255)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblComments.Size = New System.Drawing.Size(266, 21)
        Me.lblComments.TabIndex = 39
        Me.lblComments.Text = "lblComments"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'grpIsothermDB
        '
        Me.grpIsothermDB.Controls.Add(Me.grpTwo)
        Me.grpIsothermDB.Controls.Add(Me.grpOne)
        Me.grpIsothermDB.Location = New System.Drawing.Point(15, 227)
        Me.grpIsothermDB.Name = "grpIsothermDB"
        Me.grpIsothermDB.Size = New System.Drawing.Size(750, 360)
        Me.grpIsothermDB.TabIndex = 82
        Me.grpIsothermDB.TabStop = False
        Me.grpIsothermDB.Text = "{What} PhaseIsotherm Database"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripDirty, Me.ToolStripStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 635)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(772, 22)
        Me.StatusStrip1.TabIndex = 83
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripDirty
        '
        Me.ToolStripDirty.Name = "ToolStripDirty"
        Me.ToolStripDirty.Size = New System.Drawing.Size(78, 17)
        Me.ToolStripDirty.Text = "ToolStripDirty"
        '
        'ToolStripStatus
        '
        Me.ToolStripStatus.Name = "ToolStripStatus"
        Me.ToolStripStatus.Size = New System.Drawing.Size(85, 17)
        Me.ToolStripStatus.Text = "ToolStripStatus"
        '
        'frmFreundlich
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ClientSize = New System.Drawing.Size(772, 657)
        Me.ControlBox = False
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.grpIsothermDB)
        Me.Controls.Add(Me.grpUserInput)
        Me.Controls.Add(Me.grpIPES)
        Me.Controls.Add(Me.grpSource)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Option1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(272, 45)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFreundlich"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Freundlich Isotherm Parameters for {ComponentName}"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblInput, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblText, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lstRange, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtInput, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSource.ResumeLayout(False)
        Me.grpSource.PerformLayout()
        Me.grpIPES.ResumeLayout(False)
        Me.grpAdditional.ResumeLayout(False)
        Me.grpAdditional.PerformLayout()
        Me.grpPolanyi.ResumeLayout(False)
        Me.grpPolanyi.PerformLayout()
        Me.grpUserInput.ResumeLayout(False)
        Me.grpUserInput.PerformLayout()
        Me.grpOne.ResumeLayout(False)
        Me.grpTwo.ResumeLayout(False)
        Me.grpIsothermDB.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpSource As GroupBox
    Friend WithEvents grpIPES As GroupBox
    Friend WithEvents grpAdditional As GroupBox
    Friend WithEvents grpPolanyi As GroupBox
    Friend WithEvents grpUserInput As GroupBox
    Friend WithEvents grpOne As GroupBox
    Public WithEvents lblEmpty_lstCompo As Label
    Public WithEvents lstCompo As ListBox
    Public WithEvents _lblInput_8 As Label
    Public WithEvents cboSortMethod As ComboBox
    Friend WithEvents grpTwo As GroupBox
    Public WithEvents _lblValue_2 As Label
    Public WithEvents _lblText_1 As Label
    Public WithEvents _lblValue_1 As Label
    Public WithEvents _lblText_0 As Label
    Public WithEvents _lblValue_0 As Label
    Public WithEvents _Label5_0 As Label
    Public WithEvents Label1 As Label
    Public WithEvents Label2 As Label
    Public WithEvents _lstRange_0 As ListBox
    Public WithEvents _lstRange_1 As ListBox
    Public WithEvents Label3 As Label
    Public WithEvents Label4 As Label
    Public WithEvents Label7 As Label
    Public WithEvents Label8 As Label
    Public WithEvents Label9 As Label
    Public WithEvents lblTemperature As Label
    Public WithEvents lblPhase As Label
    Public WithEvents lblComments As Label
    Public WithEvents _lblValue_3 As Label
    Friend WithEvents grpIsothermDB As GroupBox
    Friend WithEvents _cmdCancelOK_0 As Button
    Friend WithEvents _cmdCancelOK_1 As Button
    Friend WithEvents cmdEditPolanyi As Button
    Friend WithEvents _cmdFind_0 As Button
    Friend WithEvents _cmdFind_1 As Button
    Friend WithEvents cmdSelect As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripDirty As ToolStripStatusLabel
    Friend WithEvents ToolStripStatus As ToolStripStatusLabel
    Friend WithEvents cmdCalculate As Button
    Friend WithEvents RadioButton3 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
#End Region
End Class