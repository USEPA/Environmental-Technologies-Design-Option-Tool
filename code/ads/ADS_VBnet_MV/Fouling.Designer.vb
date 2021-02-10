<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFouling
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
    Public chkUse(10) As AxThreed.AxSSCheck
    Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents cboType As System.Windows.Forms.ComboBox
    Public WithEvents fraWater As AxThreed.AxSSFrame
    Public WithEvents _cboCorrel_9 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_8 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_7 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_6 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_5 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_4 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_3 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_2 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_1 As System.Windows.Forms.ComboBox
    Public WithEvents _cboCorrel_0 As System.Windows.Forms.ComboBox
    Public WithEvents cmdEditCompo As AxThreed.AxSSCommand
    Public WithEvents _chkUse_0 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_1 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_2 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_3 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_4 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_5 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_6 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_7 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_8 As AxThreed.AxSSCheck
    Public WithEvents _chkUse_9 As AxThreed.AxSSCheck
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblName_9 As System.Windows.Forms.Label
    Public WithEvents _lblName_8 As System.Windows.Forms.Label
    Public WithEvents _lblName_7 As System.Windows.Forms.Label
    Public WithEvents _lblName_6 As System.Windows.Forms.Label
    Public WithEvents _lblName_5 As System.Windows.Forms.Label
    Public WithEvents _lblName_4 As System.Windows.Forms.Label
    Public WithEvents _lblName_3 As System.Windows.Forms.Label
    Public WithEvents _lblName_2 As System.Windows.Forms.Label
    Public WithEvents _lblName_1 As System.Windows.Forms.Label
    Public WithEvents _lblName_0 As System.Windows.Forms.Label
    Public WithEvents fraCompo As AxThreed.AxSSFrame
    Public WithEvents _cmdCancelOK_1 As AxThreed.AxSSCommand
    Public WithEvents _cmdCancelOK_0 As AxThreed.AxSSCommand
    Public WithEvents cboCorrel As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
    '   Public WithEvents chkUse As SSCheckArray
    '   Public WithEvents cmdCancelOK As SSCommandArray
    Public WithEvents lblName As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFouling))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.cboCorrel = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(Me.components)
        Me._cboCorrel_9 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_8 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_7 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_6 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_5 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_4 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_3 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_2 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_1 = New System.Windows.Forms.ComboBox()
        Me._cboCorrel_0 = New System.Windows.Forms.ComboBox()
        Me.lblName = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me._lblName_9 = New System.Windows.Forms.Label()
        Me._lblName_8 = New System.Windows.Forms.Label()
        Me._lblName_7 = New System.Windows.Forms.Label()
        Me._lblName_6 = New System.Windows.Forms.Label()
        Me._lblName_5 = New System.Windows.Forms.Label()
        Me._lblName_4 = New System.Windows.Forms.Label()
        Me._lblName_3 = New System.Windows.Forms.Label()
        Me._lblName_2 = New System.Windows.Forms.Label()
        Me._lblName_1 = New System.Windows.Forms.Label()
        Me._lblName_0 = New System.Windows.Forms.Label()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fraWater = New AxThreed.AxSSFrame()
        Me.fraCompo = New AxThreed.AxSSFrame()
        Me._chkUse_0 = New AxThreed.AxSSCheck()
        Me._chkUse_1 = New AxThreed.AxSSCheck()
        Me._chkUse_2 = New AxThreed.AxSSCheck()
        Me._chkUse_3 = New AxThreed.AxSSCheck()
        Me._chkUse_4 = New AxThreed.AxSSCheck()
        Me._chkUse_5 = New AxThreed.AxSSCheck()
        Me._chkUse_6 = New AxThreed.AxSSCheck()
        Me._chkUse_7 = New AxThreed.AxSSCheck()
        Me._chkUse_8 = New AxThreed.AxSSCheck()
        Me._chkUse_9 = New AxThreed.AxSSCheck()
        Me.cmdEditCompo = New AxThreed.AxSSCommand()
        Me._cmdCancelOK_1 = New AxThreed.AxSSCommand()
        Me._cmdCancelOK_0 = New AxThreed.AxSSCommand()
        Me.cmdEdit = New AxThreed.AxSSCommand()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboCorrel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraWater, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraWater.SuspendLayout()
        CType(Me.fraCompo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraCompo.SuspendLayout()
        CType(Me._chkUse_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._chkUse_9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdEditCompo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(143, 446)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(165, 22)
        Me.Command4.TabIndex = 40
        Me.Command4.Text = "Print Screen"
        Me.ToolTip1.SetToolTip(Me.Command4, "Click here to print current screen to selected printer")
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Location = New System.Drawing.Point(641, 374)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(89, 57)
        Me.Picture1.TabIndex = 41
        Me.Picture1.TabStop = False
        Me.Picture1.Visible = False
        '
        'cboCorrel
        '
        '
        '_cboCorrel_9
        '
        Me._cboCorrel_9.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_9, CType(9, Short))
        Me._cboCorrel_9.Location = New System.Drawing.Point(350, 249)
        Me._cboCorrel_9.Name = "_cboCorrel_9"
        Me._cboCorrel_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_9.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_9.TabIndex = 16
        '
        '_cboCorrel_8
        '
        Me._cboCorrel_8.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_8, CType(8, Short))
        Me._cboCorrel_8.Location = New System.Drawing.Point(350, 225)
        Me._cboCorrel_8.Name = "_cboCorrel_8"
        Me._cboCorrel_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_8.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_8.TabIndex = 15
        '
        '_cboCorrel_7
        '
        Me._cboCorrel_7.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_7, CType(7, Short))
        Me._cboCorrel_7.Location = New System.Drawing.Point(350, 201)
        Me._cboCorrel_7.Name = "_cboCorrel_7"
        Me._cboCorrel_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_7.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_7.TabIndex = 14
        '
        '_cboCorrel_6
        '
        Me._cboCorrel_6.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_6, CType(6, Short))
        Me._cboCorrel_6.Location = New System.Drawing.Point(350, 177)
        Me._cboCorrel_6.Name = "_cboCorrel_6"
        Me._cboCorrel_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_6.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_6.TabIndex = 13
        '
        '_cboCorrel_5
        '
        Me._cboCorrel_5.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_5, CType(5, Short))
        Me._cboCorrel_5.Location = New System.Drawing.Point(350, 153)
        Me._cboCorrel_5.Name = "_cboCorrel_5"
        Me._cboCorrel_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_5.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_5.TabIndex = 12
        '
        '_cboCorrel_4
        '
        Me._cboCorrel_4.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_4, CType(4, Short))
        Me._cboCorrel_4.Location = New System.Drawing.Point(350, 129)
        Me._cboCorrel_4.Name = "_cboCorrel_4"
        Me._cboCorrel_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_4.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_4.TabIndex = 11
        '
        '_cboCorrel_3
        '
        Me._cboCorrel_3.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_3, CType(3, Short))
        Me._cboCorrel_3.Location = New System.Drawing.Point(350, 105)
        Me._cboCorrel_3.Name = "_cboCorrel_3"
        Me._cboCorrel_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_3.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_3.TabIndex = 10
        '
        '_cboCorrel_2
        '
        Me._cboCorrel_2.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_2, CType(2, Short))
        Me._cboCorrel_2.Location = New System.Drawing.Point(350, 81)
        Me._cboCorrel_2.Name = "_cboCorrel_2"
        Me._cboCorrel_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_2.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_2.TabIndex = 9
        '
        '_cboCorrel_1
        '
        Me._cboCorrel_1.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_1, CType(1, Short))
        Me._cboCorrel_1.Location = New System.Drawing.Point(350, 57)
        Me._cboCorrel_1.Name = "_cboCorrel_1"
        Me._cboCorrel_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_1.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_1.TabIndex = 8
        '
        '_cboCorrel_0
        '
        Me._cboCorrel_0.BackColor = System.Drawing.SystemColors.Window
        Me._cboCorrel_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cboCorrel_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me._cboCorrel_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cboCorrel_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCorrel.SetIndex(Me._cboCorrel_0, CType(0, Short))
        Me._cboCorrel_0.Location = New System.Drawing.Point(350, 33)
        Me._cboCorrel_0.Name = "_cboCorrel_0"
        Me._cboCorrel_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cboCorrel_0.Size = New System.Drawing.Size(166, 24)
        Me._cboCorrel_0.TabIndex = 7
        '
        '_lblName_9
        '
        Me._lblName_9.BackColor = System.Drawing.Color.Transparent
        Me._lblName_9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_9, CType(9, Short))
        Me._lblName_9.Location = New System.Drawing.Point(146, 249)
        Me._lblName_9.Name = "_lblName_9"
        Me._lblName_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_9.Size = New System.Drawing.Size(206, 24)
        Me._lblName_9.TabIndex = 36
        Me._lblName_9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_8
        '
        Me._lblName_8.BackColor = System.Drawing.Color.Transparent
        Me._lblName_8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_8, CType(8, Short))
        Me._lblName_8.Location = New System.Drawing.Point(146, 225)
        Me._lblName_8.Name = "_lblName_8"
        Me._lblName_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_8.Size = New System.Drawing.Size(206, 24)
        Me._lblName_8.TabIndex = 35
        Me._lblName_8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_7
        '
        Me._lblName_7.BackColor = System.Drawing.Color.Transparent
        Me._lblName_7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_7, CType(7, Short))
        Me._lblName_7.Location = New System.Drawing.Point(146, 201)
        Me._lblName_7.Name = "_lblName_7"
        Me._lblName_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_7.Size = New System.Drawing.Size(206, 24)
        Me._lblName_7.TabIndex = 34
        Me._lblName_7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_6
        '
        Me._lblName_6.BackColor = System.Drawing.Color.Transparent
        Me._lblName_6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_6, CType(6, Short))
        Me._lblName_6.Location = New System.Drawing.Point(146, 177)
        Me._lblName_6.Name = "_lblName_6"
        Me._lblName_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_6.Size = New System.Drawing.Size(206, 24)
        Me._lblName_6.TabIndex = 33
        Me._lblName_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_5
        '
        Me._lblName_5.BackColor = System.Drawing.Color.Transparent
        Me._lblName_5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_5, CType(5, Short))
        Me._lblName_5.Location = New System.Drawing.Point(146, 153)
        Me._lblName_5.Name = "_lblName_5"
        Me._lblName_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_5.Size = New System.Drawing.Size(206, 24)
        Me._lblName_5.TabIndex = 32
        Me._lblName_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_4
        '
        Me._lblName_4.BackColor = System.Drawing.Color.Transparent
        Me._lblName_4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_4, CType(4, Short))
        Me._lblName_4.Location = New System.Drawing.Point(146, 129)
        Me._lblName_4.Name = "_lblName_4"
        Me._lblName_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_4.Size = New System.Drawing.Size(206, 24)
        Me._lblName_4.TabIndex = 31
        Me._lblName_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_3
        '
        Me._lblName_3.BackColor = System.Drawing.Color.Transparent
        Me._lblName_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_3, CType(3, Short))
        Me._lblName_3.Location = New System.Drawing.Point(146, 105)
        Me._lblName_3.Name = "_lblName_3"
        Me._lblName_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_3.Size = New System.Drawing.Size(206, 24)
        Me._lblName_3.TabIndex = 30
        Me._lblName_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_2
        '
        Me._lblName_2.BackColor = System.Drawing.Color.Transparent
        Me._lblName_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_2, CType(2, Short))
        Me._lblName_2.Location = New System.Drawing.Point(146, 81)
        Me._lblName_2.Name = "_lblName_2"
        Me._lblName_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_2.Size = New System.Drawing.Size(206, 24)
        Me._lblName_2.TabIndex = 29
        Me._lblName_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_1
        '
        Me._lblName_1.BackColor = System.Drawing.Color.Transparent
        Me._lblName_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_1, CType(1, Short))
        Me._lblName_1.Location = New System.Drawing.Point(146, 57)
        Me._lblName_1.Name = "_lblName_1"
        Me._lblName_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_1.Size = New System.Drawing.Size(206, 24)
        Me._lblName_1.TabIndex = 28
        Me._lblName_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_lblName_0
        '
        Me._lblName_0.BackColor = System.Drawing.Color.Transparent
        Me._lblName_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblName_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblName_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblName_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblName.SetIndex(Me._lblName_0, CType(0, Short))
        Me._lblName_0.Location = New System.Drawing.Point(146, 33)
        Me._lblName_0.Name = "_lblName_0"
        Me._lblName_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblName_0.Size = New System.Drawing.Size(206, 24)
        Me._lblName_0.TabIndex = 27
        Me._lblName_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cboType
        '
        Me.cboType.BackColor = System.Drawing.SystemColors.Window
        Me.cboType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboType.Location = New System.Drawing.Point(62, 24)
        Me.cboType.Name = "cboType"
        Me.cboType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboType.Size = New System.Drawing.Size(464, 24)
        Me.cboType.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label5.Location = New System.Drawing.Point(6, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(57, 17)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Apply"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(250, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(153, 17)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Type of correlation used"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(66, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(145, 17)
        Me.Label1.TabIndex = 37
        Me.Label1.Text = "Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'fraWater
        '
        Me.fraWater.Controls.Add(Me.cboType)
        Me.fraWater.Location = New System.Drawing.Point(12, 12)
        Me.fraWater.Name = "fraWater"
        Me.fraWater.OcxState = CType(resources.GetObject("fraWater.OcxState"), System.Windows.Forms.AxHost.State)
        Me.fraWater.Size = New System.Drawing.Size(556, 89)
        Me.fraWater.TabIndex = 0
        '
        'fraCompo
        '
        Me.fraCompo.Controls.Add(Me._cboCorrel_0)
        Me.fraCompo.Controls.Add(Me._lblName_0)
        Me.fraCompo.Controls.Add(Me._chkUse_0)
        Me.fraCompo.Controls.Add(Me._cboCorrel_1)
        Me.fraCompo.Controls.Add(Me._lblName_1)
        Me.fraCompo.Controls.Add(Me._chkUse_1)
        Me.fraCompo.Controls.Add(Me._cboCorrel_2)
        Me.fraCompo.Controls.Add(Me._lblName_2)
        Me.fraCompo.Controls.Add(Me._chkUse_2)
        Me.fraCompo.Controls.Add(Me._cboCorrel_3)
        Me.fraCompo.Controls.Add(Me._lblName_3)
        Me.fraCompo.Controls.Add(Me._chkUse_3)
        Me.fraCompo.Controls.Add(Me._cboCorrel_4)
        Me.fraCompo.Controls.Add(Me._lblName_4)
        Me.fraCompo.Controls.Add(Me._chkUse_4)
        Me.fraCompo.Controls.Add(Me._cboCorrel_5)
        Me.fraCompo.Controls.Add(Me._lblName_5)
        Me.fraCompo.Controls.Add(Me._chkUse_5)
        Me.fraCompo.Controls.Add(Me._cboCorrel_6)
        Me.fraCompo.Controls.Add(Me._lblName_6)
        Me.fraCompo.Controls.Add(Me._chkUse_6)
        Me.fraCompo.Controls.Add(Me._cboCorrel_7)
        Me.fraCompo.Controls.Add(Me._lblName_7)
        Me.fraCompo.Controls.Add(Me._chkUse_7)
        Me.fraCompo.Controls.Add(Me._cboCorrel_8)
        Me.fraCompo.Controls.Add(Me._lblName_8)
        Me.fraCompo.Controls.Add(Me._chkUse_8)
        Me.fraCompo.Controls.Add(Me._cboCorrel_9)
        Me.fraCompo.Controls.Add(Me._lblName_9)
        Me.fraCompo.Controls.Add(Me._chkUse_9)
        Me.fraCompo.Location = New System.Drawing.Point(12, 114)
        Me.fraCompo.Name = "fraCompo"
        Me.fraCompo.OcxState = CType(resources.GetObject("fraCompo.OcxState"), System.Windows.Forms.AxHost.State)
        Me.fraCompo.Size = New System.Drawing.Size(552, 317)
        Me.fraCompo.TabIndex = 1
        '
        '_chkUse_0
        '
        Me._chkUse_0.Location = New System.Drawing.Point(22, 33)
        Me._chkUse_0.Name = "_chkUse_0"
        Me._chkUse_0.OcxState = CType(resources.GetObject("_chkUse_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_0.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_0.TabIndex = 17
        chkUse(0) = _chkUse_0
        '
        '_chkUse_1
        '
        Me._chkUse_1.Location = New System.Drawing.Point(22, 58)
        Me._chkUse_1.Name = "_chkUse_1"
        Me._chkUse_1.OcxState = CType(resources.GetObject("_chkUse_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_1.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_1.TabIndex = 18
        chkUse(1) = _chkUse_1
        '
        '_chkUse_2
        '
        Me._chkUse_2.Location = New System.Drawing.Point(22, 82)
        Me._chkUse_2.Name = "_chkUse_2"
        Me._chkUse_2.OcxState = CType(resources.GetObject("_chkUse_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_2.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_2.TabIndex = 19
        chkUse(2) = _chkUse_2
        '
        '_chkUse_3
        '
        Me._chkUse_3.Location = New System.Drawing.Point(22, 106)
        Me._chkUse_3.Name = "_chkUse_3"
        Me._chkUse_3.OcxState = CType(resources.GetObject("_chkUse_3.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_3.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_3.TabIndex = 20
        chkUse(3) = _chkUse_3
        '
        '_chkUse_4
        '
        Me._chkUse_4.Location = New System.Drawing.Point(22, 132)
        Me._chkUse_4.Name = "_chkUse_4"
        Me._chkUse_4.OcxState = CType(resources.GetObject("_chkUse_4.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_4.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_4.TabIndex = 21
        chkUse(4) = _chkUse_4
        '
        '_chkUse_5
        '
        Me._chkUse_5.Location = New System.Drawing.Point(22, 156)
        Me._chkUse_5.Name = "_chkUse_5"
        Me._chkUse_5.OcxState = CType(resources.GetObject("_chkUse_5.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_5.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_5.TabIndex = 22
        chkUse(5) = _chkUse_5


        '
        '_chkUse_6
        '
        Me._chkUse_6.Location = New System.Drawing.Point(22, 180)
        Me._chkUse_6.Name = "_chkUse_6"
        Me._chkUse_6.OcxState = CType(resources.GetObject("_chkUse_6.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_6.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_6.TabIndex = 23
        chkUse(6) = _chkUse_6
        '
        '_chkUse_7
        '
        Me._chkUse_7.Location = New System.Drawing.Point(22, 203)
        Me._chkUse_7.Name = "_chkUse_7"
        Me._chkUse_7.OcxState = CType(resources.GetObject("_chkUse_7.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_7.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_7.TabIndex = 24
        chkUse(7) = _chkUse_7
        '
        '_chkUse_8
        '
        Me._chkUse_8.Location = New System.Drawing.Point(22, 228)
        Me._chkUse_8.Name = "_chkUse_8"
        Me._chkUse_8.OcxState = CType(resources.GetObject("_chkUse_8.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_8.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_8.TabIndex = 25
        chkUse(8) = _chkUse_8
        '
        '_chkUse_9
        '
        Me._chkUse_9.Location = New System.Drawing.Point(22, 252)
        Me._chkUse_9.Name = "_chkUse_9"
        Me._chkUse_9.OcxState = CType(resources.GetObject("_chkUse_9.OcxState"), System.Windows.Forms.AxHost.State)
        Me._chkUse_9.Size = New System.Drawing.Size(100, 25)
        Me._chkUse_9.TabIndex = 26
        chkUse(9) = _chkUse_9
        '
        'cmdEditCompo
        '
        Me.cmdEditCompo.Location = New System.Drawing.Point(115, 393)
        Me.cmdEditCompo.Name = "cmdEditCompo"
        Me.cmdEditCompo.OcxState = CType(resources.GetObject("cmdEditCompo.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdEditCompo.Size = New System.Drawing.Size(363, 24)
        Me.cmdEditCompo.TabIndex = 43
        '
        '_cmdCancelOK_1
        '
        Me._cmdCancelOK_1.Location = New System.Drawing.Point(45, 474)
        Me._cmdCancelOK_1.Name = "_cmdCancelOK_1"
        Me._cmdCancelOK_1.OcxState = CType(resources.GetObject("_cmdCancelOK_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_1.Size = New System.Drawing.Size(100, 50)
        Me._cmdCancelOK_1.TabIndex = 2
        Me._cmdCancelOK_1.TabStop = False
        '
        '_cmdCancelOK_0
        '
        Me._cmdCancelOK_0.Location = New System.Drawing.Point(294, 474)
        Me._cmdCancelOK_0.Name = "_cmdCancelOK_0"
        Me._cmdCancelOK_0.OcxState = CType(resources.GetObject("_cmdCancelOK_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._cmdCancelOK_0.Size = New System.Drawing.Size(100, 50)
        Me._cmdCancelOK_0.TabIndex = 3
        Me._cmdCancelOK_0.TabStop = False
        '
        'cmdEdit
        '
        Me.cmdEdit.Location = New System.Drawing.Point(115, 66)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.OcxState = CType(resources.GetObject("cmdEdit.OcxState"), System.Windows.Forms.AxHost.State)
        Me.cmdEdit.Size = New System.Drawing.Size(363, 24)
        Me.cmdEdit.TabIndex = 42
        '
        'frmFouling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(628, 558)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdEditCompo)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.fraWater)
        Me.Controls.Add(Me.fraCompo)
        Me.Controls.Add(Me._cmdCancelOK_1)
        Me.Controls.Add(Me._cmdCancelOK_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(167, 115)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFouling"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Fouling of GAC"
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboCorrel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraWater, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraWater.ResumeLayout(False)
        CType(Me.fraCompo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraCompo.ResumeLayout(False)
        CType(Me._chkUse_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._chkUse_9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdEditCompo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdCancelOK_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._cmdCancelOK_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdEdit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents cmdEdit As AxThreed.AxSSCommand

#End Region
End Class