<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout
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
	Public WithEvents cmdLaunchWebSite As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents picIcon As System.Windows.Forms.PictureBox
	Public WithEvents lblSerialNumber As System.Windows.Forms.Label
	Public WithEvents lblUserName As System.Windows.Forms.Label
	Public WithEvents lblUserCompany As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_5 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_5 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_4 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_3 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_2 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_1 As System.Windows.Forms.Label
    Public WithEvents _lblWarning_0 As System.Windows.Forms.Label
    Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
    Public WithEvents _lbldesc_0 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_4 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_3 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_2 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_1 As System.Windows.Forms.Label
    Public WithEvents _lblVersionInfo_0 As System.Windows.Forms.Label
    Public WithEvents lblProgramName As System.Windows.Forms.Label
    Public WithEvents lblVersionInfo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblWarning As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lbldesc As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.cmdLaunchWebSite = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.lblSerialNumber = New System.Windows.Forms.Label()
        Me.lblUserName = New System.Windows.Forms.Label()
        Me.lblUserCompany = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_5 = New System.Windows.Forms.Label()
        Me._lblWarning_5 = New System.Windows.Forms.Label()
        Me._lblWarning_4 = New System.Windows.Forms.Label()
        Me._lblWarning_3 = New System.Windows.Forms.Label()
        Me._lblWarning_2 = New System.Windows.Forms.Label()
        Me._lblWarning_1 = New System.Windows.Forms.Label()
        Me._lblWarning_0 = New System.Windows.Forms.Label()
        Me._lbldesc_0 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_4 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_3 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_2 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_1 = New System.Windows.Forms.Label()
        Me._lblVersionInfo_0 = New System.Windows.Forms.Label()
        Me.lblProgramName = New System.Windows.Forms.Label()
        Me.lblVersionInfo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblWarning = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lbldesc = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblWarning, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.Line1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(650, 446)
        Me.ShapeContainer1.TabIndex = 22
        Me.ShapeContainer1.TabStop = False
        '
        'Line1
        '
        Me.Line1.BorderColor = System.Drawing.SystemColors.WindowText
        Me.Line1.BorderWidth = 2
        Me.Line1.Name = "Line1"
        Me.Line1.X1 = 4
        Me.Line1.X2 = 370
        Me.Line1.Y1 = 206
        Me.Line1.Y2 = 206
        '
        'cmdLaunchWebSite
        '
        Me.cmdLaunchWebSite.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLaunchWebSite.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLaunchWebSite.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdLaunchWebSite.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLaunchWebSite.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLaunchWebSite.Location = New System.Drawing.Point(280, 250)
        Me.cmdLaunchWebSite.Name = "cmdLaunchWebSite"
        Me.cmdLaunchWebSite.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLaunchWebSite.Size = New System.Drawing.Size(91, 23)
        Me.cmdLaunchWebSite.TabIndex = 21
        Me.cmdLaunchWebSite.TabStop = False
        Me.cmdLaunchWebSite.Text = "Go to web site"
        Me.cmdLaunchWebSite.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(280, 216)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(91, 23)
        Me.cmdOK.TabIndex = 8
        Me.cmdOK.Text = "&OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'picIcon
        '
        Me.picIcon.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.picIcon.Cursor = System.Windows.Forms.Cursors.Default
        Me.picIcon.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picIcon.ForeColor = System.Drawing.SystemColors.WindowText
        Me.picIcon.Image = CType(resources.GetObject("picIcon.Image"), System.Drawing.Image)
        Me.picIcon.Location = New System.Drawing.Point(10, 22)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picIcon.Size = New System.Drawing.Size(41, 37)
        Me.picIcon.TabIndex = 7
        Me.picIcon.TabStop = False
        '
        'lblSerialNumber
        '
        Me.lblSerialNumber.BackColor = System.Drawing.Color.Transparent
        Me.lblSerialNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSerialNumber.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSerialNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSerialNumber.Location = New System.Drawing.Point(3, 44)
        Me.lblSerialNumber.Name = "lblSerialNumber"
        Me.lblSerialNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSerialNumber.Size = New System.Drawing.Size(355, 15)
        Me.lblSerialNumber.TabIndex = 19
        Me.lblSerialNumber.Text = "{Z_SERIALNUMBER}"
        Me.lblSerialNumber.UseMnemonic = False
        '
        'lblUserName
        '
        Me.lblUserName.BackColor = System.Drawing.Color.Transparent
        Me.lblUserName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUserName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUserName.Location = New System.Drawing.Point(3, 0)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUserName.Size = New System.Drawing.Size(355, 15)
        Me.lblUserName.TabIndex = 18
        Me.lblUserName.Text = "{Z_USERNAME}"
        Me.lblUserName.UseMnemonic = False
        '
        'lblUserCompany
        '
        Me.lblUserCompany.BackColor = System.Drawing.Color.Transparent
        Me.lblUserCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUserCompany.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserCompany.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblUserCompany.Location = New System.Drawing.Point(3, 14)
        Me.lblUserCompany.Name = "lblUserCompany"
        Me.lblUserCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUserCompany.Size = New System.Drawing.Size(355, 15)
        Me.lblUserCompany.TabIndex = 17
        Me.lblUserCompany.Text = "{Z_USERCOMPANY}"
        Me.lblUserCompany.UseMnemonic = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label1.Location = New System.Drawing.Point(3, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(71, 15)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Serial Number:"
        '
        '_lblVersionInfo_5
        '
        Me._lblVersionInfo_5.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_5, CType(5, Short))
        Me._lblVersionInfo_5.Location = New System.Drawing.Point(84, 90)
        Me._lblVersionInfo_5.Name = "_lblVersionInfo_5"
        Me._lblVersionInfo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_5.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_5.TabIndex = 20
        Me._lblVersionInfo_5.Text = "(Build Code XX)"
        '
        '_lblWarning_5
        '
        Me._lblWarning_5.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_5, CType(5, Short))
        Me._lblWarning_5.Location = New System.Drawing.Point(4, 286)
        Me._lblWarning_5.Name = "_lblWarning_5"
        Me._lblWarning_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_5.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_5.TabIndex = 14
        Me._lblWarning_5.Text = "extent possible under law."
        '
        '_lblWarning_4
        '
        Me._lblWarning_4.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_4, CType(4, Short))
        Me._lblWarning_4.Location = New System.Drawing.Point(4, 272)
        Me._lblWarning_4.Name = "_lblWarning_4"
        Me._lblWarning_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_4.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_4.TabIndex = 13
        Me._lblWarning_4.Text = "penalties, and will be prosecuted to the maximum"
        '
        '_lblWarning_3
        '
        Me._lblWarning_3.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_3, CType(3, Short))
        Me._lblWarning_3.Location = New System.Drawing.Point(4, 258)
        Me._lblWarning_3.Name = "_lblWarning_3"
        Me._lblWarning_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_3.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_3.TabIndex = 12
        Me._lblWarning_3.Text = "portion of it, may result in severe civil and criminal"
        '
        '_lblWarning_2
        '
        Me._lblWarning_2.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_2, CType(2, Short))
        Me._lblWarning_2.Location = New System.Drawing.Point(4, 244)
        Me._lblWarning_2.Name = "_lblWarning_2"
        Me._lblWarning_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_2.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_2.TabIndex = 11
        Me._lblWarning_2.Text = "reproduction or distribution of this program, or any"
        '
        '_lblWarning_1
        '
        Me._lblWarning_1.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_1, CType(1, Short))
        Me._lblWarning_1.Location = New System.Drawing.Point(4, 230)
        Me._lblWarning_1.Name = "_lblWarning_1"
        Me._lblWarning_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_1.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_1.TabIndex = 10
        Me._lblWarning_1.Text = "copyright law and international treaties.  Unauthorized"
        '
        '_lblWarning_0
        '
        Me._lblWarning_0.BackColor = System.Drawing.Color.Transparent
        Me._lblWarning_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblWarning_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblWarning_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblWarning.SetIndex(Me._lblWarning_0, CType(0, Short))
        Me._lblWarning_0.Location = New System.Drawing.Point(4, 216)
        Me._lblWarning_0.Name = "_lblWarning_0"
        Me._lblWarning_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblWarning_0.Size = New System.Drawing.Size(265, 15)
        Me._lblWarning_0.TabIndex = 9
        Me._lblWarning_0.Text = "Warning: This computer program is protected by"
        '
        '_lbldesc_0
        '
        Me._lbldesc_0.BackColor = System.Drawing.Color.Transparent
        Me._lbldesc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lbldesc_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lbldesc_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbldesc.SetIndex(Me._lbldesc_0, CType(0, Short))
        Me._lbldesc_0.Location = New System.Drawing.Point(4, 108)
        Me._lbldesc_0.Name = "_lbldesc_0"
        Me._lbldesc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lbldesc_0.Size = New System.Drawing.Size(281, 15)
        Me._lbldesc_0.TabIndex = 6
        Me._lbldesc_0.Text = "This program is licensed to:"
        '
        '_lblVersionInfo_4
        '
        Me._lblVersionInfo_4.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_4, CType(4, Short))
        Me._lblVersionInfo_4.Location = New System.Drawing.Point(84, 76)
        Me._lblVersionInfo_4.Name = "_lblVersionInfo_4"
        Me._lblVersionInfo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_4.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_4.TabIndex = 5
        Me._lblVersionInfo_4.Text = "Houghton, Michigan"
        '
        '_lblVersionInfo_3
        '
        Me._lblVersionInfo_3.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_3, CType(3, Short))
        Me._lblVersionInfo_3.Location = New System.Drawing.Point(84, 62)
        Me._lblVersionInfo_3.Name = "_lblVersionInfo_3"
        Me._lblVersionInfo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_3.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_3.TabIndex = 4
        Me._lblVersionInfo_3.Text = "Michigan Technological University"
        '
        '_lblVersionInfo_2
        '
        Me._lblVersionInfo_2.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_2, CType(2, Short))
        Me._lblVersionInfo_2.Location = New System.Drawing.Point(84, 48)
        Me._lblVersionInfo_2.Name = "_lblVersionInfo_2"
        Me._lblVersionInfo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_2.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_2.TabIndex = 3
        Me._lblVersionInfo_2.Text = "{copyright info}"
        '
        '_lblVersionInfo_1
        '
        Me._lblVersionInfo_1.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_1, CType(1, Short))
        Me._lblVersionInfo_1.Location = New System.Drawing.Point(84, 34)
        Me._lblVersionInfo_1.Name = "_lblVersionInfo_1"
        Me._lblVersionInfo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_1.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_1.TabIndex = 2
        Me._lblVersionInfo_1.Text = "{expiration info}"
        '
        '_lblVersionInfo_0
        '
        Me._lblVersionInfo_0.BackColor = System.Drawing.Color.Transparent
        Me._lblVersionInfo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVersionInfo_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblVersionInfo_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblVersionInfo.SetIndex(Me._lblVersionInfo_0, CType(0, Short))
        Me._lblVersionInfo_0.Location = New System.Drawing.Point(84, 20)
        Me._lblVersionInfo_0.Name = "_lblVersionInfo_0"
        Me._lblVersionInfo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVersionInfo_0.Size = New System.Drawing.Size(281, 15)
        Me._lblVersionInfo_0.TabIndex = 1
        Me._lblVersionInfo_0.Text = "{version info}"
        '
        'lblProgramName
        '
        Me.lblProgramName.BackColor = System.Drawing.Color.Transparent
        Me.lblProgramName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProgramName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgramName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblProgramName.Location = New System.Drawing.Point(84, 6)
        Me.lblProgramName.Name = "lblProgramName"
        Me.lblProgramName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProgramName.Size = New System.Drawing.Size(281, 15)
        Me.lblProgramName.TabIndex = 0
        Me.lblProgramName.Text = "{AppName}"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.lblUserName)
        Me.Panel1.Controls.Add(Me.lblSerialNumber)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lblUserCompany)
        Me.Panel1.Location = New System.Drawing.Point(15, 135)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(356, 64)
        Me.Panel1.TabIndex = 23
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.CancelButton = Me.cmdLaunchWebSite
        Me.ClientSize = New System.Drawing.Size(650, 446)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdLaunchWebSite)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.picIcon)
        Me.Controls.Add(Me._lblVersionInfo_5)
        Me.Controls.Add(Me._lblWarning_5)
        Me.Controls.Add(Me._lblWarning_4)
        Me.Controls.Add(Me._lblWarning_3)
        Me.Controls.Add(Me._lblWarning_2)
        Me.Controls.Add(Me._lblWarning_1)
        Me.Controls.Add(Me._lblWarning_0)
        Me.Controls.Add(Me._lbldesc_0)
        Me.Controls.Add(Me._lblVersionInfo_4)
        Me.Controls.Add(Me._lblVersionInfo_3)
        Me.Controls.Add(Me._lblVersionInfo_2)
        Me.Controls.Add(Me._lblVersionInfo_1)
        Me.Controls.Add(Me._lblVersionInfo_0)
        Me.Controls.Add(Me.lblProgramName)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(52, 214)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "About"
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVersionInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblWarning, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lbldesc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
#End Region
End Class