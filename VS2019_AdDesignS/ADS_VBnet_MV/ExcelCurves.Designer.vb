<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmExcelCurves
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
    Public WithEvents _mnuFileItem_40 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFileItem_49 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFileItem_50 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFileItem_55 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFileItem_60 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFileItem_198 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFileItem_199 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuEditItem_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuEditItem_20 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    '   Public WithEvents f1book As VCIF1Lib.F1Book
    Public WithEvents f1bookDataGrid As DataGridView  'Replace f1book
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
        Me._mnuFileItem_40 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_49 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFileItem_50 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_55 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_60 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFileItem_198 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFileItem_199 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditItem_20 = New System.Windows.Forms.ToolStripMenuItem()
        Me.f1bookDataGrid = New System.Windows.Forms.DataGridView()
        Me.mnuEditItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuFileItem = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.MainMenu1.SuspendLayout()
        CType(Me.f1bookDataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.MainMenu1.Size = New System.Drawing.Size(796, 24)
        Me.MainMenu1.TabIndex = 2
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFileItem_40, Me._mnuFileItem_49, Me._mnuFileItem_50, Me._mnuFileItem_55, Me._mnuFileItem_60, Me._mnuFileItem_198, Me._mnuFileItem_199})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        '_mnuFileItem_40
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_40, CType(40, Short))
        Me._mnuFileItem_40.Name = "_mnuFileItem_40"
        Me._mnuFileItem_40.Size = New System.Drawing.Size(222, 22)
        Me._mnuFileItem_40.Text = "Save &As ..."
        '
        '_mnuFileItem_49
        '
        Me._mnuFileItem_49.Name = "_mnuFileItem_49"
        Me._mnuFileItem_49.Size = New System.Drawing.Size(219, 6)
        '
        '_mnuFileItem_50
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_50, CType(50, Short))
        Me._mnuFileItem_50.Name = "_mnuFileItem_50"
        Me._mnuFileItem_50.Size = New System.Drawing.Size(222, 22)
        Me._mnuFileItem_50.Text = "Page Setup ..."
        '
        '_mnuFileItem_55
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_55, CType(55, Short))
        Me._mnuFileItem_55.Name = "_mnuFileItem_55"
        Me._mnuFileItem_55.Size = New System.Drawing.Size(222, 22)
        Me._mnuFileItem_55.Text = "Printer Setup ..."
        '
        '_mnuFileItem_60
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_60, CType(60, Short))
        Me._mnuFileItem_60.Name = "_mnuFileItem_60"
        Me._mnuFileItem_60.Size = New System.Drawing.Size(222, 22)
        Me._mnuFileItem_60.Text = "&Print (Current Sheet Only) ..."
        '
        '_mnuFileItem_198
        '
        Me._mnuFileItem_198.Name = "_mnuFileItem_198"
        Me._mnuFileItem_198.Size = New System.Drawing.Size(219, 6)
        '
        '_mnuFileItem_199
        '
        Me.mnuFileItem.SetIndex(Me._mnuFileItem_199, CType(199, Short))
        Me._mnuFileItem_199.Name = "_mnuFileItem_199"
        Me._mnuFileItem_199.Size = New System.Drawing.Size(222, 22)
        Me._mnuFileItem_199.Text = "&Close"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuEditItem_10, Me._mnuEditItem_20})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(47, 20)
        Me.mnuEdit.Text = "&Copy"
        '
        '_mnuEditItem_10
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_10, CType(10, Short))
        Me._mnuEditItem_10.Name = "_mnuEditItem_10"
        Me._mnuEditItem_10.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me._mnuEditItem_10.Size = New System.Drawing.Size(274, 22)
        Me._mnuEditItem_10.Text = "&Copy Selection to Clipboard"
        '
        '_mnuEditItem_20
        '
        Me.mnuEditItem.SetIndex(Me._mnuEditItem_20, CType(20, Short))
        Me._mnuEditItem_20.Name = "_mnuEditItem_20"
        Me._mnuEditItem_20.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me._mnuEditItem_20.Size = New System.Drawing.Size(274, 22)
        Me._mnuEditItem_20.Text = "Copy &Entire Table to Clipboard"
        '
        'f1bookDataGrid
        '
        Me.f1bookDataGrid.AllowUserToOrderColumns = True
        Me.f1bookDataGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.f1bookDataGrid.ColumnHeadersHeight = 29
        Me.f1bookDataGrid.Location = New System.Drawing.Point(12, 37)
        Me.f1bookDataGrid.Name = "f1bookDataGrid"
        Me.f1bookDataGrid.RowHeadersWidth = 51
        Me.f1bookDataGrid.Size = New System.Drawing.Size(772, 511)
        Me.f1bookDataGrid.TabIndex = 0
        '
        'mnuEditItem
        '
        '
        'mnuFileItem
        '
        '
        'frmExcelCurves
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(796, 560)
        Me.Controls.Add(Me.f1bookDataGrid)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(116, 219)
        Me.MinimizeBox = False
        Me.Name = "frmExcelCurves"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "frmExcelCurves"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.f1bookDataGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuEditItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFileItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SaveFileDialog1 As SaveFileDialog
#End Region
End Class