<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(main))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuQuit = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.newfilepathBrowse = New System.Windows.Forms.Button()
        Me.oldfilepathBrowse = New System.Windows.Forms.Button()
        Me.newfilepathText = New System.Windows.Forms.TextBox()
        Me.oldfilepathText = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.runconfigBtn = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog3 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(503, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuQuit})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'menuQuit
        '
        Me.menuQuit.Name = "menuQuit"
        Me.menuQuit.Size = New System.Drawing.Size(200, 22)
        Me.menuQuit.Text = "Quit                          (Esc)"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HelpToolStripMenuItem1, Me.ToolStripMenuItem1, Me.menuAbout})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(24, 20)
        Me.HelpToolStripMenuItem.Text = "?"
        '
        'HelpToolStripMenuItem1
        '
        Me.HelpToolStripMenuItem1.Name = "HelpToolStripMenuItem1"
        Me.HelpToolStripMenuItem1.Size = New System.Drawing.Size(211, 22)
        Me.HelpToolStripMenuItem1.Text = "Help                              (F2)"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(208, 6)
        '
        'menuAbout
        '
        Me.menuAbout.Name = "menuAbout"
        Me.menuAbout.Size = New System.Drawing.Size(211, 22)
        Me.menuAbout.Text = "About...                         (F1)"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.newfilepathBrowse)
        Me.GroupBox1.Controls.Add(Me.oldfilepathBrowse)
        Me.GroupBox1.Controls.Add(Me.newfilepathText)
        Me.GroupBox1.Controls.Add(Me.oldfilepathText)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(22, 37)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(473, 89)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "MNSure Script Installation"
        '
        'newfilepathBrowse
        '
        Me.newfilepathBrowse.Location = New System.Drawing.Point(411, 51)
        Me.newfilepathBrowse.Name = "newfilepathBrowse"
        Me.newfilepathBrowse.Size = New System.Drawing.Size(56, 20)
        Me.newfilepathBrowse.TabIndex = 5
        Me.newfilepathBrowse.Text = "Browse"
        Me.newfilepathBrowse.UseVisualStyleBackColor = True
        '
        'oldfilepathBrowse
        '
        Me.oldfilepathBrowse.Location = New System.Drawing.Point(411, 19)
        Me.oldfilepathBrowse.Name = "oldfilepathBrowse"
        Me.oldfilepathBrowse.Size = New System.Drawing.Size(56, 20)
        Me.oldfilepathBrowse.TabIndex = 4
        Me.oldfilepathBrowse.Text = "Browse"
        Me.oldfilepathBrowse.UseVisualStyleBackColor = True
        '
        'newfilepathText
        '
        Me.newfilepathText.Location = New System.Drawing.Point(110, 51)
        Me.newfilepathText.Name = "newfilepathText"
        Me.newfilepathText.Size = New System.Drawing.Size(295, 20)
        Me.newfilepathText.TabIndex = 3
        '
        'oldfilepathText
        '
        Me.oldfilepathText.Location = New System.Drawing.Point(110, 19)
        Me.oldfilepathText.Name = "oldfilepathText"
        Me.oldfilepathText.Size = New System.Drawing.Size(295, 20)
        Me.oldfilepathText.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(105, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Installation Directory:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Current Directory:"
        '
        'runconfigBtn
        '
        Me.runconfigBtn.Location = New System.Drawing.Point(383, 132)
        Me.runconfigBtn.Name = "runconfigBtn"
        Me.runconfigBtn.Size = New System.Drawing.Size(112, 29)
        Me.runconfigBtn.TabIndex = 2
        Me.runconfigBtn.Text = "Run Configuration"
        Me.runconfigBtn.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog1
        '
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(503, 167)
        Me.Controls.Add(Me.runconfigBtn)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "main"
        Me.Text = "MNSure Scripts Configuration"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuQuit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents newfilepathText As System.Windows.Forms.TextBox
    Friend WithEvents oldfilepathText As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents runconfigBtn As System.Windows.Forms.Button
    Friend WithEvents newfilepathBrowse As System.Windows.Forms.Button
    Friend WithEvents oldfilepathBrowse As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialog2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialog3 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents HelpToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents FileSystemWatcher1 As System.IO.FileSystemWatcher

End Class
