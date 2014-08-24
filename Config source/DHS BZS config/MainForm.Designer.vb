<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class scripts_config_form
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.county_selection = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.access_DB_check = New System.Windows.Forms.CheckBox()
        Me.EDMS_check = New System.Windows.Forms.CheckBox()
        Me.EDMS_choice = New System.Windows.Forms.TextBox()
        Me.county_address_line_01 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.county_address_line_02 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.new_file_path = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.old_file_path = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.run_configuration_button = New System.Windows.Forms.Button()
        Me.intake_dates_check = New System.Windows.Forms.CheckBox()
        Me.Tab_Control_Main_Form = New System.Windows.Forms.TabControl()
        Me.basic_settings_tab = New System.Windows.Forms.TabPage()
        Me.advanced_script_mods_tab = New System.Windows.Forms.TabPage()
        Me.move_verifs_needed_check = New System.Windows.Forms.CheckBox()
        Me.advanced_file_path_mods_tab = New System.Windows.Forms.TabPage()
        Me.Update_Files_Label = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Tab_Control_Main_Form.SuspendLayout()
        Me.basic_settings_tab.SuspendLayout()
        Me.advanced_script_mods_tab.SuspendLayout()
        Me.advanced_file_path_mods_tab.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AccessibleDescription = "Menubar"
        Me.MenuStrip1.AccessibleName = "Menubar"
        Me.MenuStrip1.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.MenuStrip1.Location = New System.Drawing.Point(2, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(517, 24)
        Me.MenuStrip1.TabIndex = 7
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.AboutToolStripMenuItem.Text = "About..."
        '
        'county_selection
        '
        Me.county_selection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.county_selection.FormattingEnabled = True
        Me.county_selection.Items.AddRange(New Object() {"01 - Aitkin", "02 - Anoka", "03 - Becker", "04 - Beltrami", "05 - Benton", "06 - Big Stone", "07 - Blue Earth", "08 - Brown", "09 - Carlton", "10 - Carver", "11 - Cass", "12 - Chippewa", "13 - Chisago", "14 - Clay", "15 - Clearwater", "16 - Cook", "17 - Cottonwood", "18 - Crow Wing", "19 - Dakota", "20 - Dodge", "21 - Douglas", "22 - Faribault", "23 - Fillmore", "24 - Freeborn", "25 - Goodhue", "26 - Grant", "27 - Hennepin", "28 - Houston", "29 - Hubbard", "30 - Isanti", "31 - Itasca", "32 - Jackson", "33 - Kanabec", "34 - Kandiyohi", "35 - Kittson", "36 - Koochiching", "37 - Lac Qui Parle", "38 - Lake", "39 - Lake of the Woods", "40 - LeSueur", "41 - Lincoln", "42 - Lyon", "43 - Mcleod", "44 - Mahnomen", "45 - Marshall", "46 - Martin", "47 - Meeker", "48 - Mille Lacs", "49 - Morrison", "50 - Mower", "51 - Murray", "52 - Nicollet", "53 - Nobles", "54 - Norman", "55 - Olmsted", "56 - Otter Tail", "57 - Pennington", "58 - Pine", "59 - Pipestone", "60 - Polk", "61 - Pope", "62 - Ramsey", "63 - Red Lake", "64 - Redwood", "65 - Renville", "66 - Rice", "67 - Rock", "68 - Roseau", "69 - St. Louis", "70 - Scott", "71 - Sherburne", "72 - Sibley", "73 - Stearns", "74 - Steele", "75 - Stevens", "76 - Swift", "77 - Todd", "78 - Traverse", "79 - Wabasha", "80 - Wadena", "81 - Waseca", "82 - Washington", "83 - Watonwan", "84 - Wilkin", "85 - Winona", "86 - Wright", "87 - Yellow Medicine"})
        Me.county_selection.Location = New System.Drawing.Point(66, 19)
        Me.county_selection.Name = "county_selection"
        Me.county_selection.Size = New System.Drawing.Size(215, 21)
        Me.county_selection.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "County:"
        '
        'access_DB_check
        '
        Me.access_DB_check.AutoSize = True
        Me.access_DB_check.Enabled = False
        Me.access_DB_check.Location = New System.Drawing.Point(3, 109)
        Me.access_DB_check.Name = "access_DB_check"
        Me.access_DB_check.Size = New System.Drawing.Size(303, 17)
        Me.access_DB_check.TabIndex = 1
        Me.access_DB_check.Text = "Check here to collect statistics using an Access Database."
        Me.access_DB_check.UseVisualStyleBackColor = True
        '
        'EDMS_check
        '
        Me.EDMS_check.AutoSize = True
        Me.EDMS_check.Location = New System.Drawing.Point(3, 132)
        Me.EDMS_check.Name = "EDMS_check"
        Me.EDMS_check.Size = New System.Drawing.Size(298, 17)
        Me.EDMS_check.TabIndex = 2
        Me.EDMS_check.Text = "Check here if you use an EDMS, and enter its name here:"
        Me.EDMS_check.UseVisualStyleBackColor = True
        '
        'EDMS_choice
        '
        Me.EDMS_choice.Location = New System.Drawing.Point(307, 130)
        Me.EDMS_choice.Name = "EDMS_choice"
        Me.EDMS_choice.Size = New System.Drawing.Size(156, 20)
        Me.EDMS_choice.TabIndex = 3
        Me.EDMS_choice.Text = "ex: Compass Forms"
        '
        'county_address_line_01
        '
        Me.county_address_line_01.Location = New System.Drawing.Point(134, 46)
        Me.county_address_line_01.Name = "county_address_line_01"
        Me.county_address_line_01.Size = New System.Drawing.Size(199, 20)
        Me.county_address_line_01.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "County address line 1:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "County address line 2:"
        '
        'county_address_line_02
        '
        Me.county_address_line_02.Location = New System.Drawing.Point(134, 72)
        Me.county_address_line_02.Name = "county_address_line_02"
        Me.county_address_line_02.Size = New System.Drawing.Size(199, 20)
        Me.county_address_line_02.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "New file path:"
        '
        'new_file_path
        '
        Me.new_file_path.Location = New System.Drawing.Point(86, 32)
        Me.new_file_path.Name = "new_file_path"
        Me.new_file_path.Size = New System.Drawing.Size(409, 20)
        Me.new_file_path.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(135, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Old file path (script default):"
        '
        'old_file_path
        '
        Me.old_file_path.Location = New System.Drawing.Point(149, 6)
        Me.old_file_path.Name = "old_file_path"
        Me.old_file_path.Size = New System.Drawing.Size(346, 20)
        Me.old_file_path.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.county_selection)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.county_address_line_01)
        Me.GroupBox2.Controls.Add(Me.county_address_line_02)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(346, 100)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "County information"
        '
        'run_configuration_button
        '
        Me.run_configuration_button.Location = New System.Drawing.Point(402, 218)
        Me.run_configuration_button.Name = "run_configuration_button"
        Me.run_configuration_button.Size = New System.Drawing.Size(109, 30)
        Me.run_configuration_button.TabIndex = 6
        Me.run_configuration_button.Text = "Run Configuration"
        Me.run_configuration_button.UseVisualStyleBackColor = True
        '
        'intake_dates_check
        '
        Me.intake_dates_check.AutoSize = True
        Me.intake_dates_check.Checked = True
        Me.intake_dates_check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.intake_dates_check.Location = New System.Drawing.Point(6, 6)
        Me.intake_dates_check.Name = "intake_dates_check"
        Me.intake_dates_check.Size = New System.Drawing.Size(488, 17)
        Me.intake_dates_check.TabIndex = 4
        Me.intake_dates_check.Text = "Check here to have the ""closed progs"" and ""denied progs"" scripts case note info o" & _
    "n intake dates."
        Me.intake_dates_check.UseVisualStyleBackColor = True
        '
        'Tab_Control_Main_Form
        '
        Me.Tab_Control_Main_Form.Controls.Add(Me.basic_settings_tab)
        Me.Tab_Control_Main_Form.Controls.Add(Me.advanced_script_mods_tab)
        Me.Tab_Control_Main_Form.Controls.Add(Me.advanced_file_path_mods_tab)
        Me.Tab_Control_Main_Form.Location = New System.Drawing.Point(6, 27)
        Me.Tab_Control_Main_Form.Name = "Tab_Control_Main_Form"
        Me.Tab_Control_Main_Form.SelectedIndex = 0
        Me.Tab_Control_Main_Form.Size = New System.Drawing.Size(509, 185)
        Me.Tab_Control_Main_Form.TabIndex = 8
        '
        'basic_settings_tab
        '
        Me.basic_settings_tab.Controls.Add(Me.GroupBox2)
        Me.basic_settings_tab.Controls.Add(Me.access_DB_check)
        Me.basic_settings_tab.Controls.Add(Me.EDMS_check)
        Me.basic_settings_tab.Controls.Add(Me.EDMS_choice)
        Me.basic_settings_tab.Location = New System.Drawing.Point(4, 22)
        Me.basic_settings_tab.Name = "basic_settings_tab"
        Me.basic_settings_tab.Size = New System.Drawing.Size(501, 159)
        Me.basic_settings_tab.TabIndex = 2
        Me.basic_settings_tab.Text = "Basic settings"
        Me.basic_settings_tab.UseVisualStyleBackColor = True
        '
        'advanced_script_mods_tab
        '
        Me.advanced_script_mods_tab.Controls.Add(Me.move_verifs_needed_check)
        Me.advanced_script_mods_tab.Controls.Add(Me.intake_dates_check)
        Me.advanced_script_mods_tab.Location = New System.Drawing.Point(4, 22)
        Me.advanced_script_mods_tab.Name = "advanced_script_mods_tab"
        Me.advanced_script_mods_tab.Padding = New System.Windows.Forms.Padding(3)
        Me.advanced_script_mods_tab.Size = New System.Drawing.Size(501, 159)
        Me.advanced_script_mods_tab.TabIndex = 0
        Me.advanced_script_mods_tab.Text = "Advanced script mods"
        Me.advanced_script_mods_tab.UseVisualStyleBackColor = True
        '
        'move_verifs_needed_check
        '
        Me.move_verifs_needed_check.AutoSize = True
        Me.move_verifs_needed_check.Location = New System.Drawing.Point(6, 29)
        Me.move_verifs_needed_check.Name = "move_verifs_needed_check"
        Me.move_verifs_needed_check.Size = New System.Drawing.Size(408, 17)
        Me.move_verifs_needed_check.TabIndex = 5
        Me.move_verifs_needed_check.Text = "Check here to move the ""verifs needed"" section to the top of the CAF case note."
        Me.move_verifs_needed_check.UseVisualStyleBackColor = True
        '
        'advanced_file_path_mods_tab
        '
        Me.advanced_file_path_mods_tab.Controls.Add(Me.Label4)
        Me.advanced_file_path_mods_tab.Controls.Add(Me.old_file_path)
        Me.advanced_file_path_mods_tab.Controls.Add(Me.new_file_path)
        Me.advanced_file_path_mods_tab.Controls.Add(Me.Label5)
        Me.advanced_file_path_mods_tab.Location = New System.Drawing.Point(4, 22)
        Me.advanced_file_path_mods_tab.Name = "advanced_file_path_mods_tab"
        Me.advanced_file_path_mods_tab.Padding = New System.Windows.Forms.Padding(3)
        Me.advanced_file_path_mods_tab.Size = New System.Drawing.Size(501, 159)
        Me.advanced_file_path_mods_tab.TabIndex = 1
        Me.advanced_file_path_mods_tab.Text = "Advanced file path mods"
        Me.advanced_file_path_mods_tab.UseVisualStyleBackColor = True
        '
        'Update_Files_Label
        '
        Me.Update_Files_Label.AutoSize = True
        Me.Update_Files_Label.Location = New System.Drawing.Point(254, 227)
        Me.Update_Files_Label.Name = "Update_Files_Label"
        Me.Update_Files_Label.Size = New System.Drawing.Size(142, 13)
        Me.Update_Files_Label.TabIndex = 9
        Me.Update_Files_Label.Text = "Updating files, please wait! :)"
        Me.Update_Files_Label.Visible = False
        '
        'scripts_config_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(519, 254)
        Me.Controls.Add(Me.Update_Files_Label)
        Me.Controls.Add(Me.Tab_Control_Main_Form)
        Me.Controls.Add(Me.run_configuration_button)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "scripts_config_form"
        Me.Padding = New System.Windows.Forms.Padding(2, 0, 0, 0)
        Me.Text = "BlueZone Scripts Configuration"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Tab_Control_Main_Form.ResumeLayout(False)
        Me.basic_settings_tab.ResumeLayout(False)
        Me.basic_settings_tab.PerformLayout()
        Me.advanced_script_mods_tab.ResumeLayout(False)
        Me.advanced_script_mods_tab.PerformLayout()
        Me.advanced_file_path_mods_tab.ResumeLayout(False)
        Me.advanced_file_path_mods_tab.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents county_selection As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents access_DB_check As System.Windows.Forms.CheckBox
    Friend WithEvents EDMS_check As System.Windows.Forms.CheckBox
    Friend WithEvents EDMS_choice As System.Windows.Forms.TextBox
    Friend WithEvents county_address_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents county_address_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents new_file_path As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents old_file_path As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents run_configuration_button As System.Windows.Forms.Button
    Friend WithEvents intake_dates_check As System.Windows.Forms.CheckBox
    Friend WithEvents Tab_Control_Main_Form As System.Windows.Forms.TabControl
    Friend WithEvents advanced_script_mods_tab As System.Windows.Forms.TabPage
    Friend WithEvents advanced_file_path_mods_tab As System.Windows.Forms.TabPage
    Friend WithEvents basic_settings_tab As System.Windows.Forms.TabPage
    Friend WithEvents move_verifs_needed_check As System.Windows.Forms.CheckBox
    Friend WithEvents Update_Files_Label As System.Windows.Forms.Label

End Class
