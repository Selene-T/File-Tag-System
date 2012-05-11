<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFileTags
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
        Me.lblFilename = New System.Windows.Forms.Label()
        Me.lblFilenameDisplay = New System.Windows.Forms.Label()
        Me.txtTags = New System.Windows.Forms.TextBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.mnuMain = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDatabase = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDatabaseNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDatabaseOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDatabaseClose = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuQuery = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuQueryAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSearchTags = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.lstFiles = New System.Windows.Forms.ListBox()
        Me.ofdFile = New System.Windows.Forms.OpenFileDialog()
        Me.fbdFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.sfdDatabase = New System.Windows.Forms.SaveFileDialog()
        Me.mnuSearchFiles = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFilename
        '
        Me.lblFilename.AutoSize = True
        Me.lblFilename.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilename.Location = New System.Drawing.Point(12, 52)
        Me.lblFilename.Name = "lblFilename"
        Me.lblFilename.Size = New System.Drawing.Size(77, 18)
        Me.lblFilename.TabIndex = 1
        Me.lblFilename.Text = "Filename:"
        '
        'lblFilenameDisplay
        '
        Me.lblFilenameDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFilenameDisplay.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilenameDisplay.Location = New System.Drawing.Point(95, 52)
        Me.lblFilenameDisplay.Name = "lblFilenameDisplay"
        Me.lblFilenameDisplay.Size = New System.Drawing.Size(77, 18)
        Me.lblFilenameDisplay.TabIndex = 2
        '
        'txtTags
        '
        Me.txtTags.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTags.Location = New System.Drawing.Point(15, 73)
        Me.txtTags.MaxLength = 255
        Me.txtTags.Multiline = True
        Me.txtTags.Name = "txtTags"
        Me.txtTags.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTags.Size = New System.Drawing.Size(156, 64)
        Me.txtTags.TabIndex = 3
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(16, 143)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(156, 37)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemove.Location = New System.Drawing.Point(16, 186)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(156, 37)
        Me.btnRemove.TabIndex = 5
        Me.btnRemove.Text = "&Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(16, 229)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(156, 37)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'mnuMain
        '
        Me.mnuMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.mnuMain.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuQuery, Me.HelpToolStripMenuItem})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(792, 24)
        Me.mnuMain.TabIndex = 7
        Me.mnuMain.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDatabase, Me.mnuAdd, Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(35, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuDatabase
        '
        Me.mnuDatabase.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDatabaseNew, Me.mnuDatabaseOpen, Me.mnuDatabaseClose})
        Me.mnuDatabase.Name = "mnuDatabase"
        Me.mnuDatabase.Size = New System.Drawing.Size(120, 22)
        Me.mnuDatabase.Text = "Database"
        '
        'mnuDatabaseNew
        '
        Me.mnuDatabaseNew.Name = "mnuDatabaseNew"
        Me.mnuDatabaseNew.Size = New System.Drawing.Size(100, 22)
        Me.mnuDatabaseNew.Text = "&New"
        '
        'mnuDatabaseOpen
        '
        Me.mnuDatabaseOpen.Name = "mnuDatabaseOpen"
        Me.mnuDatabaseOpen.Size = New System.Drawing.Size(100, 22)
        Me.mnuDatabaseOpen.Text = "Open"
        '
        'mnuDatabaseClose
        '
        Me.mnuDatabaseClose.Name = "mnuDatabaseClose"
        Me.mnuDatabaseClose.Size = New System.Drawing.Size(100, 22)
        Me.mnuDatabaseClose.Text = "&Close"
        '
        'mnuAdd
        '
        Me.mnuAdd.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddFile, Me.mnuAddFolder})
        Me.mnuAdd.Name = "mnuAdd"
        Me.mnuAdd.Size = New System.Drawing.Size(120, 22)
        Me.mnuAdd.Text = "A&dd"
        '
        'mnuAddFile
        '
        Me.mnuAddFile.Name = "mnuAddFile"
        Me.mnuAddFile.Size = New System.Drawing.Size(104, 22)
        Me.mnuAddFile.Text = "Files"
        '
        'mnuAddFolder
        '
        Me.mnuAddFolder.Name = "mnuAddFolder"
        Me.mnuAddFolder.Size = New System.Drawing.Size(104, 22)
        Me.mnuAddFolder.Text = "Folder"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(120, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'mnuQuery
        '
        Me.mnuQuery.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuQueryAll, Me.mnuSearchTags, Me.mnuSearchFiles})
        Me.mnuQuery.Name = "mnuQuery"
        Me.mnuQuery.Size = New System.Drawing.Size(49, 20)
        Me.mnuQuery.Text = "&Query"
        '
        'mnuQueryAll
        '
        Me.mnuQueryAll.Name = "mnuQueryAll"
        Me.mnuQueryAll.Size = New System.Drawing.Size(152, 22)
        Me.mnuQueryAll.Text = "&All"
        '
        'mnuSearchTags
        '
        Me.mnuSearchTags.Name = "mnuSearchTags"
        Me.mnuSearchTags.Size = New System.Drawing.Size(152, 22)
        Me.mnuSearchTags.Text = "Search &Tags"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(103, 22)
        Me.mnuAbout.Text = "A&bout"
        '
        'lstFiles
        '
        Me.lstFiles.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFiles.FormattingEnabled = True
        Me.lstFiles.IntegralHeight = False
        Me.lstFiles.ItemHeight = 18
        Me.lstFiles.Location = New System.Drawing.Point(12, 27)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.Size = New System.Drawing.Size(156, 22)
        Me.lstFiles.TabIndex = 0
        '
        'ofdFile
        '
        Me.ofdFile.RestoreDirectory = True
        '
        'fbdFolder
        '
        Me.fbdFolder.ShowNewFolderButton = False
        '
        'mnuSearchFiles
        '
        Me.mnuSearchFiles.Name = "mnuSearchFiles"
        Me.mnuSearchFiles.Size = New System.Drawing.Size(152, 22)
        Me.mnuSearchFiles.Text = "Search Fi&les"
        '
        'frmFileTags
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 573)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtTags)
        Me.Controls.Add(Me.lblFilenameDisplay)
        Me.Controls.Add(Me.lblFilename)
        Me.Controls.Add(Me.lstFiles)
        Me.Controls.Add(Me.mnuMain)
        Me.MainMenuStrip = Me.mnuMain
        Me.Name = "frmFileTags"
        Me.Text = "File Tags"
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFilename As System.Windows.Forms.Label
    Friend WithEvents lblFilenameDisplay As System.Windows.Forms.Label
    Friend WithEvents txtTags As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents mnuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents lstFiles As System.Windows.Forms.ListBox
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDatabase As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDatabaseNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDatabaseOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDatabaseClose As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ofdFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents fbdFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents sfdDatabase As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuQuery As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuQueryAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSearchTags As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSearchFiles As System.Windows.Forms.ToolStripMenuItem

End Class
