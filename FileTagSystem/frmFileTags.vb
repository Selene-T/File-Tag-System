'Author:    Selene Tabacchini
'Pupose:    Catagorize and store tags for files in a database that can be searched through later
'Date:      2012-02-26

'Set options
Option Explicit On
Option Strict On

Public Class frmFileTags
    'Declare a couple public variables, one base on a custom class
    Public fileStatus As New clsFileData
    Public arlMD5 As New ArrayList
    Public fileList As New List(Of clsFileData)
    Public database As New clsDatabase
    Public blnIsMouseDown As Boolean = False
    Public intListBoxWidth As Int32 = 156

    Private Sub frmFileTags_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'When closing, close the about form, and if a database is open, close that too
        frmAbout.Close()
        If database.State = ConnectionState.Open Then database.Close()
    End Sub

    Private Sub LoadProcess(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Processing variables to centre the form on the primary monitor
        Dim intScreenWidth As Int32 = Screen.PrimaryScreen.WorkingArea.Width
        Dim intScreenHeight As Int32 = Screen.PrimaryScreen.WorkingArea.Height
        Me.Left = Convert.ToInt32(intScreenWidth / 2 - Me.Width / 2)
        Me.Top = Convert.ToInt32(intScreenHeight / 2 - Me.Height / 2)

        'Set the minumum size to the form doesn't become too squished when moved about, and centre its location
        Me.MinimumSize = New System.Drawing.Size(480, 320)
        'Set the start size
        Me.Size = New System.Drawing.Size(800, 600)
        'Allow drag-drop to the form
        Me.AllowDrop = True
        'Add event handler when draging something into the form, and dropping it
        AddHandler Me.DragEnter, AddressOf fileDragEnter
        AddHandler Me.DragDrop, AddressOf fileDragDrop
        'Set the border for the label that displays the actual file name to none
        'I set the border to on by default so that it is easy to see and move on the form
        lblFilenameDisplay.BorderStyle = BorderStyle.None

        'Run sub to resize the form's objects
        Call setObjectsEnables(False)
        Call FormResize()
    End Sub

    Private Sub fileDragEnter(ByVal sender As Object, ByVal e As DragEventArgs)
        'Checks if information is being dragged
        If e.Data.GetDataPresent(DataFormats.FileDrop) And lstFiles.Enabled Then
            'If there is a file being dragged in, and the listbox is enabled (ready to accept), copy it
            e.Effect = DragDropEffects.Copy
        Else
            'If not, make sure the data area is cleared
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub fileDragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'Checks if there is data stored form the drag when information is dropped
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'Declare a string array for all files being dragged in
            Dim strFilePath As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            'Run through every entry in the array
            For Each strFilesDropped As String In strFilePath
                'Adds file to database
                Call addFile(strFilesDropped)
            Next
            Exit Sub
        End If
    End Sub

    Private Sub addFile(ByRef strFilePath As String)
        'Checks and confirms that the file exists
        If System.IO.File.Exists(strFilePath) Then
            'Declare a variable to hold md5 as called by the function
            With fileStatus
                'Want to check that the file isn't already added, because then this would
                'just clear the tags, but instead want to load the tags
                'Store the MD5 for quick calling
                .MD5 = GetMD5(strFilePath)
                If .MD5 = "ERROR" Then Exit Sub
                If Not .MD5 = Nothing Then
                    database.Command("SELECT * FROM FileTags WHERE (MD5 = '" & .MD5 & "')")
                    'If the select statement returned anything, then it's already added
                    If Not database.IsDataEmpty Then
                        Dim tblData As DataTable = database.Data
                        txtTags.Text = tblData(0)("Tags").ToString
                    Else
                        txtTags.Text = String.Empty
                    End If
                End If
                'Split the filepath into an array series of string with the data between the \'s
                'Then use that again to grab how far down the final array string is, and use that for the filename
                .FileName = Split(strFilePath, "\")(UBound(Split(strFilePath, "\")))
                'The full filepath, remove the fileName
                .FilePath = Strings.Left(strFilePath, strFilePath.Length - .FileName.Length)
                'Change the label to display the filename to show it was added
                lblFilenameDisplay.Text = .FileName
                'Add the file to the database
                Call DBAddFile()
            End With
        End If
    End Sub

    Private Function GetMD5(ByVal strFileLoc As String) As String
        'Call DoEvents whenever this function is accessed, this process can take a while on large files
        Application.DoEvents()
        On Error GoTo FileIsOpen
        'Opens the file in read-only mode
        Using strOpenedFile As New System.IO.FileStream(strFileLoc, IO.FileMode.Open, IO.FileAccess.Read)
            'Declare the service for getting the MD5
            Using md5Calc As New System.Security.Cryptography.MD5CryptoServiceProvider
                'Get the MD5 for the file
                Dim hashFile() As Byte = md5Calc.ComputeHash(strOpenedFile)
                'Use a nifty function found on evilzone.org that converts the btye into a hexidecimal string
                Return ByteArrayToString(hashFile)
            End Using
        End Using
        Exit Function
FileIsOpen:
        Return "ERROR"
    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        'Credit goes to user ande from evilzon.org, but they were also asking who made this and who to credit

        Dim sb As New System.Text.StringBuilder(arrInput.Length * 2)
        For i As Integer = 0 To arrInput.Length - 1
            sb.Append(arrInput(i).ToString("X2"))
        Next
        Return sb.ToString().ToLower
    End Function

    Private Sub FormResize() Handles Me.Resize
        'When the window is minimized, it's set for a tiny size, so ignore that
        If Me.WindowState = FormWindowState.Minimized Then Exit Sub


        'Use the forms' border and size of other buttons to place and size them in the proper locations
        'Menu is to be the width of the form, keeping it's own height; sitting in the top left across the whole form
        mnuMain.Location = New System.Drawing.Point(0, 0)
        mnuMain.Size = New System.Drawing.Size(ClientSize.Width, mnuMain.Height)
        'Listbox won't change width, not here at least; height will be the remainder of the form below the menu
        'It will sit just under the menu and to the left of the form
        lstFiles.Location = New System.Drawing.Point(0, mnuMain.Height)
        lstFiles.Height = ClientSize.Height - mnuMain.Height
        'Checks to make sure the listbox's width isn't too large, and resizes it if needed
        'This occurs if the listbox was stretched wide when maximized, then switched to normal view
        If lstFiles.Width > ClientSize.Width - 255 Then lstFiles.Width = ClientSize.Width - 255
        'The pair of labels will not change height, but be the remaining width of the form after the listbox
        'They will sit just to the right of the listbox and right under the menu.  One label to the right of the other
        lblFilename.Location = New System.Drawing.Point(lstFiles.Width, mnuMain.Height)
        lblFilenameDisplay.Size = New System.Drawing.Size(ClientSize.Width - lstFiles.Width - lblFilename.Width, lblFilename.Height)
        lblFilenameDisplay.Location = New System.Drawing.Point(lstFiles.Width + lblFilename.Width, mnuMain.Height)
        'The textbox will be the width left after the listbox, the combined width of both labels, height the remainder
        'of the form's height after the menu and labels, but then leaving room for the buttons below it.
        'It will sit to the right edge of the listbox and right under the labels.
        txtTags.Location = New System.Drawing.Point(lstFiles.Width, lblFilename.Height + mnuMain.Height)
        txtTags.Size = New System.Drawing.Size(ClientSize.Width - lstFiles.Width, lstFiles.Height - lblFilename.Height - btnSave.Height)
        'Processing variable for button width, their width will be a third the width of the textbox, to fit three of
        'them under it. Calculation has to be done now and not declared at the beginning. Their height won't change
        Dim intButtonWidth As Int32 = Convert.ToInt32(Math.Floor(txtTags.Width / 3))
        'They will be located below the textbox, and anchored to the bottom-right of the form. This is because with 3
        'buttons, there will be a little whitespace, putting that whitespace between them and the listbox as it is not
        'so noticable there.
        btnSave.Location = New System.Drawing.Point(ClientSize.Width - intButtonWidth * 3, ClientSize.Height - btnSave.Height)
        btnSave.Size = New System.Drawing.Size(intButtonWidth, btnSave.Height)
        btnRemove.Location = New System.Drawing.Point(ClientSize.Width - btnSave.Width * 2, ClientSize.Height - btnRemove.Height)
        btnRemove.Size = New System.Drawing.Size(intButtonWidth, btnSave.Height)
        btnExit.Location = New System.Drawing.Point(ClientSize.Width - intButtonWidth, ClientSize.Height - btnExit.Height)
        btnExit.Size = New System.Drawing.Size(intButtonWidth, btnSave.Height)
    End Sub

    Private Sub lstFiles_DoubleClick(sender As Object, e As System.EventArgs) Handles lstFiles.DoubleClick
        On Error GoTo OpenError
        'Handles the error in the case that windows cannot open the file
        System.Diagnostics.Process.Start(fileStatus.FilePath & fileStatus.FileName)
        'Exit before the error handling message
        Exit Sub
OpenError:
        'Displays message to user
        MsgBox("No associated application for this file.")
    End Sub

    Private Sub resizeFormLeftSide(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstFiles.MouseDown, lstFiles.MouseMove
        'Set blnIsMouseDown to true if the left mouse button is down and are within 10 measures from the right boarder of lstFiles
        'This activates only if near the edge, but stays on as long as the mouse button is held
        If e.Button = Windows.Forms.MouseButtons.Left And e.X > lstFiles.Width - 10 Then blnIsMouseDown = True
        'As long as the mouse button down is active...
        If blnIsMouseDown = True Then
            'Make sure the neither box is too small
            If e.X > 150 And e.X < ClientSize.Width - 255 Then
                'Resize the listbox and then the whole form accordingly
                lstFiles.Width = e.X
                Call FormResize()
            End If
        End If
    End Sub

    Private Sub resizeFormRightSide(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtTags.MouseDown, txtTags.MouseMove
        'Set blnIsMouseDown to true if the left mouse button is down and are within 10 measures from the left boarder of txtTags
        'This activates only if near the edge, but stays on as long as the mouse button is held
        If e.Button = Windows.Forms.MouseButtons.Left And e.X < 10 Then blnIsMouseDown = True
        'As long as the mouse button down is active...
        If blnIsMouseDown = True Then
            'Make sure the neither box is too small
            If lstFiles.Width + e.X > 150 And lstFiles.Width + e.X < ClientSize.Width - 255 Then
                'Resize the listbox and then the whole form accordingly
                lstFiles.Width = lstFiles.Width + e.X
                Call FormResize()
            End If
        End If
    End Sub

    Private Sub confirmMouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtTags.MouseUp, lstFiles.MouseUp
        'When the mouse is released, set it to false
        blnIsMouseDown = False
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Call DBAddFile()
    End Sub

    Private Sub exitApplication(sender As System.Object, e As System.EventArgs) Handles btnExit.Click, mnuExit.Click
        'Closes the form/application
        Me.Close()
    End Sub

    Private Sub mnuAddFile_Click(sender As System.Object, e As System.EventArgs) Handles mnuAddFile.Click
        'Set the multiselect option to true because the user can open more than 1 file at a time
        ofdFile.Multiselect = True
        ofdFile.Filter = "All Files (*.*)|*.*"
        ofdFile.FilterIndex = 1
        'Opens the open file dialog and confirms a selection was made
        If ofdFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Count through each of the files selected
            For Each strFileSelected In ofdFile.FileNames
                'Adds file to database
                Call addFile(strFileSelected)
            Next
        End If
    End Sub

    Private Sub mnuAddFolder_Click(sender As System.Object, e As System.EventArgs) Handles mnuAddFolder.Click
        'Opens the folder browser dialog and confirms a selection was made
        If fbdFolder.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            'Goes through each file in the selected folder
            For Each strAllFilesInFolder As String In System.IO.Directory.GetFiles(fbdFolder.SelectedPath)
                'Adds each the database
                Call addFile(strAllFilesInFolder)
            Next
        End If
    End Sub

    Private Sub setObjectsEnables(ByVal blnState As Boolean)
        'Sets the enabled value of every object (except the menus)
        For Each ctrl As System.Windows.Forms.Control In Me.Controls
            If Not ctrl.Name = "btnExit" And Not ctrl.Name = "mnuMain" Then ctrl.Enabled = blnState
        Next
        mnuAddFile.Enabled = blnState
        mnuAddFolder.Enabled = blnState
        mnuDatabaseClose.Enabled = blnState
        mnuQuery.Enabled = blnState

        'If things were enabled, then focus the tags textbox
        If blnState Then
            txtTags.Focus()
        Else
            lblFilenameDisplay.Text = String.Empty
            txtTags.Text = String.Empty
        End If

        'Clear the fields on the form
        lblFilenameDisplay.Text = String.Empty
        txtTags.Text = String.Empty
        lstFiles.Items.Clear()
    End Sub

    Private Sub mnuDatabaseNew_Click(sender As System.Object, e As System.EventArgs) Handles mnuDatabaseNew.Click
        'Creates a save file entry and confirms there was a positive selection
        sfdDatabase.Filter = "Database Files (*.accdb)|*.accdb|All Files (*.*)|*.*"
        sfdDatabase.FilterIndex = 1
        If sfdDatabase.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Closes the database if it is already open
            If database.State = ConnectionState.Open Then
                database.Close()
                MsgBox("Closing old DB")
            End If
            'Sends a delete command to the file if saving over another file, or just ignores it
            If System.IO.File.Exists(sfdDatabase.FileName) Then System.IO.File.Delete(sfdDatabase.FileName)
            database.CreateDB(sfdDatabase.FileName)
            database.Database = sfdDatabase.FileName

            database.Command("CREATE TABLE FileTags(" &
                              "MD5 char(32) primary key," &
                              "File varchar(255) not null," &
                              "Path varchar(255) not null," &
                              "Tags varchar(255))")
            lstFiles.Items.Clear()
            Call setObjectsEnables(True)
        End If
    End Sub

    Private Sub mnuDatabaseOpen_Click(sender As System.Object, e As System.EventArgs) Handles mnuDatabaseOpen.Click
        'Set the multiselect option to false as we only want to open 1 database at a time
        ofdFile.Multiselect = False
        ofdFile.Filter = "Database Files (*.accdb)|*.accdb|All Files (*.*)|*.*"
        ofdFile.FilterIndex = 1
        'Opens the open file dialog and confirms that a selection was made
        If ofdFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            If database.State = ConnectionState.Open Then database.Close()
            database.Database = ofdFile.FileName
            lstFiles.Items.Clear()
            Call setObjectsEnables(True)
            'Calls to list all the database entries
            database.Command("SELECT * FROM FileTags")
            Call FillListBox()
        End If
    End Sub

    Private Sub mnuDatabaseClose_Click(sender As System.Object, e As System.EventArgs) Handles mnuDatabaseClose.Click
        'Check if the database is open before clearing it, then clear the other fields and disable the controls
        If database.State = ConnectionState.Open Then database.Close()
        lstFiles.Items.Clear()
        Call setObjectsEnables(False)
    End Sub

    Private Sub mnuAbout_Click(sender As System.Object, e As System.EventArgs) Handles mnuAbout.Click
        'Shows and focuses the About form
        frmAbout.Show()
        frmAbout.Focus()
    End Sub

    Public Sub displayMessage(ByVal strMessage As String)
        'Displays message to user
        MessageBox.Show(strMessage, "File Tags: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        'Removes the currently selected item from the list and database

        'If -1 (no selection) then exit the sub
        If lstFiles.SelectedIndex < 0 Then Exit Sub
        With fileStatus
            'Makes sure there is an MD5 to search by
            If Not .MD5 = Nothing Then
                'Runs a select to see if there is an entry with matching MD5 (primary key)
                database.Command("SELECT * FROM FileTags WHERE (MD5 = '" & .MD5 & "')")
                'If there is no returned data, then the entry is not there
                If Not database.IsDataEmpty Then
                    'But if it is there, delete it
                    database.Command("DELETE FROM FileTags WHERE (MD5 = '" & .MD5 & "')")
                    'Remove the entry from the listbox and clear the fields
                    lblFilenameDisplay.Text = String.Empty
                    txtTags.Text = String.Empty
                    arlMD5.RemoveAt(lstFiles.SelectedIndex)
                    lstFiles.Items.RemoveAt(lstFiles.SelectedIndex)
                End If
            End If
        End With
    End Sub

    Private Sub DBAddFile()
        'Adds a file or files to the database
        With fileStatus
            If Not .MD5 = Nothing Then
                'Cleck to make sure the file isn't already there
                database.Command("SELECT * FROM FileTags WHERE (MD5 = '" & .MD5 & "')")
                If database.IsDataEmpty Then
                    'If it isn't add it
                    database.Command("INSERT INTO FileTags VALUES ('" & .MD5 & "', '" & .FileName & "', '" & .FilePath & "', '" & .Tags & "')")
                Else
                    'If it is, update it
                    database.Command("UPDATE FileTags SET File = '" & .FileName & "' WHERE (MD5 = '" & .MD5 & "')")
                    database.Command("UPDATE FileTags SET Path = '" & .FilePath & "' WHERE (MD5 = '" & .MD5 & "')")
                    database.Command("UPDATE FileTags SET Tags = '" & .Tags & "' WHERE (MD5 = '" & .MD5 & "')")
                End If
            End If
        End With
    End Sub
    Private Sub txtTags_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtTags.TextChanged
        'When the tags are changed, store them in a nice clean sorted format to the tags string, which will be used to the store to the database
        fileStatus.Tags = SortString(System.Text.RegularExpressions.Regex.Replace(txtTags.Text, "^\s+|[^\w ]|\s+$", ""))
    End Sub

    Private Sub mnuQueryAll_Click(sender As System.Object, e As System.EventArgs) Handles mnuQueryAll.Click
        'Calls for all files in the database
        database.Command("SELECT * FROM FileTags")
        Call FillListBox()
    End Sub

    Private Sub FillListBox()
        'Fills the listbox with all entries from the previous search command
        If Not database.IsDataEmpty Then
            'Only fill it is it isn't an empty search

            'Declare variables
            Dim tblData As DataTable = database.Data
            Dim fileTemp As New clsFileData

            'Clear lists and labels
            fileList.Clear()
            arlMD5.Clear()
            lstFiles.Items.Clear()
            lblFilenameDisplay.Text = String.Empty
            txtTags.Text = String.Empty

            'For each entry, add it to the lists
            For i As Int32 = 0 To tblData.Rows.Count - 1
                arlMD5.Add(tblData(i)("MD5").ToString)
                lstFiles.Items.Add(tblData(i)("File").ToString)
            Next

            'Set the listbox selection to the first entry
            lstFiles.SelectedIndex = 0
        Else
            'If the search was empty, just clear everything
            fileList.Clear()
            arlMD5.Clear()
            lstFiles.Items.Clear()
            lblFilenameDisplay.Text = String.Empty
            txtTags.Text = String.Empty
        End If
    End Sub

    Private Sub lstFiles_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstFiles.SelectedIndexChanged
        'When the selection on the listbox is changed, update the other controls to show its filename and tags

        'Exit the sub if there is no selection
        If lstFiles.SelectedIndex < 0 Then Exit Sub
        With fileStatus
            'Checks if the file entry is in the database
            .MD5 = arlMD5(lstFiles.SelectedIndex).ToString
            database.Command("SELECT * FROM FileTags WHERE (MD5 = '" & .MD5 & "')")
            If database.IsDataEmpty Then
                'Warn if it is a missing entry
                MsgBox("Entry Missing")
            Else
                'Set the current file status and displays to match the selected file
                Dim tblData As DataTable = database.Data
                .FileName = tblData(0)("File").ToString
                .FilePath = tblData(0)("Path").ToString
                .Tags = SortString(tblData(0)("Tags").ToString)
                lblFilenameDisplay.Text = .FileName
                txtTags.Text = .Tags
            End If
        End With

        'Checks to make sure the file exists
        If Not System.IO.File.Exists(fileStatus.FilePath & fileStatus.FileName) Then
            With fileStatus
                Dim blnDidFindNewFile As Boolean = False
                For Each strAllFilesInFolder As String In System.IO.Directory.GetFiles(.FilePath)
                    'If it doesn't, scan every file in the folder for a matching MD5
                    If GetMD5(strAllFilesInFolder) = .MD5 Then
                        'If found, update and notify the user
                        Dim intListSpotHold As Int32 = lstFiles.SelectedIndex
                        MsgBox("Found a entry where the filename was changed, auto-correcting this entry.")
                        .FileName = Split(strAllFilesInFolder, "\")(UBound(Split(strAllFilesInFolder, "\")))
                        .FilePath = Strings.Left(strAllFilesInFolder, strAllFilesInFolder.Length - .FileName.Length)
                        database.Command("UPDATE FileTags SET File = '" & .FileName & "' WHERE (MD5 = '" & .MD5 & "')")
                        database.Command("UPDATE FileTags SET Path = '" & .FilePath & "' WHERE (MD5 = '" & .MD5 & "')")
                        lstFiles.Items.RemoveAt(intListSpotHold)
                        lstFiles.Items.Insert(intListSpotHold, .FileName)
                        lstFiles.SelectedIndex = intListSpotHold
                        blnDidFindNewFile = True
                        Exit For
                    End If
                Next
                'If not, remove notify the user and suggest a removal of the entry
                If blnDidFindNewFile = False Then MsgBox("File for current entry not found, it may be wise to remove it.")
            End With
        End If
    End Sub

    Private Sub mnuSearchTags_Click(sender As System.Object, e As System.EventArgs) Handles mnuSearchTags.Click
        Dim strQuery As String = InputBox("Enter tag to search for: (Separate tags with a space)", "Search Database", "")
        If strQuery = String.Empty Then Exit Sub
        'Regex to remove all whitespace from the begining and end of the string, and remove anything
        'that is not Alphanumeric or a space
        strQuery = System.Text.RegularExpressions.Regex.Replace(strQuery, "^\s+|[^\w ]|\s+$", "")
        'Sorts all tags searched for by ABC, this is to query the database will be easier
        strQuery = SortString(strQuery)
        'Regex replaces all whitespace with a % now, used in the LIKE operator in a query
        strQuery = System.Text.RegularExpressions.Regex.Replace(strQuery, "\s+", "%")
        'Calls the search
        database.Command("SELECT * FROM FileTags WHERE Tags LIKE '%" & strQuery & "%'")
        'And parse what was found
        Call FillListBox()
    End Sub

    Private Function SortString(ByRef strOriginal As String) As String
        'Sorts a string's words by alphabetical order

        'Declare string for the return
        Dim strSorted As String = String.Empty
        'Declare and fill an array. Splits the original string by each word for each array entry
        Dim arrWords As String() = Split(strOriginal, " ")
        'Sort the array
        System.Array.Sort(arrWords)
        For Each strWord As String In arrWords
            'For each array word, add it to the output string
            If strSorted = String.Empty Then
                'For the first word, we don't need a space before the word
                strSorted += strWord
            Else
                strSorted += Space(1) & strWord
            End If
        Next
        Return strSorted
    End Function

    Private Sub mnuSearchFiles_Click(sender As System.Object, e As System.EventArgs) Handles mnuSearchFiles.Click
        Dim strQuery As String = InputBox("Enter part of the filename to search for:" & vbCrLf & "Example: 'png' or 'exe'", "Search Database", "")
        If strQuery = String.Empty Then Exit Sub
        'Regex to remove all whitespace from the string, and remove anything that is not Alphanumeric
        strQuery = System.Text.RegularExpressions.Regex.Replace(strQuery, "[^\w]", "")
        'Calls the search
        database.Command("SELECT * FROM FileTags WHERE File LIKE '%" & strQuery & "%'")
        'And parse what was found
        Call FillListBox()
    End Sub
End Class