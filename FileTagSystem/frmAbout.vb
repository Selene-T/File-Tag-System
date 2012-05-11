'Author:    Selene Tabacchini
'Pupose:    The about page with some basic information
'Date:      2012-02-26

'Set options
Option Explicit On
Option Strict On

Public Class frmAbout

    Private Sub frmAbout_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Dim frmAbout As System.Windows.Forms.Form = CType(sender, System.Windows.Forms.Form)
        'Processing variables to centre the form on the main form
        'Dim intScreenWidth As Int32 = Screen.PrimaryScreen.WorkingArea.Width
        'Dim intScreenHeight As Int32 = Screen.PrimaryScreen.WorkingArea.Height
        Me.Left = frmFileTags.Left + Convert.ToInt32((frmFileTags.Width - Me.Width) / 2)
        Me.Top = frmFileTags.Top + Convert.ToInt32((frmFileTags.Height - Me.Height) / 2)

        'frmAbout.Size = New System.Drawing.Size(240, 160)
        'frmAbout.Text = "File Tags: About"
        'Me.Font = New System.Drawing.Font("Arial", 12)
        'Disable the main form while the about form is up
        ''Me.Enabled = False

        'Dynamically create each label for the dynamic form
        Dim lblAboutBuild As New System.Windows.Forms.Label
        With lblAboutBuild
            .Name = "lblAboutBuild"
            .Text = "File Tags v0.0.1"
            .AutoSize = True
            .Location = New System.Drawing.Point(48, Convert.ToInt32(Me.Height / 100 * 25 - .Height / 2))
        End With

        Dim lblAboutAuthor As New System.Windows.Forms.Label
        With lblAboutAuthor
            .Name = "lblAboutAuthor"
            .Text = "Author:" & vbCrLf & Space(5) & "Selene Tabacchini"
            .AutoSize = True
            .Location = New System.Drawing.Point(48, Convert.ToInt32(Me.Height / 100 * 50 - .Height))
        End With

        Dim lblAboutContact As New System.Windows.Forms.Label
        With lblAboutContact
            .Name = "lblAboutContact"
            .Text = "Contact:"
            .AutoSize = True
            .Location = New System.Drawing.Point(48, Convert.ToInt32(Me.Height / 100 * 75 - .Height))
        End With

        Dim lblAboutEmail As New System.Windows.Forms.LinkLabel
        With lblAboutEmail
            .Name = "lblAboutEmail"
            .Text = "selene.tabacchini@gmail.com"
            .AutoSize = True
            .Location = New System.Drawing.Point(68, Convert.ToInt32(Me.Height / 100 * 75))
            'Bind the click event handler to the sub
            AddHandler .Click, AddressOf emailClicked
        End With

        'Add the labels to the form
        With Me.Controls
            .Add(lblAboutBuild)
            .Add(lblAboutAuthor)
            .Add(lblAboutContact)
            .Add(lblAboutEmail)
        End With
    End Sub

    Private Sub emailClicked(ByVal sender As Object, ByVal e As System.EventArgs)
        'Process variable ctrl of the linklabel, will be used to call upon closing the form
        Dim ctrl As System.Windows.Forms.LinkLabel = CType(sender, System.Windows.Forms.LinkLabel)
        'Handles error of no email application being set
        On Error GoTo EmailError
        'Attempts to open email address
        System.Diagnostics.Process.Start("mailto:selene.tabacchini@gmail.com")
        'Go around the error handling message
        GoTo ClosingSub
EmailError:
        'Displays message to user
        Call frmFileTags.displayMessage("No associated application for emails.")
        Exit Sub
ClosingSub:
    End Sub
End Class