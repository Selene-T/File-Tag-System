'Author:    Selene Tabacchini
'Pupose:    Handles database interaction
'Date:      2012-03-02

'Set options
Option Explicit On
Option Strict On

Public Class clsDatabase
    Private mobjOleDbConnection As System.Data.OleDb.OleDbConnection
    Private mobjOleDbCommand As System.Data.OleDb.OleDbCommand
    Private mobjOleDbDataReader As System.Data.OleDb.OleDbDataReader
    Private mstrDatabaseFile As String
    Private mstrConnectionString As String
    Private mtblData As DataTable

    Public ReadOnly Property Conn As System.Data.OleDb.OleDbConnection
        'Connection object, Read Only
        Get
            Return mobjOleDbConnection
        End Get
    End Property

    Public Property Cmd As System.Data.OleDb.OleDbCommand
        'Command object
        Get
            Return mobjOleDbCommand
        End Get
        Set(objValue As System.Data.OleDb.OleDbCommand)
            mobjOleDbCommand = objValue
        End Set
    End Property

    Public Property Database As String
        'Database's location
        Get
            Return mstrDatabaseFile
        End Get
        Set(strValue As String)
            mstrDatabaseFile = strValue
            'Append the .accdb extention is it isn't there
            If Not Strings.Right(mstrDatabaseFile, 6) = ".accdb" Then mstrDatabaseFile += ".accdb"
            'Append the string to the connection string with proper connection code
            mstrConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mstrDatabaseFile & "; Persist Security Info=False;"
        End Set
    End Property

    Public ReadOnly Property Data As DataTable
        'Read Only for the data table, don't want them altering that
        Get
            Return mtblData
        End Get
    End Property

    'Create default constructor
    Public Sub New()
        'Initialize varialbes
        mobjOleDbConnection = Nothing
        mobjOleDbCommand = Nothing
        mobjOleDbDataReader = Nothing
        mstrConnectionString = String.Empty
        mstrDatabaseFile = String.Empty
        mtblData = Nothing
    End Sub

    'Create Methods
    Public Function IsDataEmpty() As Boolean
        'A funtion to return whether or not the database is empty
        If mtblData Is Nothing Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub Command(ByVal strCommand As String)
        'Runs the command given

        'Exits the sub if there is no command
        If mstrConnectionString = String.Empty Or strCommand = String.Empty Then Exit Sub

        'If there's an error with the command, handle it
        On Error GoTo BadCommand

        'Create a new database connection and open
        Using mobjOleDbConnection = New System.Data.OleDb.OleDbConnection(mstrConnectionString)
            mobjOleDbCommand = New System.Data.OleDb.OleDbCommand(strCommand, mobjOleDbConnection)
            mobjOleDbCommand.Connection.Open()

            'Checks to see what kind of command it is
            Select Case Strings.Left(strCommand, 6).ToLower
                Case "insert", "create", "update", "delete"
                    'Commands just ran to the database with no feedback from the database
                    mobjOleDbCommand.ExecuteNonQuery()
                Case "select"
                    'Select needs to be exectuted into a reader for the feedback
                    mobjOleDbDataReader = mobjOleDbCommand.ExecuteReader()

                    'Check if there was any returned values
                    If Not mobjOleDbDataReader.HasRows Then
                        'If not, clear the table and exit the sub
                        mtblData = Nothing
                        Exit Sub
                    End If

                    'Initialize new variables and a new data table
                    Dim intReads As Int32 = 0
                    mtblData = New DataTable("Data")
                    With mobjOleDbDataReader
                        'While we are still reading from the feedback...
                        Do While .Read()
                            'On the first read through, we want to grab the column names and create them in the data table
                            If intReads = 0 Then
                                For i As Int32 = 0 To .FieldCount - 1
                                    'Grab the name of each field as the column title
                                    mtblData.Columns.Add(.GetName(i), GetType(String))
                                Next
                            End If
                            'Add an empty row before it can be filled
                            mtblData.Rows.Add()
                            For i As Int32 = 0 To .FieldCount - 1
                                'Add the item from each field into the data table
                                mtblData.Rows.Item(intReads).Item(i) = .Item(i)
                            Next

                            'Increase the reads, or depth/row
                            intReads += 1
                        Loop

                        'Close the reader
                        .Close()
                    End With
            End Select

            'Close the connection to the database
            mobjOleDbCommand.Connection.Close()
            mobjOleDbConnection.Close()
        End Using
        Exit Sub
BadCommand:
        'Error handler, notify user of the bad command and show it to them
        Msg("Command string is invalid:" & vbNewLine & strCommand)
    End Sub

    Public Sub CreateDB(ByVal strCreateDatabase As String)
        'Call to create a new database file
        Dim dbNew As New ADOX.Catalog

        'Checks to make sure the file doesn't already exist
        If Not strCreateDatabase = String.Empty And Not System.IO.File.Exists(strCreateDatabase) Then
            dbNew.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strCreateDatabase)
        End If
    End Sub

    Public Sub Close()
        'Call to close the connection
        If mobjOleDbConnection.State = ConnectionState.Open Then mobjOleDbConnection.Close()
    End Sub

    Public Function State() As ConnectionState
        'Call to gather the connection state of the database
        On Error GoTo StateNull
        Return mobjOleDbConnection.State
StateNull:
        'Handles the error is there is no state
        Return ConnectionState.Closed
    End Function

    Private Sub Msg(ByVal strMessage As String)
        'Call to display the error message
        MessageBox.Show(strMessage, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub
End Class
