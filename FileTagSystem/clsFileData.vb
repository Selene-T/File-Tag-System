'Author:    Selene Tabacchini
'Pupose:    Handles the current entry's information
'Date:      2012-02-26

'Set options
Option Explicit On
Option Strict On

Public Class clsFileData
    Private mstrFileName As String
    Private mstrFilePath As String
    Private mstrMd5 As String
    Private mstrTags As String

    Public Property FileName As String
        'Handles the file name
        Get
            Return mstrFileName
        End Get
        Set(strValue As String)
            mstrFileName = strValue
        End Set
    End Property

    Public Property FilePath As String
        'Handles the file path
        Get
            Return mstrFilePath
        End Get
        Set(strValue As String)
            mstrFilePath = strValue
        End Set
    End Property

    Public Property MD5 As String
        'Handles the MD5
        Get
            Return mstrMd5
        End Get
        Set(strValue As String)
            mstrMd5 = strValue
        End Set
    End Property

    Public Property Tags As String
        'Handles the list of tags
        Get
            Return mstrTags
        End Get
        Set(strValue As String)
            mstrTags = strValue
        End Set
    End Property
End Class
