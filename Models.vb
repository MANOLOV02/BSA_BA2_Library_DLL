Option Strict On
Imports System.Text

Namespace BethesdaArchive.Core

    Public NotInheritable Class VirtualEntry
        Public Property Directory As String = ""
        Public Property FileName As String = ""
        Public ReadOnly Property FullPath As String
            Get
                If String.IsNullOrEmpty(Directory) Then Return FileName
                Return Directory.TrimEnd(Correct_Path_separator, InCorrect_Path_separator) & Correct_Path_separator & FileName
            End Get
        End Property
        Public Property Data As Byte()             ' Contenido lógico
        Public Property PreferCompress As Boolean  ' Para BSA/BA2.GNRL

        ' DX10 metadata
        Public Property DxgiFormat As Integer
        Public Property Width As Integer
        Public Property Height As Integer
        Public Property MipCount As Integer
        Public Property IsCubemap As Boolean
        Public Property Faces As Integer
    End Class

    Public NotInheritable Class ArchiveChangeSet
        Public ReadOnly Adds As New Dictionary(Of String, VirtualEntry)(StringComparer.OrdinalIgnoreCase)
        Public ReadOnly Deletes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Public ReadOnly Renames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public ReadOnly Replacements As New Dictionary(Of String, VirtualEntry)(StringComparer.OrdinalIgnoreCase)
    End Class

    Public Enum GameKind
        SSE_BSA
        FO4_BA2
    End Enum

    Public NotInheritable Class ArchiveHandle
        Public ReadOnly Property Kind As GameKind
        Public ReadOnly Property SourcePath As String
        Public ReadOnly Property Encoding As Encoding
        Public ReadOnly ChangeSet As ArchiveChangeSet = New ArchiveChangeSet()
        Public ReadOnly Entries As List(Of VirtualEntry)

        Public Sub New(kind As GameKind, sourcePath As String, encoding As Encoding, entries As List(Of VirtualEntry))
            Me.Kind = kind
            Me.SourcePath = sourcePath
            Me.Encoding = If(encoding, Encoding.UTF8)
            Me.Entries = entries
        End Sub
    End Class

    Public NotInheritable Class MultiArchiveEditor
        Private ReadOnly _handles As New List(Of ArchiveHandle)
        Public ReadOnly Property [Handles] As IReadOnlyList(Of ArchiveHandle)
            Get
                Return _handles
            End Get
        End Property

        Public Sub OpenMany(kind As GameKind, archives As IEnumerable(Of (path As String, entries As List(Of VirtualEntry), enc As Encoding)))
            For Each a In archives
                _handles.Add(New ArchiveHandle(kind, a.path, a.enc, a.entries))
            Next
        End Sub
    End Class

    Public Module PathUtil
        Public Const Correct_Path_separator As Char = "\"c
        Public Const InCorrect_Path_separator As Char = "/"c
        Public Function NormalizeSlash(p As String) As String
            If String.IsNullOrEmpty(p) Then Return ""
            Dim result = p.Replace(InCorrect_Path_separator, Correct_Path_separator)
            Return result
        End Function
        Public Function JoinDirFile(dir As String, file As String) As String
            If String.IsNullOrEmpty(dir) Then Return file
            Return NormalizeSlash(dir).TrimEnd(Correct_Path_separator) & Correct_Path_separator & file
        End Function
        Public Function MakeRelativeUnderDataRoot(absPath As String, dataRoot As String) As String
            Dim root = If(dataRoot, "")
            Try
                If root.Length > 0 Then
                    Dim fullRoot = IO.Path.GetFullPath(root).TrimEnd(IO.Path.DirectorySeparatorChar, IO.Path.AltDirectorySeparatorChar)
                    Dim fullAbs = IO.Path.GetFullPath(absPath)
                    If fullAbs.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase) Then
                        Dim rel = fullAbs.Substring(fullRoot.Length).TrimStart(IO.Path.DirectorySeparatorChar, IO.Path.AltDirectorySeparatorChar)
                        Return NormalizeSlash(rel)
                    Else
                        Return fullAbs
                    End If
                Else
                    Return absPath
                End If
            Catch
            End Try
            Return NormalizeSlash(IO.Path.GetFileName(absPath))
        End Function

        Public Sub SplitDirFile(relPath As String, ByRef dir As String, ByRef file As String)
            Dim relNorm As String = NormalizeSlash(relPath)
            dir = ""
            file = relNorm
            Dim slash As Integer = relNorm.LastIndexOf(Correct_Path_separator)
            If slash >= 0 Then
                dir = If(slash = 0, "", relNorm.Substring(0, slash))
                file = relNorm.Substring(slash + 1)
            End If
        End Sub
    End Module

End Namespace

