

'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program core procedures.
Public Module CoreModule
   'This enumeration lists the supported data file formats.
   Private Enum FormatsE As Byte
      Uncompressed = &H0%   'Defines uncompressed data.
      Compressed = &H80%    'Defines compressed data.
      CBMImage = &HC0%      'Defines CBM image data.
   End Enum

   'This structure defines a directory file entry.
   Private Structure DirFileEntryStr
      Public Offset As Integer    'Defines a directory file's offset.
      Public Length As Integer    'Defines a directory file's length.
      Public FileName As String   'Defines a directory file's name.
      Public Format As FormatsE   'Defines a directory file's format.
   End Structure

   Private ReadOnly DAT_FILE_SIGNATURE() As Byte = {&HCA%, &HCA%, &HD0%, &HD0%}                                   'Defines the data file signature.
   Private ReadOnly INVALID_CHARACTERS() As Char = {"*"c, "/"c, "<"c, ">"c, "?"c, "["c, "\"c, "]"c, "|"c, " "c}   'Defines characters that are invalid in file names in MS-DOS.
   Private ReadOnly PADDING As Char = ToChar(&H0%)                                                                'Defines the null character used to terminate and pad file names.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim DatFile As String = Nothing
         Dim DirFileEntries As List(Of DirFileEntryStr) = Nothing
         Dim DirFiles() As FileInfo = {}
         Dim SourcePath As String = Nothing
         Dim TargetPath As String = Nothing

         If GetCommandLineArgs().Count = 2 Then
            SourcePath = GetCommandLineArgs().Last()
            If Directory.Exists(SourcePath) Then
               TargetPath = Path.Combine(SourcePath, "Data")
               Console.WriteLine($"Extracting to: {TargetPath}")

               If Directory.Exists(TargetPath) Then
                  Console.WriteLine($"WARNING: {TargetPath} already exists. - Any previously extracted files will be overwritten.")
               Else
                  Directory.CreateDirectory(TargetPath)
               End If

               DirFiles = My.Computer.FileSystem.GetDirectoryInfo(SourcePath).GetFiles("*.dir")
               If DirFiles.Count = 0 Then
                  Throw New Exception("No *.dir files found.")
               Else
                  For Each DirFile As FileInfo In DirFiles
                     Console.WriteLine($"Reading directory: {DirFile.FullName}")
                     DatFile = $"{Path.Combine(SourcePath, Path.GetFileNameWithoutExtension(DirFile.FullName))}.DAT"
                     DirFileEntries = GetDirectoryEntries(DirFile.FullName)
                     CheckDataSize(DirFileEntries, DatFile)
                     CheckForDataOverlap(DirFileEntries)
                     Console.WriteLine($"Extracting data file: {DatFile}")
                     For Each DirFileEntry As DirFileEntryStr In DirFileEntries
                        With DirFileEntry
                           Console.WriteLine($"Writing: { .FileName} ({ .Format.ToString()})")
                           File.WriteAllBytes(Path.Combine(TargetPath, .FileName), GetDatFileData(DatFile).GetRange(.Offset, .Length).ToArray())
                        End With
                     Next DirFileEntry
                  Next DirFile
               End If
            Else
               Throw New Exception($"Could not find: {SourcePath}.")
            End If
         Else
            With My.Application.Info
               Console.WriteLine($"{ .Title} v{ .Version}, by: { .CompanyName}, ***{ .Copyright }***")
               Console.WriteLine()
               Console.WriteLine($"Usage: ""{ .AssemblyName}.exe"" path")
               Console.WriteLine()
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure checks whether all data has been read from the specified file and displays a warning if not.
   Private Sub CheckDataSize(DirFileEntries As List(Of DirFileEntryStr), DatFile As String)
      Try
         Dim CalculatedDatFileLength As Integer = DirFileEntries.Sum(Function(DirFileEntry) DirFileEntry.Length)
         Dim DatFileLength As Integer = CInt(New FileInfo(DatFile).Length) - DAT_FILE_SIGNATURE.Length

         If Not CalculatedDatFileLength = DatFileLength Then
            Console.WriteLine($"WARNING: Not all data has been read: {DatFileLength - CalculatedDatFileLength} bytes.")
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure checks whether the data files overlap and displays a warning if so.
   Private Sub CheckForDataOverlap(DirFileEntries As List(Of DirFileEntryStr))
      Try
         Dim DirFileLength As New Integer
         Dim DirFileNextOffset As New Integer
         Dim DirFileOffset As New Integer

         DirFileEntries = New List(Of DirFileEntryStr)(From DirectoryEntry In DirFileEntries Order By DirectoryEntry.Offset)
         For Index As Integer = 0 To DirFileEntries.Count - 2
            DirFileOffset = DirFileEntries(Index).Offset
            DirFileLength = DirFileEntries(Index).Length
            DirFileNextOffset = DirFileEntries(Index + 1).Offset

            If DirFileOffset + DirFileLength < DirFileNextOffset Then
               Console.WriteLine($"Offset {DirFileOffset} overlaps with {DirFileNextOffset} by {DirFileNextOffset - (DirFileOffset + DirFileLength)} bytes.")
            End If
         Next Index
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure retrieves and returns the specified data file's contents.
   Private Function GetDatFileData(DatFile As String) As List(Of Byte)
      Try
         Dim Data As New List(Of Byte)(File.ReadAllBytes(DatFile))

         If Not Data.GetRange(0, DAT_FILE_SIGNATURE.Length).SequenceEqual(DAT_FILE_SIGNATURE) Then
            Throw New Exception("Invalid data file signature.")
         End If

         Return Data
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure retrieves the directory entries from the specified directory file and returns the result.
   Private Function GetDirectoryEntries(DirFile As String) As List(Of DirFileEntryStr)
      Try
         Dim DirFileEntries As New List(Of DirFileEntryStr)
         Dim FileName As String = Nothing
         Dim Format As New FormatsE
         Dim Length As New Integer
         Dim Offset As New Integer

         Using FileO As New BinaryReader(File.Open(DirFile, FileMode.Open))
            Do Until FileO.BaseStream.Position >= FileO.BaseStream.Length
               Offset = FileO.ReadInt32()
               Length = FileO.ReadUInt16()
               FileName = FileO.ReadChars(13)
               FileName = FileName.Substring(0, FileName.IndexOf(PADDING))
               Format = DirectCast(FileO.ReadByte(), FormatsE)

               If FileName.Intersect(INVALID_CHARACTERS).Count > 0 Then
                  Throw New Exception("Invalid characters found in filename.")
               Else
                  DirFileEntries.Add(New DirFileEntryStr With {.Offset = Offset, .Length = Length, .FileName = FileName, .Format = Format})
               End If

               If Array.IndexOf([Enum].GetValues(GetType(FormatsE)), Format) < 0 Then
                  Console.WriteLine($"WARNING: Unsupported format: 0x{Format:X}.")
               End If
            Loop
         End Using

         Return DirFileEntries
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of DirFileEntryStr)
   End Function

   'This procedure handles any errors that occur.
   Private Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine($"ERROR: {ExceptionO.Message}")
         [Exit](0)
      Catch
         [Exit](0)
      End Try
   End Sub
End Module
