' Version Uploaded of Wardrobe 2.1.3
Imports System.IO
Imports System.Runtime
Imports System.Text
Imports DirectXTexWrapperCLI ' Loader.EncodeDDSHeader(...)
Imports ICSharpCode.SharpZipLib.Zip.Compression
Imports K4os.Compression.LZ4.Streams

' ======== API pública ========
Namespace BethesdaArchive.Core
    Public Class BethesdaReader
        Implements IDisposable

        Public ReadOnly Property EntriesFiles As IReadOnlyList(Of ArchiveEntry)

        Private ReadOnly _impl As IArchiveImpl
        Private ReadOnly _entries As List(Of ArchiveEntry)

        ''' <summary>Abre BSA (SSE v105) o BA2 (FO4 GNRL/DX10).</summary>
        Public Sub New(fs As Stream, Optional encoding As Encoding = Nothing)
            If encoding Is Nothing Then encoding = Encoding.UTF8
            _impl = DetectAndOpen(fs, encoding)
            _entries = _impl.ListEntries()
            EntriesFiles = _entries.AsReadOnly()
        End Sub

        ''' <summary>a) Lista expuesta en .EntriesFiles; b) descomprime en memoria por índice.</summary>
        Public Function ExtractToMemory(index As Integer) As Byte()
            If index < 0 OrElse index >= _entries.Count Then Throw New ArgumentOutOfRangeException(NameOf(index))
            Return _impl.ExtractByIndex(index)
        End Function

        Private Shared Function DetectAndOpen(fs As Stream, enc As Encoding) As IArchiveImpl
            If Not fs.CanSeek Then Throw New ArgumentException("El Stream debe soportar Seek")
            Dim start = fs.Position
            Using br As New BinaryReader(fs, enc, leaveOpen:=True)
                ' BA2: "BTDX"
                If fs.Length - fs.Position >= 4 Then
                    Dim m = Encoding.ASCII.GetString(br.ReadBytes(4))
                    fs.Position = start
                    If m = "BTDX" Then Return New Ba2Impl(fs, enc)
                End If
                ' BSA: 0x00415342 = "BSA\0"
                If fs.Length - fs.Position >= 4 Then
                    Dim magic = br.ReadUInt32()
                    fs.Position = start
                    If magic = &H415342UI Then Return New BsaImpl(fs, enc)
                End If
            End Using
            Throw New InvalidDataException("Formato no reconocido (BSA SSE o BA2 FO4).")
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _impl?.Dispose()
        End Sub
    End Class

    Public Class ArchiveEntry
        Public Property Index As Integer
        Public Property Directory As String
        Public Property FileName As String
        Public ReadOnly Property FullPath As String
            Get
                If String.IsNullOrEmpty(Directory) Then Return FileName
                Return Directory.TrimEnd(Correct_Path_separator, InCorrect_Path_separator) & Correct_Path_separator & FileName
            End Get
        End Property
    End Class

    Friend Interface IArchiveImpl
        Inherits IDisposable
        Function ListEntries() As List(Of ArchiveEntry)
        Function ExtractByIndex(index As Integer) As Byte()
    End Interface

    ' ============ DECOMPRESSERS =============
    Friend Module Lz4Strict
        ''' <summary>
        ''' Descomprime LZ4 de forma determinista.
        ''' - Si el buffer comienza con magic de FRAME (0x184D2204) o SKIPPABLE (0x184D2A50..5F), procesa como FRAME:
        '''   soporta múltiples frames concatenados y bloques skippables.
        ''' - En caso contrario, trata el buffer como RAW BLOCK.
        ''' En ambos modos exige que la salida tenga exactamente 'expected' bytes.
        ''' </summary>
        Friend Function Lz4DecompressStrict(packed As Byte(), expected As Integer) As Byte()
            If expected <= 0 Then
                Throw New InvalidDataException("LZ4: 'expected' debe ser > 0 para decodificación estricta.")
            End If
            If packed Is Nothing OrElse packed.Length = 0 Then
                Throw New InvalidDataException("LZ4: buffer vacío.")
            End If

            ' ---- Detección de FRAME / SKIPPABLE por magic ----
            Dim isFrame As Boolean = False
            Dim isSkippableStart As Boolean = False
            If packed.Length >= 4 Then
                Dim b0 = packed(0) : Dim b1 = packed(1) : Dim b2 = packed(2) : Dim b3 = packed(3)
                isFrame = (b0 = &H4 AndAlso b1 = &H22 AndAlso b2 = &H4D AndAlso b3 = &H18) ' 04-22-4D-18
                isSkippableStart = (b1 = &H2A AndAlso b2 = &H4D AndAlso b3 = &H18 AndAlso b0 >= &H50 AndAlso b0 <= &H5F) ' 5x-2A-4D-18
            End If

            If isFrame OrElse isSkippableStart Then
                ' --------- FRAME path (múltiples frames + skippables) ---------
                Using msIn As New MemoryStream(packed, writable:=False),
                  msOut As New MemoryStream(expected)

                    While msIn.Position < msIn.Length
                        If msIn.Length - msIn.Position < 4 Then Exit While

                        ' Peek 4 bytes
                        Dim savePos = msIn.Position
                        Dim h0 = msIn.ReadByte() : Dim h1 = msIn.ReadByte() : Dim h2 = msIn.ReadByte() : Dim h3 = msIn.ReadByte()
                        msIn.Position = savePos
                        If h0 < 0 OrElse h1 < 0 OrElse h2 < 0 OrElse h3 < 0 Then Exit While

                        Dim hdr0 = CByte(h0) : Dim hdr1 = CByte(h1) : Dim hdr2 = CByte(h2) : Dim hdr3 = CByte(h3)
                        Dim frame = (hdr0 = &H4 AndAlso hdr1 = &H22 AndAlso hdr2 = &H4D AndAlso hdr3 = &H18)
                        Dim skippable = (hdr1 = &H2A AndAlso hdr2 = &H4D AndAlso hdr3 = &H18 AndAlso hdr0 >= &H50 AndAlso hdr0 <= &H5F)

                        If skippable Then
                            ' magic + length (u32 LE) + payload skippable
                            msIn.Position += 4
                            If msIn.Length - msIn.Position < 4 Then Throw New InvalidDataException("LZ4 frame: skippable sin longitud.")
                            Dim len0 = msIn.ReadByte(), len1 = msIn.ReadByte(), len2 = msIn.ReadByte(), len3 = msIn.ReadByte()
                            If len0 < 0 OrElse len1 < 0 OrElse len2 < 0 OrElse len3 < 0 Then Throw New EndOfStreamException()
                            Dim skipLen As Integer = CInt(len0) Or (CInt(len1) << 8) Or (CInt(len2) << 16) Or (CInt(len3) << 24)
                            Dim newPos = msIn.Position + skipLen
                            If newPos > msIn.Length Then Throw New InvalidDataException("LZ4 frame: skippable excede el buffer.")
                            msIn.Position = newPos
                            Continue While
                        ElseIf frame Then
                            ' decodificar UN frame completo; el decoder avanza msIn hasta el fin del frame
                            Using dec As Stream = LZ4Stream.Decode(msIn, leaveOpen:=True)
                                dec.CopyTo(msOut)
                            End Using
                            Continue While
                        Else
                            ' bytes inesperados entre frames → error duro; no “probamos” nada
                            Throw New InvalidDataException("LZ4 frame: bytes inesperados entre frames/skippables.")
                        End If
                    End While

                    Dim outB = msOut.ToArray()
                    If outB.Length <> expected Then
                        Throw New InvalidDataException($"LZ4 frame: tamaño descomprimido {outB.Length} != esperado {expected}.")
                    End If
                    Return outB
                End Using
            Else
                ' --------- RAW BLOCK path (sin knownOutputLength) ---------
                Dim outBytes As Byte() = New Byte(expected - 1) {}
                Dim written As Integer = K4os.Compression.LZ4.LZ4Codec.Decode(packed, 0, packed.Length, outBytes, 0, expected)
                If written < 0 Then
                    Throw New InvalidDataException("LZ4 raw: error de decodificación.")
                End If
                If written <> expected Then
                    Throw New InvalidDataException($"LZ4 raw: tamaño descomprimido {written} != esperado {expected}.")
                End If
                Return outBytes
            End If
        End Function
    End Module
    Friend Module ZlibStrict

        ''' <summary>
        ''' Descomprime un stream ZLIB (RFC1950) de forma estricta:
        ''' - Verifica header CMF/FLG (CM=8 y (CMF<<8+FLG) mod 31 == 0).
        ''' - Descomprime con header y Adler-32.
        ''' - Exige que el output tenga exactamente 'expected' bytes y que no queden bytes sin consumir.
        ''' </summary>
        Friend Function ZlibDecompressStrict(packed As Byte(), expected As Integer) As Byte()
            If expected <= 0 Then Throw New InvalidDataException("Zlib: 'expected' debe ser > 0.")
            If packed Is Nothing OrElse packed.Length = 0 Then Throw New InvalidDataException("Zlib: buffer vacío.")
            If packed.Length < 2 Then Throw New InvalidDataException("Zlib: buffer demasiado corto (faltan CMF/FLG).")

            ' Validación de cabecera ZLIB: CMF/FLG
            Dim cmf As Integer = packed(0)
            Dim flg As Integer = packed(1)
            If (cmf And &HF) <> 8 Then Throw New InvalidDataException("Zlib: CM != 8 (no es deflate).")
            If (((cmf << 8) + flg) Mod 31) <> 0 Then Throw New InvalidDataException("Zlib: header CMF/FLG inválido (mod31).")

            Dim inflater As New Inflater(noHeader:=False)
            inflater.SetInput(packed)

            Dim outBytes As Byte() = New Byte(expected - 1) {}
            Dim total As Integer = 0

            ' Inflar hasta alcanzar exactamente expected, o hasta quedar sin input útil
            While total < expected AndAlso Not inflater.IsNeedingInput
                Dim wrote = inflater.Inflate(outBytes, total, expected - total)
                If wrote = 0 Then Exit While ' no avanza más (input agotado o estado final)
                total += wrote
            End While

            If total <> expected Then
                Throw New InvalidDataException($"Zlib: tamaño descomprimido {total} != esperado {expected}.")
            End If

            Return outBytes
        End Function

        ''' <summary>
        ''' Descomprime un stream DEFLATE RAW (RFC1951, sin header zlib) de forma estricta:
        ''' - No acepta header zlib.
        ''' - Exige output de tamaño exacto y fin de stream.
        ''' </summary>
        Friend Function DeflateRawDecompressStrict(packed As Byte(), expected As Integer) As Byte()
            If expected <= 0 Then Throw New InvalidDataException("DeflateRaw: 'expected' debe ser > 0.")
            If packed Is Nothing OrElse packed.Length = 0 Then Throw New InvalidDataException("DeflateRaw: buffer vacío.")

            ' Si detectás header zlib, consideralo un error aquí (es RAW)
            If packed.Length >= 2 Then
                Dim cmf As Integer = packed(0)
                Dim flg As Integer = packed(1)
                Dim looksZlib As Boolean = ((cmf And &HF) = 8) AndAlso ((((cmf << 8) + flg) Mod 31) = 0)
                If looksZlib Then Throw New InvalidDataException("DeflateRaw: se detectó header zlib (usar ZlibDecompressStrict).")
            End If

            Dim inflater As New Inflater(noHeader:=True)
            inflater.SetInput(packed)

            Dim outBytes As Byte() = New Byte(expected - 1) {}
            Dim total As Integer = 0
            While total < expected AndAlso Not inflater.IsFinished AndAlso Not inflater.IsNeedingInput
                total += inflater.Inflate(outBytes, total, expected - total)
            End While

            If total <> expected Then
                Throw New InvalidDataException($"DeflateRaw: tamaño descomprimido {total} != esperado {expected}.")
            End If
            If Not inflater.IsFinished Then
                Throw New InvalidDataException("DeflateRaw: flujo truncado (no alcanzó estado Finished).")
            End If
            If inflater.RemainingInput <> 0 Then
                Throw New InvalidDataException("DeflateRaw: quedaron bytes extra tras terminar (input no consumido).")
            End If

            Return outBytes
        End Function

    End Module


    ' ======== BSA (Skyrim SE v105) ========


    Friend Class BsaImpl
        Implements IArchiveImpl

        Private ReadOnly _fs As Stream
        Private ReadOnly _br As BinaryReader
        Private ReadOnly _enc As Encoding

        Private _hdr As BsaHeader
        Private _fileDataBase As Long
        Private ReadOnly _records As New List(Of FileRec)

        ' Flags
        Private Const ARCHIVE_COMPRESS As UInteger = &H4UI
        Private Const ARCHIVE_EMBED_NAMES As UInteger = &H100UI
        Private Shared ReadOnly ICOMPRESSION As UInteger = CUInt(1UI << 30)
        Private Shared ReadOnly ICHECKED As UInteger = CUInt(1UI << 31)

        Private Structure BsaHeader
            Public Magic As UInteger          ' "BSA\0"
            Public Version As UInteger        ' 105 (0x69)
            Public HeaderSize As UInteger     ' 0x24
            Public ArchiveFlags As UInteger
            Public FolderCount As UInteger
            Public FileCount As UInteger
            Public FolderNameLength As UInteger
            Public FileNameLength As UInteger
            Public FileFlags As UInteger
        End Structure

        Private Class FileRec
            Public Index As Integer
            Public Directory As String
            Public FileName As String
            Public Offset As UInteger         ' ABSOLUTO
            Public SizeField As UInteger      ' con bits 30/31
            Public Compressed As Boolean
            Public HasEmbeddedName As Boolean ' <<< NUEVO
        End Class

        Public Sub New(fs As Stream, enc As Encoding)
            _fs = fs : _enc = enc
            _br = New BinaryReader(_fs, _enc, leaveOpen:=True)
            Open()
        End Sub

        Private Shared Function ReadZString(br As BinaryReader, enc As Encoding) As String
            Dim tmp As New List(Of Byte)(64)
            While True
                Dim b = br.ReadByte()
                If b = 0 Then Exit While
                tmp.Add(b)
            End While
            Return enc.GetString(tmp.ToArray())
        End Function

        Private Shared Function ReadBZ(br As BinaryReader, enc As Encoding) As String
            Dim l = br.ReadByte() ' incluye NULL
            If l = 0 Then Return ""
            Dim raw = br.ReadBytes(l - 1)
            Dim _null = br.ReadByte()
            Return enc.GetString(raw)
        End Function

        Private Sub Open()
            ' --- Header ---
            _hdr.Magic = _br.ReadUInt32()
            If _hdr.Magic <> &H415342UI Then Throw New InvalidDataException("BSA: magic inválido")
            _hdr.Version = _br.ReadUInt32()
            If _hdr.Version <> &H69UI Then Throw New NotSupportedException("BSA: sólo SSE v105")
            _hdr.HeaderSize = _br.ReadUInt32()         ' 0x24
            _hdr.ArchiveFlags = _br.ReadUInt32()
            _hdr.FolderCount = _br.ReadUInt32()
            _hdr.FileCount = _br.ReadUInt32()
            _hdr.FolderNameLength = _br.ReadUInt32()
            _hdr.FileNameLength = _br.ReadUInt32()
            _hdr.FileFlags = _br.ReadUInt32()

            ' --- Offsets v105 ---
            Dim directoriesOffset As Long = _hdr.HeaderSize
            Dim dirEntrySize As Integer = &H18
            Dim filesOffset As Long = directoriesOffset + CLng(dirEntrySize) * CLng(_hdr.FolderCount)

            Dim hasDirStrings As Boolean = ((_hdr.ArchiveFlags And 1UI) <> 0UI)
            Dim hasFileStrings As Boolean = ((_hdr.ArchiveFlags And 2UI) <> 0UI)

            Dim dirStrSz As Long = If(hasDirStrings, CLng(_hdr.FolderNameLength) + CLng(_hdr.FolderCount), 0L)
            Dim fileStringsStart As Long = filesOffset + dirStrSz + CLng(_hdr.FileCount) * &H10
            _fileDataBase = fileStringsStart + CLng(_hdr.FileNameLength) ' (no se usa para calcular offsets de archivo)

            ' --- 1) Leer sólo los counts de carpetas ---
            _fs.Position = directoriesOffset
            Dim folderCounts As New List(Of UInteger)(CInt(_hdr.FolderCount))
            For i = 0UI To _hdr.FolderCount - 1UI
                Dim _hash = _br.ReadUInt64()
                Dim cnt = _br.ReadUInt32()
                folderCounts.Add(cnt)
                _fs.Position += 12 ' resto del folder record v105 (0x18 total)
            Next

            ' --- 2) Recorrido intercalado: [Dir BZString] + [count * FileEntry] ---
            Dim filesCursor As Long = filesOffset
            Dim namesCursor As Long = fileStringsStart

            _records.Clear()
            _records.Capacity = CInt(_hdr.FileCount)
            Dim idx As Integer = 0
            Dim globCompressed As Boolean = ((_hdr.ArchiveFlags And ARCHIVE_COMPRESS) <> 0UI)

            For d = 0 To folderCounts.Count - 1
                ' 2.a) BZString de directorio (si existen nombres de carpeta)
                Dim dirName As String = ""
                If hasDirStrings Then
                    _fs.Position = filesCursor
                    dirName = PathUtil.NormalizeSlash(ReadBZ(_br, _enc))   ' avanza según longitud+1 (incluye null)
                    filesCursor = _fs.Position               ' avanzar cursor tras el BZString
                End If

                ' 2.b) 'count' file entries de esta carpeta
                For j = 0UI To folderCounts(d) - 1UI
                    _fs.Position = filesCursor

                    Dim fhash = _br.ReadUInt64()
                    Dim sizeField = _br.ReadUInt32()
                    Dim offAbs As UInteger = _br.ReadUInt32()  ' offset ABSOLUTO (v105)
                    filesCursor = _fs.Position

                    ' nombre de archivo desde la tabla de nombres (zstring)
                    Dim fileName As String = ""
                    If hasFileStrings Then
                        Dim ret = _fs.Position
                        _fs.Position = namesCursor
                        fileName = ReadZString(_br, _enc)
                        namesCursor = _fs.Position
                        _fs.Position = ret
                    End If

                    ' compresión efectiva = flag global XOR bit30
                    Dim comp As Boolean = globCompressed
                    If (sizeField And ICOMPRESSION) <> 0UI Then comp = Not comp
                    Dim embedNames As Boolean = ((_hdr.ArchiveFlags And ARCHIVE_EMBED_NAMES) <> 0UI)

                    ' al crear cada registro:
                    _records.Add(New FileRec With {
    .Index = idx, .Directory = dirName, .FileName = fileName,
    .Offset = offAbs, .SizeField = sizeField, .Compressed = comp,
    .HasEmbeddedName = embedNames   ' <<< SIEMPRE según flag 0x100
})
                    idx += 1
                Next
            Next
        End Sub



        Public Function ListEntries() As List(Of ArchiveEntry) Implements IArchiveImpl.ListEntries
            Dim list As New List(Of ArchiveEntry)(_records.Count)
            For Each r In _records
                list.Add(New ArchiveEntry With {.Index = r.Index, .Directory = r.Directory, .FileName = r.FileName})
            Next
            Return list
        End Function

        Public Function ExtractByIndex(index As Integer) As Byte() Implements IArchiveImpl.ExtractByIndex
            Dim r = _records(index)

            ' 1) Posición física (offset absoluto con bit 31 limpio por posible “secondary archive”)
            Dim offPhys As UInteger = (r.Offset And &H7FFFFFFFUI)
            _fs.Position = CLng(offPhys)

            ' 2) Trabajar con copia del size del entry
            Dim sizeWork As UInteger = (r.SizeField And &H3FFFFFFFUI)

            ' 3) Nombre embebido (BSTRING: u8 len + len bytes, SIN terminador)
            Dim hasEmbeddedName As Boolean = r.HasEmbeddedName
            If hasEmbeddedName Then
                Dim nameLen As Integer = _fs.ReadByte()
                If nameLen < 0 Then Throw New EndOfStreamException("BSA: EOF leyendo longitud de nombre embebido.")
                If _fs.Position + nameLen > _fs.Length Then
                    Throw New EndOfStreamException("BSA: nombre embebido excede el archivo.")
                End If
                If nameLen > 0 Then _fs.Position += nameLen
                ' Descontar: byte de longitud + bytes del nombre
                sizeWork -= CUInt(nameLen + 1)
            End If

            ' 4) ¿Comprimido?
            Dim isComp As Boolean = r.Compressed
            Dim uncompLen As UInteger = 0UI
            If isComp Then
                If sizeWork < 4UI Then Throw New InvalidDataException("BSA: size < 4 en archivo comprimido.")
                uncompLen = _br.ReadUInt32() ' u32 LE de tamaño descomprimido
                If uncompLen = 0UI Then Throw New InvalidDataException("BSA: tamaño descomprimido = 0.")
                sizeWork -= 4UI ' restar antes de limpiar flags
            End If

            ' 5) Leer payload
            Dim payloadLen As Integer = CInt(sizeWork)
            If payloadLen < 0 Then Throw New InvalidDataException("BSA: payloadLen negativo.")
            Dim payload As Byte() = If(payloadLen > 0, _br.ReadBytes(payloadLen), Array.Empty(Of Byte)())
            If payload.Length <> payloadLen Then
                Throw New EndOfStreamException($"BSA: lectura corta del payload (read={payload.Length}, need={payloadLen}).")
            End If

            ' 7) No comprimido → devolver
            If Not isComp Then
                Return payload
            End If

            ' 8) LZ4 (FRAME/RAW) estricto con tamaño esperado
            Dim outData As Byte()
            outData = Lz4Strict.Lz4DecompressStrict(payload, CInt(uncompLen))
            Return outData
        End Function



        Public Sub Dispose() Implements IArchiveImpl.Dispose
            _br?.Dispose()
        End Sub
    End Class

    ' ======== BA2 (FO4: GNRL / DX10) ========

    Friend Class Ba2Impl
        Implements IArchiveImpl

        Private ReadOnly _fs As Stream
        Private ReadOnly _br As BinaryReader
        Private ReadOnly _enc As Encoding

        Private _hdr As Ba2Header
        Private ReadOnly _entries As New List(Of EntryBase)

        Private Enum Ba2Type
            GNRL
            DX10
            GNMF
        End Enum

        Private Structure Ba2Header
            Public Version As UInteger
            Public Type As Ba2Type
            Public NumFiles As UInteger
            Public NameTableOffset As ULong
            Public V2V3_UnknownAlways1 As ULong ' v2/v3: suele ser 1
            Public V3_Compression As UInteger   ' v3: 3=LZ4, otro=zlib
        End Structure

        Private MustInherit Class EntryBase
            Public Property Index As Integer
            Public Property Directory As String = ""
            Public Property FileName As String = ""
            Public MustOverride Function Extract(fs As Stream, hdr As Ba2Header) As Byte()
        End Class

        ' ---------- GNRL ----------
        Private Class EntryGNRL
            Inherits EntryBase

            Public HashFile As UInteger
            Public HashExt As UInteger
            Public HashDir As UInteger
            Public DataFileIndex As Byte
            Public ChunkCount As Byte
            Public ChunkHeaderSize As UShort

            Public Chunks As New List(Of Chunk)

            Public Structure Chunk
                Public Offset As ULong
                Public CompressedSize As UInteger ' 0 => sin compresión
                Public DecompressedSize As UInteger
                Public Sentinel As UInteger       ' 0xBAADF00D
            End Structure

            Public Overrides Function Extract(fs As Stream, hdr As Ba2Header) As Byte()
                If Chunks Is Nothing OrElse Chunks.Count = 0 Then
                    Return Array.Empty(Of Byte)()
                End If

                ' Orden físico por Offset por robustez (no debería alterar si ya viene ordenado)
                Dim ordered = Chunks.OrderBy(Function(c) c.Offset).ToList()

                Using ms As New MemoryStream()
                    For Each ch In ordered

                        If CLng(ch.Offset) < 0 OrElse CLng(ch.Offset) >= fs.Length Then
                            Throw New InvalidDataException($"BA2.GNRL: offset fuera de rango (offset={ch.Offset}).")
                        End If
                        fs.Position = CLng(ch.Offset)

                        Dim isComp As Boolean = (ch.CompressedSize <> 0UI)
                        Dim needUL As ULong = If(isComp, CULng(ch.CompressedSize), CULng(ch.DecompressedSize))
                        If needUL > CULng(Integer.MaxValue) Then
                            Throw New InvalidDataException($"BA2.GNRL: tamaño de chunk excede Int32 (need={needUL}).")
                        End If
                        Dim need As Integer = CInt(needUL)

                        Dim packed As Byte() = If(need > 0, New Byte(need - 1) {}, Array.Empty(Of Byte)())
                        If need > 0 Then ReadExact(fs, packed, 0, need)

                        Dim outChunk As Byte()
                        If isComp Then
                            If hdr.Version >= 3UI AndAlso hdr.V3_Compression = 3UI Then
                                ' LZ4
                                outChunk = Lz4Strict.Lz4DecompressStrict(packed, CInt(ch.DecompressedSize))
                            Else
                                ' zlib/deflate
                                outChunk = ZlibStrict.ZlibDecompressStrict(packed, CInt(ch.DecompressedSize))
                            End If
                        Else
                            outChunk = packed
                        End If
                        If isComp AndAlso ch.DecompressedSize = 0UI Then
                            Throw New InvalidDataException("BA2.GNRL: chunk comprimido con DecompressedSize=0 no es válido.")
                        End If
                        If ch.DecompressedSize > 0UI AndAlso outChunk.Length <> CInt(ch.DecompressedSize) Then
                            ' ajustar o validar
                            If outChunk.Length < CInt(ch.DecompressedSize) Then
                                Throw New InvalidDataException($"BA2.GNRL: tamaño descomprimido menor al esperado (got={outChunk.Length}, exp={ch.DecompressedSize}).")
                            End If
                            Array.Resize(outChunk, CInt(ch.DecompressedSize))
                        End If

                        ms.Write(outChunk, 0, outChunk.Length)

                    Next

                    Return ms.ToArray()
                End Using
            End Function

        End Class

        ' ---------- DX10 ----------
        Private Class EntryDX10
            Inherits EntryBase

            Public HashFile As UInteger
            Public HashExt As UInteger
            Public HashDir As UInteger
            Public DataFileIndex As Byte
            Public ChunkCount As Byte
            Public ChunkHeaderSize As UShort

            Public Height As UShort
            Public Width As UShort
            Public MipCount As Byte
            Public DxgiFormatU8 As Byte
            Public Flags As Byte           ' bit0: cubemap
            Public TileMode As Byte

            Public Chunks As New List(Of Chunk)

            Public Structure Chunk
                Public Offset As ULong
                Public CompressedSize As UInteger
                Public DecompressedSize As UInteger
                Public MipFirst As UShort
                Public MipLast As UShort
                Public Sentinel As UInteger
            End Structure

            Public Overrides Function Extract(fs As Stream, hdr As Ba2Header) As Byte()
                Dim isCube As Boolean = ((Flags And 1) <> 0)
                Dim arraySize As Integer = If(isCube, 6, 1)
                If Chunks Is Nothing OrElse Chunks.Count = 0 Then
                    Return Array.Empty(Of Byte)()
                End If

                If TileMode <> 8 Then
                    Throw New NotSupportedException($"BA2.DX10: TileMode={TileMode} no soportado (requiere destileo).")
                End If

                Dim ddsHeader As Byte() = Loader.EncodeDDSHeader(CInt(DxgiFormatU8), CInt(Width), CInt(Height), arraySize, CInt(If(MipCount = 0, 1, MipCount)), isCube)

                ' Los chunks DX10 pueden no venir en orden lógico de mips; ordenamos por MipFirst asc
                Dim ordered = Chunks.OrderBy(Function(c) CInt(c.MipFirst)).ThenBy(Function(c) CLng(c.Offset)).ToList()

                Using ms As New MemoryStream()
                    ms.Write(ddsHeader, 0, ddsHeader.Length)

                    For Each ch In ordered
                        fs.Position = CLng(ch.Offset)
                        Dim isComp As Boolean = (ch.CompressedSize <> 0UI)
                        Dim needUL As ULong = If(isComp, CULng(ch.CompressedSize), CULng(ch.DecompressedSize))
                        If needUL > CULng(Integer.MaxValue) Then
                            Throw New InvalidDataException($"BA2.DX10: tamaño de chunk excede Int32 (need={needUL}).")
                        End If
                        Dim need As Integer = CInt(needUL)

                        Dim packed = If(need > 0, New Byte(need - 1) {}, Array.Empty(Of Byte)())
                        If need > 0 Then ReadExact(fs, packed, 0, need)

                        Dim outChunk As Byte()
                        If isComp Then
                            If hdr.Version >= 3UI AndAlso hdr.V3_Compression = 3UI Then
                                outChunk = Lz4Strict.Lz4DecompressStrict(packed, CInt(ch.DecompressedSize))
                            Else
                                outChunk = ZlibStrict.ZlibDecompressStrict(packed, CInt(ch.DecompressedSize))
                            End If
                        Else
                            outChunk = packed
                        End If
                        If isComp AndAlso ch.DecompressedSize = 0UI Then
                            Throw New InvalidDataException("BA2.DX10: chunk comprimido con DecompressedSize=0 no es válido.")
                        End If
                        If ch.DecompressedSize > 0UI AndAlso outChunk.Length <> CInt(ch.DecompressedSize) Then
                            If outChunk.Length < CInt(ch.DecompressedSize) Then
                                Throw New InvalidDataException($"BA2.DX10: tamaño descomprimido menor al esperado (got={outChunk.Length}, exp={ch.DecompressedSize}).")
                            End If
                            Array.Resize(outChunk, CInt(ch.DecompressedSize))
                        End If

                        ms.Write(outChunk, 0, outChunk.Length)
                    Next

                    Return ms.ToArray()
                End Using
            End Function

        End Class

        Public Sub New(fs As Stream, enc As Encoding)
            _fs = fs : _enc = enc
            _br = New BinaryReader(_fs, _enc, leaveOpen:=True)
            Open()
        End Sub

        Private Sub Open()
            ' Magic "BTDX"
            Dim magic = Encoding.ASCII.GetString(_br.ReadBytes(4))
            If magic <> "BTDX" Then Throw New InvalidDataException("BA2: magic inválido")

            _hdr.Version = _br.ReadUInt32()

            Dim typeStr = Encoding.ASCII.GetString(_br.ReadBytes(4))
            Select Case typeStr
                Case "GNRL" : _hdr.Type = Ba2Type.GNRL
                Case "DX10" : _hdr.Type = Ba2Type.DX10
                Case "GNMF" : _hdr.Type = Ba2Type.GNMF
                Case Else : Throw New InvalidDataException("BA2: tipo desconocido " & typeStr)
            End Select

            _hdr.NumFiles = _br.ReadUInt32()
            _hdr.NameTableOffset = _br.ReadUInt64()

            If _hdr.Version = 2UI OrElse _hdr.Version = 3UI Then
                _hdr.V2V3_UnknownAlways1 = _br.ReadUInt64()
                If _hdr.Version = 3UI Then _hdr.V3_Compression = _br.ReadUInt32()
            End If

            _entries.Capacity = CInt(_hdr.NumFiles)

            For i = 0UI To _hdr.NumFiles - 1UI
                Dim hashFile = _br.ReadUInt32()
                Dim hashExt = _br.ReadUInt32()
                Dim hashDir = _br.ReadUInt32()
                Dim dataFileIdx = _br.ReadByte()
                Dim chunkCount = _br.ReadByte()
                Dim chunkHdrSize = _br.ReadUInt16()


                Select Case _hdr.Type
                    Case Ba2Type.GNRL
                        If chunkHdrSize <> CUShort(16) AndAlso chunkHdrSize <> CUShort(20) Then
                            Throw New NotSupportedException($"BA2.GNRL: chunk header size inesperado ({chunkHdrSize}).")
                        End If
                        Dim e As New EntryGNRL With {
                            .Index = CInt(i),
                            .HashFile = hashFile, .HashExt = hashExt, .HashDir = hashDir,
                            .DataFileIndex = dataFileIdx, .ChunkCount = chunkCount, .ChunkHeaderSize = chunkHdrSize
                        }
                        For c = 0 To e.ChunkCount - 1
                            Dim ch As New EntryGNRL.Chunk
                            ch.Offset = _br.ReadUInt64()           ' 8
                            ch.CompressedSize = _br.ReadUInt32()   ' +4 = 12
                            ch.DecompressedSize = _br.ReadUInt32() ' +4 = 16

                            ' Sentinel SIEMPRE va inmediatamente después del header base.
                            Dim s As UInteger = _br.ReadUInt32()   ' +4 = 20 consumidos
                            If s <> &HBAADF00DUI Then
                                Throw New InvalidDataException($"BA2.GNRL: sentinel inválido (0x{s:X8}).")
                            End If
                            ch.Sentinel = s
                            e.Chunks.Add(ch)
                        Next
                        e.FileName = $"{e.HashFile:X8}.bin"
                        If e.HashDir <> 0UI Then e.Directory = $"{e.HashDir:X8}"
                        _entries.Add(e)

                    Case Ba2Type.DX10
                        If chunkHdrSize <> CUShort(24) AndAlso chunkHdrSize <> CUShort(20) Then
                            Throw New NotSupportedException($"BA2.DX10: chunk header size inesperado ({chunkHdrSize}).")
                        End If
                        Dim e As New EntryDX10 With {
                            .Index = CInt(i),
                            .HashFile = hashFile, .HashExt = hashExt, .HashDir = hashDir,
                            .DataFileIndex = dataFileIdx, .ChunkCount = chunkCount, .ChunkHeaderSize = chunkHdrSize
                        }
                        e.Height = _br.ReadUInt16()
                        e.Width = _br.ReadUInt16()
                        e.MipCount = _br.ReadByte()
                        e.DxgiFormatU8 = _br.ReadByte()
                        e.Flags = _br.ReadByte()
                        e.TileMode = _br.ReadByte()

                        For c = 0 To e.ChunkCount - 1
                            Dim ch As New EntryDX10.Chunk
                            ch.Offset = _br.ReadUInt64()           ' 8
                            ch.CompressedSize = _br.ReadUInt32()   ' +4 = 12
                            ch.DecompressedSize = _br.ReadUInt32() ' +4 = 16
                            ch.MipFirst = _br.ReadUInt16()         ' +2 = 18
                            ch.MipLast = _br.ReadUInt16()          ' +2 = 20

                            Dim s As UInteger = &HBAADF00DUI

                            If chunkHdrSize = CUShort(24) Then
                                ' Sentinel forma parte del header “declarado”
                                s = _br.ReadUInt32()               ' +4 = 24
                                If s <> &HBAADF00DUI Then
                                    Throw New InvalidDataException($"BA2.DX10: sentinel inválido (0x{s:X8}).")
                                End If
                            Else
                                ' chunkHdrSize=20: algunos siguen poniendo el sentinel inmediatamente después.
                                ' Lo detectamos sin romper el alineamiento.
                                Dim pos = _fs.Position
                                Dim maybe = _br.ReadUInt32()
                                If maybe = &HBAADF00DUI Then
                                    s = maybe
                                Else
                                    _fs.Position = pos ' no había sentinel; retroceder
                                End If
                            End If

                            ch.Sentinel = s
                            e.Chunks.Add(ch)
                        Next

                        e.FileName = $"{e.HashFile:X8}.dds"
                        If e.HashDir <> 0UI Then e.Directory = $"{e.HashDir:X8}"
                        _entries.Add(e)

                    Case Else
                        Throw New NotSupportedException("BA2: GNMF fuera de alcance")
                End Select
            Next

            ' Name table (como lo tenías: Int16 + bytes usando _enc)
            If _hdr.NameTableOffset > 0UL Then
                _fs.Position = CLng(_hdr.NameTableOffset)
                For i = 0 To _entries.Count - 1
                    Dim nlen As UShort = _br.ReadUInt16()
                    Dim nb = If(nlen > 0US, _br.ReadBytes(nlen), Array.Empty(Of Byte)())
                    Dim full = PathUtil.NormalizeSlash(_enc.GetString(nb))
                    Dim slash = full.LastIndexOf(Correct_Path_separator)
                    If slash >= 0 Then
                        _entries(i).Directory = full.Substring(0, slash)
                        _entries(i).FileName = full.Substring(slash + 1)
                    Else
                        _entries(i).Directory = ""
                        _entries(i).FileName = full
                    End If
                Next
            End If
        End Sub

        Public Function ListEntries() As List(Of ArchiveEntry) Implements IArchiveImpl.ListEntries
            Dim list As New List(Of ArchiveEntry)(_entries.Count)
            For Each e In _entries
                list.Add(New ArchiveEntry With {.Index = e.Index, .Directory = e.Directory, .FileName = e.FileName})
            Next
            Return list
        End Function

        Public Function ExtractByIndex(index As Integer) As Byte() Implements IArchiveImpl.ExtractByIndex
            Return _entries(index).Extract(_fs, _hdr)
        End Function

        Public Sub Dispose() Implements IArchiveImpl.Dispose
            _br?.Dispose()
        End Sub

        Private Shared Sub ReadExact(s As Stream, buf As Byte(), offset As Integer, count As Integer)
            Dim read As Integer = 0
            While read < count
                Dim n = s.Read(buf, offset + read, count - read)
                If n <= 0 Then Throw New EndOfStreamException($"Faltan {count - read} bytes.")
                read += n
            End While
        End Sub


    End Class
End Namespace

