Option Strict On
Imports System.IO
Imports System.Linq
Imports System.Text
Imports K4os.Compression.LZ4.Streams

Namespace BethesdaArchive.Core
    Public NotInheritable Class BsaWriter
        Public Shared Event Writed()
        Friend Class BASTes4Hashing
            Private Shared Function GetBytesLatin1(text As String) As Byte()
                If text Is Nothing OrElse text.Length = 0 Then Return Array.Empty(Of Byte)()
                Return Encoding.Latin1.GetBytes(text)
            End Function

            ' CRC estilo 0x1003F con aritmética en 64 bits y máscara a 32
            Private Shared Function Crc1003F(bytes As Byte()) As UInteger
                Dim crc As ULong = 0UL
                If bytes IsNot Nothing Then
                    For i As Integer = 0 To bytes.Length - 1
                        crc = (crc * &H1003FUL + CULng(bytes(i))) And &HFFFFFFFFUL
                    Next
                End If
                Return CUInt(crc)
            End Function

            Friend Shared Function Tes4HashDirectoryBytes(path As String) As Byte()
                path = PathUtil.NormalizeSlash(path).Trim(Correct_Path_separator).ToLowerInvariant()
                ' SSE: si tras normalizar queda vacío, el hash se calcula sobre "."
                If String.IsNullOrEmpty(path) Then path = "."
                Dim b = GetBytesLatin1(path)
                Dim last As Byte = If(b.Length >= 1, b(b.Length - 1), CByte(0))
                Dim last2 As Byte = If(b.Length >= 2, b(b.Length - 2), CByte(0))
                Dim first As Byte = If(b.Length >= 1, b(0), CByte(0))
                Dim length As Byte = CByte(Math.Min(b.Length, 255))

                Dim crc As UInteger = 0UI
                ' Excluir primer byte y los dos últimos → índices 1..(Length-3)
                If b.Length > 3 Then
                    Dim sliceLen As Integer = b.Length - 3
                    Dim tmp(sliceLen - 1) As Byte
                    Array.Copy(b, 1, tmp, 0, sliceLen)
                    crc = Crc1003F(tmp)
                End If

                Dim out(7) As Byte
                out(0) = last : out(1) = last2 : out(2) = length : out(3) = first
                Dim c = BitConverter.GetBytes(crc) ' LE
                Buffer.BlockCopy(c, 0, out, 4, 4)
                Return out
            End Function

            Friend Shared Function Tes4HashFileBytes(path As String) As Byte()
                path = PathUtil.NormalizeSlash(path).ToLowerInvariant()
                ' nombre sin carpeta
                Dim posSlash = path.LastIndexOf(Correct_Path_separator)
                If posSlash >= 0 Then path = path.Substring(posSlash + 1)
                Dim stem As String = path
                Dim ext As String = ""
                Dim p = path.LastIndexOf("."c)
                If p >= 0 Then
                    stem = path.Substring(0, p)
                    ext = path.Substring(p)
                End If

                If stem.Length = 0 OrElse stem.Length >= 260 OrElse ext.Length >= 16 Then
                    Return New Byte(7) {} ' cero
                End If

                Dim h = Tes4HashDirectoryBytes(stem) ' << mantiene tu función actual

                Dim crcBase As UInteger = BitConverter.ToUInt32(h, 4)
                Dim add As UInteger = Crc1003F(GetBytesLatin1(ext))
                Dim sum As ULong = CULng(crcBase) + CULng(add)
                crcBase = CUInt(sum And &HFFFFFFFFUL)
                Dim crcBytes = BitConverter.GetBytes(crcBase)
                Buffer.BlockCopy(crcBytes, 0, h, 4, 4)

                ' LUT de extensiones (mismos ajustes de bytes que tenías)
                Dim lut = New String() {"", ".nif", ".kf", ".dds", ".wav", ".adp"}
                Dim i As Integer = Array.IndexOf(lut, ext)
                If i >= 0 Then
                    Dim first = CInt(h(3))
                    Dim last = CInt(h(0))
                    Dim last2 = CInt(h(1))
                    first = (first + 32 * (i And &HFC)) And &HFF
                    last = (last + ((i And &HFE) << 6)) And &HFF
                    last2 = (last2 + (i << 7)) And &HFF
                    h(3) = CByte(first) : h(0) = CByte(last) : h(1) = CByte(last2)
                End If

                Return h
            End Function


        End Class
        <Flags>
        Private Enum ArchiveFlags As UInteger
            IncludeDirectoryNames = 1UI   ' bit 0
            IncludeFileNames = 2UI        ' bit 1
            Compressed = 4UI              ' bit 2 (global)
            EmbedFileNames = &H100UI      ' bit 8
        End Enum

        Private Const VERSION As UInteger = &H69UI                  ' SSE v105
        Private Const HEADER_SIZE As UInteger = &H24UI              ' 36
        Private Const DIR_ENTRY_SIZE_V105 As Integer = &H18         ' 24
        Private Const FILE_ENTRY_SIZE As Integer = &H10             ' 16
        Private Const ICOMPRESSION As UInteger = 1UI << 30          ' bit30 (invert per-file)

        Public NotInheritable Class Options
            Public Property Encoding As Encoding = Encoding.UTF8
            Public Property UseDirectoryStrings As Boolean = True
            Public Property UseFileStrings As Boolean = True
            Public Property EmbedNames As Boolean = True
            Public Property GlobalCompressed As Boolean = True
        End Class

        '===================== PARTE 1: HEADER FIELDS + PAYLOAD BUILDER =====================
        Public Shared Sub Write(output As Stream, entriesIn As IEnumerable(Of VirtualEntry), opts As Options)
            If output Is Nothing OrElse Not output.CanWrite Then Throw New ArgumentException("Stream inválido.")
            If opts Is Nothing Then opts = New Options()
            Dim enc = If(opts.Encoding, Encoding.UTF8)

            ' ---- Normalización de entradas ----
            Dim entries = entriesIn?.ToList()
            If entries Is Nothing OrElse entries.Count = 0 Then Throw New InvalidOperationException("No hay entradas.")
            For Each e In entries
                e.Directory = PathUtil.NormalizeSlash(e.Directory)
                If String.IsNullOrEmpty(e.FileName) Then Throw New InvalidDataException("Entry sin FileName.")
                If e.Data Is Nothing Then e.Data = Array.Empty(Of Byte)()
            Next

            ' ---- Agrupar por carpeta (orden alfabético) ----
            Dim groups = entries.GroupBy(Function(e) e.Directory, StringComparer.OrdinalIgnoreCase).OrderBy(Function(g)
                                                                                                                Dim h = BASTes4Hashing.Tes4HashDirectoryBytes(g.Key)
                                                                                                                Return BitConverter.ToUInt64(h, 0) ' LE → 64-bit
                                                                                                            End Function).ToList()

            Dim folderCount As Integer = groups.Count
            Dim fileCount As Integer = entries.Count

            ' ---- Sumas "raw" para header ----
            ' FolderNameLength (campo del header) debe ser: sum(rawLenDir) + folderCount
            ' (tu lector calcula dirStrSz = FolderNameLength + FolderCount → da raw + 2*folders = bytes reales de BZString)
            Dim folderNamesRawSum As Integer = groups.Sum(Function(g) enc.GetByteCount(If(g.Key, "")))
            Dim fileNamesRawSum As Integer = entries.Sum(Function(e) enc.GetByteCount(e.FileName))

            ' ---- Layout de secciones (sin escribir todavía) ----
            Dim posHeader As Long = 0
            Dim posDirEntries As Long = posHeader + HEADER_SIZE
            Dim posAfterDirEntries As Long = posDirEntries + CLng(DIR_ENTRY_SIZE_V105) * folderCount

            Dim hasDirStrings As Boolean = opts.UseDirectoryStrings
            Dim hasFileStrings As Boolean = opts.UseFileStrings

            ' BZStrings de carpetas: bytes reales = rawLen + 2 (len+1 + NUL) por carpeta
            Dim dirBzTotal As Long = 0
            Dim dirBzLenByDir As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            If hasDirStrings Then
                For Each g In groups
                    Dim raw = enc.GetBytes(If(g.Key, ""))
                    If raw.Length > 255 Then Throw New InvalidDataException("Directory BZString > 255.")
                    Dim totalThis = raw.Length + 2
                    dirBzLenByDir(g.Key) = totalThis
                    dirBzTotal += totalThis
                Next
            End If

            ' Región "filesOffset" (como la usa tu BsaImpl.Open): [por carpeta: BZString?] + [N * 16]
            Dim filesOffset As Long = posAfterDirEntries
            Dim fileEntriesBytes As Long = CLng(FILE_ENTRY_SIZE) * fileCount
            Dim posFileNameStrings As Long = filesOffset + dirBzTotal + fileEntriesBytes

            ' Tabla ZSTRING real = sum(rawLen) + fileCount (por los NULs)
            Dim fileStringsBytesActual As Long = If(hasFileStrings, CLng(fileNamesRawSum) + fileCount, 0L)
            Dim posFileData As Long = posFileNameStrings + fileStringsBytesActual

            ' ---- Preparar HEADER fields (no se escriben aún; los usamos en Parte 2) ----
            Dim header_Flags As ArchiveFlags = CType(0UI, ArchiveFlags)
            If hasDirStrings Then header_Flags = header_Flags Or ArchiveFlags.IncludeDirectoryNames
            If hasFileStrings Then header_Flags = header_Flags Or ArchiveFlags.IncludeFileNames
            If opts.GlobalCompressed Then header_Flags = header_Flags Or ArchiveFlags.Compressed
            If opts.EmbedNames Then header_Flags = header_Flags Or ArchiveFlags.EmbedFileNames

            Dim header_FolderCount As UInteger = CUInt(folderCount)
            Dim header_FileCount As UInteger = CUInt(fileCount)
            Dim header_FolderNameLength As UInteger = If(hasDirStrings, CUInt(folderNamesRawSum + folderCount), 0UI)
            Dim header_FileNameLength As UInteger = If(hasFileStrings, CUInt(fileNamesRawSum + fileCount), 0UI)

            ' ---- Construcción de PAYLOADS (con BSTRING si corresponde) + SizeField ----
            ' Guardamos por key lógico → (SizeField, OffsetAbs=0 temporal, Payload)
            Dim recByKey As New Dictionary(Of String, (SizeField As UInteger, OffsetAbs As UInteger, Payload As Byte()))(StringComparer.OrdinalIgnoreCase)

            For Each e In entries
                Using ms As New MemoryStream()
                    Dim raw As Byte() = e.Data

                    ' === Decide compresión efectiva (override vs global) ===
                    ' - Si PreferCompress=True y global=False → compress
                    ' - Si PreferCompress=False y global=True → no compress
                    ' - Sino, seguir global.
                    Dim wantCompressed As Boolean = opts.GlobalCompressed
                    If e.PreferCompress AndAlso Not opts.GlobalCompressed Then wantCompressed = True
                    If (Not e.PreferCompress) AndAlso opts.GlobalCompressed Then wantCompressed = False

                    ' === [BSTRING embebido] u8 len + (dir + '\' + file) (SIN NUL) ===
                    Dim namesBytes As Integer = 0
                    Dim dirBytes() As Byte = Array.Empty(Of Byte)()
                    Dim fileBytes() As Byte = Array.Empty(Of Byte)()
                    If opts.EmbedNames Then
                        Dim dirTxt = If(e.Directory, "").Trim(Correct_Path_separator)
                        dirBytes = enc.GetBytes(dirTxt)
                        fileBytes = enc.GetBytes(e.FileName)

                        ' SSE/C++: SIEMPRE incluir el separador '\' aunque la carpeta esté vacía (raíz)
                        Dim bstrlen As Integer = dirBytes.Length + 1 + fileBytes.Length
                        If bstrlen > 255 Then
                            Throw New InvalidDataException($"BSTRING > 255: '{dirTxt}\{e.FileName}'")
                        End If
                        ms.WriteByte(CByte(bstrlen))
                        If dirBytes.Length > 0 Then ms.Write(dirBytes, 0, dirBytes.Length)
                        ms.WriteByte(&H5C) ' '\'
                        If fileBytes.Length > 0 Then ms.Write(fileBytes, 0, fileBytes.Length)

                        namesBytes = 1 + bstrlen
                        End If


                        ' === Parte de DATOS (lo que hay "en disco"): si comprimido → u32 decompSize + LZ4 frame; si no → raw ===
                        Dim posAfterNames As Long = ms.Position
                    If wantCompressed Then
                        ' u32 LE con tamaño descomprimido + LZ4F frame
                        Dim u = BitConverter.GetBytes(CUInt(raw.Length))
                        ms.Write(u, 0, 4)

                        ' Ajustes de LZ4 para asemejarse al writer de referencia (HC default, sin content length/checksums).
                        Using lz As Stream = LZ4Stream.Encode(ms,
                    New LZ4EncoderSettings() With {
                        .CompressionLevel = K4os.Compression.LZ4.LZ4Level.L04_HC, ' ~ default HC
                        .ChainBlocks = True,
                        .ContentChecksum = False,
                        .BlockSize = 4 * 1024 * 1024,
                        .BlockChecksum = False,
                        .ContentLength = Nothing
                    },
                    leaveOpen:=True)
                            If raw.Length > 0 Then lz.Write(raw, 0, raw.Length)
                        End Using
                    Else
                        If raw.Length > 0 Then ms.Write(raw, 0, raw.Length)
                    End If

                    Dim payload As Byte() = ms.ToArray()

                    ' === sizeField (lo que "informan" los lectores) = bytes en disco del BLOQUE ===
                    ' fsize = (Embed? 1+len(dir)+1+len(file) : 0) + (Comprimido? 4 + LZ4FrameLen : rawLen)
                    Dim dataBytes As Integer = CInt(ms.Length - posAfterNames) ' incluye el +4 del decomp size si comprimido
                    Dim sizeNoFlags As UInteger = CUInt(namesBytes + dataBytes)
                    If sizeNoFlags > &H3FFFFFFFUI Then Throw New OverflowException($"BSA: sizeField demasiado grande ({sizeNoFlags}).")

                    Dim sizeField As UInteger = sizeNoFlags
                    ' bit30 override si la compresión efectiva difiere de la global
                    If (wantCompressed Xor opts.GlobalCompressed) Then
                        sizeField = sizeField Or ICOMPRESSION
                    End If

                    ' Guardar registro (Size, OffsetAbs=0 temporal, Payload a escribir)
                    recByKey(PathUtil.JoinDirFile(e.Directory, e.FileName)) =
                (sizeField, 0UI, payload)
                End Using
                RaiseEvent Writed()
            Next

            ' ===================== PARTE 2: OFFSET ASSIGN + ESCRITURA HEADER/ENTRIES =====================

            ' Orden estable: primero por carpeta, luego por nombre de archivo
            Dim orderedPairs As New List(Of (Dir As String, Name As String))(fileCount)
            For Each g In groups
                For Each e In g.OrderBy(Function(x) BitConverter.ToUInt64(BASTes4Hashing.Tes4HashFileBytes(x.FileName), 0))
                    orderedPairs.Add((g.Key, e.FileName))
                Next
            Next

            ' Asignar OffsetAbs a cada payload (zona DATA empieza en posFileData)
            Dim runData As Long = posFileData
            For Each p In orderedPairs
                Dim key = PathUtil.JoinDirFile(p.Dir, p.Name)
                Dim rec = recByKey(key)
                If runData < 0 OrElse runData > UInteger.MaxValue Then Throw New OverflowException("Offset > 4GB.")
                rec.OffsetAbs = CUInt(runData)
                recByKey(key) = rec
                ' Avanzar exactamente lo que ocupa el BLOQUE en disco (BSTRING + datos [+4 si comp]), sin padding.
                runData += If(rec.Payload Is Nothing, 0, rec.Payload.Length)
            Next

            ' Calcular inicio del bloque de carpeta (BZString) y el primer file-entry
            Dim offsetFolderBlockByDir As New Dictionary(Of String, UInteger)(StringComparer.OrdinalIgnoreCase)
            Dim offsetFileEntryByDir As New Dictionary(Of String, UInteger)(StringComparer.OrdinalIgnoreCase)

            Dim cursor As Long = filesOffset
            For Each g In groups
                Dim bzLen As Integer = If(hasDirStrings, dirBzLenByDir(g.Key), 0)

                If cursor < 0 OrElse cursor > UInteger.MaxValue Then Throw New OverflowException("Offset carpeta > 4GB.")
                offsetFolderBlockByDir(g.Key) = CUInt(cursor)          ' inicio del BZString de la carpeta

                Dim firstEntry = cursor + bzLen
                If firstEntry < 0 OrElse firstEntry > UInteger.MaxValue Then Throw New OverflowException("Offset file-entry > 4GB.")
                offsetFileEntryByDir(g.Key) = CUInt(firstEntry)        ' (por si lo necesitás en otras partes)

                cursor += bzLen + CLng(FILE_ENTRY_SIZE) * g.Count()
            Next

            ' ========= HEADER =========
            output.Position = posHeader
            output.Write(BitConverter.GetBytes(&H415342UI), 0, 4)          ' "BSA\0"
            output.Write(BitConverter.GetBytes(VERSION), 0, 4)             ' 0x69
            output.Write(BitConverter.GetBytes(HEADER_SIZE), 0, 4)         ' 0x24
            output.Write(BitConverter.GetBytes(CUInt(header_Flags)), 0, 4) ' ArchiveFlags
            output.Write(BitConverter.GetBytes(header_FolderCount), 0, 4)
            output.Write(BitConverter.GetBytes(header_FileCount), 0, 4)
            output.Write(BitConverter.GetBytes(header_FolderNameLength), 0, 4)
            output.Write(BitConverter.GetBytes(header_FileNameLength), 0, 4)
            output.Write(BitConverter.GetBytes(0UI), 0, 4)                 ' FileFlags (no usados)

            ' ========= DIRECTORY ENTRIES =========
            output.Position = posDirEntries
            For Each g In groups
                Dim dirHashBytes = BASTes4Hashing.Tes4HashDirectoryBytes(g.Key)
                output.Write(dirHashBytes, 0, dirHashBytes.Length)                  ' 8 bytes
                output.Write(BitConverter.GetBytes(CUInt(g.Count())), 0, 4)         ' u32 count
                output.Write(BitConverter.GetBytes(0UI), 0, 4)                      ' unused

                ' *** CAMBIO: offset debe incluir totalFileNameLength del HEADER ***
                Dim firstEntryAbs As UInteger = offsetFileEntryByDir(g.Key)
                Dim offsetField As UInteger = offsetFolderBlockByDir(g.Key) + header_FileNameLength
                output.Write(BitConverter.GetBytes(offsetField), 0, 4)              ' u32 offset (con +FileNameLength)

                output.Write(BitConverter.GetBytes(0UI), 0, 4)                      ' unused
            Next

            ' ========= (INTERCALADO) BZStrings de carpeta + FILE ENTRIES =========
            output.Position = filesOffset
            For Each g In groups
                ' BZString de directorio
                If hasDirStrings Then
                    Dim raw = enc.GetBytes(If(g.Key, ""))
                    output.WriteByte(CByte(raw.Length + 1))     ' incluye +1 para NUL
                    If raw.Length > 0 Then output.Write(raw, 0, raw.Length)
                    output.WriteByte(0)                         ' NUL
                End If

                ' File entries de esta carpeta
                For Each e In g.OrderBy(Function(x) BitConverter.ToUInt64(BASTes4Hashing.Tes4HashFileBytes(x.FileName), 0))
                    Dim key = PathUtil.JoinDirFile(g.Key, e.FileName)
                    Dim rec = recByKey(key)

                    Dim fileHashBytes = BASTes4Hashing.Tes4HashFileBytes(e.FileName) ' hash del nombre
                    output.Write(fileHashBytes, 0, fileHashBytes.Length)          ' 8 bytes
                    output.Write(BitConverter.GetBytes(rec.SizeField), 0, 4)      ' u32 sizeField (fsize)
                    output.Write(BitConverter.GetBytes(rec.OffsetAbs), 0, 4)      ' u32 offset ABS al bloque en DATA
                Next
            Next

            ' ========= File name strings (zstring) =========
            output.Position = posFileNameStrings
            If hasFileStrings Then
                For Each g In groups
                    For Each e In g.OrderBy(Function(x) BitConverter.ToUInt64(BASTes4Hashing.Tes4HashFileBytes(x.FileName), 0))
                        Dim raw = enc.GetBytes(e.FileName)
                        If raw.Length > 0 Then output.Write(raw, 0, raw.Length)
                        output.WriteByte(0) ' NUL
                    Next
                Next
            End If

            ' ===================== PARTE 3: ESCRITURA DE DATA (PAYLOADS) =====================

            ' Al terminar los zstrings, debemos estar exactamente donde empieza la DATA
            Dim expectedPos As Long = posFileData
            If output.Position <> expectedPos Then
                Throw New InvalidDataException($"Desalineado: fin de tabla de nombres={output.Position}, esperado posFileData={expectedPos}.")
            End If

            ' DATA: escribir cada bloque en el mismo orden usado para asignar OffsetAbs
            output.Position = posFileData
            For Each p In orderedPairs
                Dim key = PathUtil.JoinDirFile(p.Dir, p.Name)
                Dim rec = recByKey(key)
                If rec.Payload IsNot Nothing AndAlso rec.Payload.Length > 0 Then
                    output.Write(rec.Payload, 0, rec.Payload.Length)
                End If
            Next
        End Sub


    End Class

End Namespace



