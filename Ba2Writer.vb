Option Strict On
Imports System.IO
Imports System.Text
Imports K4os.Compression.LZ4.Streams
Imports K4os.Compression.LZ4

Namespace BethesdaArchive.Core


    Public Class Ba2WriterCommon

        ' ====== Presets ZLIB, equivalentes a cabeceras típicas ======
        Public Enum ZlibPreset
            Fast = 0        ' cabecera 0x78 0x01 ; DeflateStream Fastest
            [Default] = 1   ' cabecera 0x78 0x9C ; DeflateStream Optimal  (COMPORTAMIENTO ACTUAL)
            Max = 2         ' cabecera 0x78 0xDA ; DeflateStream Optimal (sesgo a mejor compresión)
        End Enum
        ' ====== Formatos de compresión admitidos por BA2 ======
        Public Enum CompressionFormat
            Zip = 0   ' zlib/deflate (FO4, FO76, SF v2)
            Lz4 = 3   ' LZ4 raw block (solo BA2 v3)
        End Enum
        ' ====== Hashing BA2 exacto ======
        Private Shared ReadOnly Crc32Table As UInteger() = BuildCrc32Table()
        Friend Shared Sub Fo4SplitPath(norm As String, ByRef parent As String, ByRef stem As String, ByRef extNoDot As String)
            ' Reimplementado usando PathUtil.SplitDirFile para unificar lógica de split.
            ' 1) Normalizar y recortar separadores extremos
            Dim relSlash As String = PathUtil.NormalizeSlash(norm).Trim(Correct_Path_separator, InCorrect_Path_separator)
            ' 2) Dividir en dir / archivo
            Dim dir As String = "", fileOnly As String = ""
            PathUtil.SplitDirFile(relSlash, dir, fileOnly)
            ' 3) Parent con separador FO4 (Correct_Path_separator)
            parent = dir.Replace(InCorrect_Path_separator, Correct_Path_separator)
            ' 4) stem y ext (sin el punto)
            stem = System.IO.Path.GetFileNameWithoutExtension(fileOnly)
            Dim ext As String = System.IO.Path.GetExtension(fileOnly)
            If Not String.IsNullOrEmpty(ext) AndAlso ext.StartsWith("."c) Then
                extNoDot = ext.Substring(1)
            Else
                extNoDot = ""
            End If
        End Sub

        Friend Shared Function Crc32Ascii(text As String) As UInteger
            Dim bytes = Encoding.ASCII.GetBytes(text)
            Dim crc As UInteger = &HFFFFFFFFUI
            For Each bt In bytes
                Dim idx = (crc Xor bt) And &HFFUI
                crc = (crc >> 8) Xor Crc32Table(CInt(idx))
            Next
            Return Not crc
        End Function

        Friend Shared Function PackExt(ext As String) As UInteger
            ext = ext.ToLowerInvariant()
            Dim b0 As Byte = 0, b1 As Byte = 0, b2 As Byte = 0, b3 As Byte = 0
            If ext.Length > 0 Then b0 = CByte(AscW(ext(0)) And &HFF)
            If ext.Length > 1 Then b1 = CByte(AscW(ext(1)) And &HFF)
            If ext.Length > 2 Then b2 = CByte(AscW(ext(2)) And &HFF)
            If ext.Length > 3 Then b3 = CByte(AscW(ext(3)) And &HFF)
            Return CUInt(b0 Or (CUInt(b1) << 8) Or (CUInt(b2) << 16) Or (CUInt(b3) << 24))
        End Function
        Friend Shared Function BuildCrc32Table() As UInteger()
            Const poly As UInteger = &HEDB88320UI
            Dim t(255) As UInteger
            For i = 0 To 255
                Dim c As UInteger = CUInt(i)
                For j = 0 To 7
                    If (c And 1UI) <> 0UI Then
                        c = poly Xor (c >> 1)
                    Else
                        c >>= 1
                    End If
                Next
                t(i) = c
            Next
            Return t
        End Function

        ' ====== helpers ======
        Friend Shared Sub WriteAscii(s As Stream, txt As String)
            Dim b = Encoding.ASCII.GetBytes(txt)
            s.Write(b, 0, b.Length)
        End Sub
        Friend Shared Sub WriteU16(s As Stream, v As UShort)
            Dim b = BitConverter.GetBytes(v) : s.Write(b, 0, 2)
        End Sub
        Friend Shared Sub WriteU32(s As Stream, v As UInteger)
            Dim b = BitConverter.GetBytes(v) : s.Write(b, 0, 4)
        End Sub
        Friend Shared Sub WriteU64(s As Stream, v As ULong)
            Dim b = BitConverter.GetBytes(v) : s.Write(b, 0, 8)
        End Sub

        ' LZ4 RAW (bloque "naked", sin frame) compatible con BA2 v3
        Friend Shared Function CompressLz4(data As Byte()) As Byte()
            If data Is Nothing OrElse data.Length = 0 Then Return Array.Empty(Of Byte)()
            Dim srcLen As Integer = data.Length
            Dim maxOut As Integer = LZ4Codec.MaximumOutputSize(srcLen)
            Dim dst As Byte() = New Byte(maxOut - 1) {}
            ' Nota: usar la sobrecarga sin 'level' para máxima compatibilidad entre versiones del paquete.
            Dim written As Integer = LZ4Codec.Encode(data, 0, srcLen, dst, 0, dst.Length)
            If written <= 0 Then Return Array.Empty(Of Byte)()
            If written = dst.Length Then Return dst
            ReDim Preserve dst(written - 1)
            Return dst
        End Function
        Friend Shared Function CompressLz4Frame(data As Byte()) As Byte()
            If data Is Nothing OrElse data.Length = 0 Then Return Array.Empty(Of Byte)()
            Using ms As New MemoryStream()
                Using lz As Stream = LZ4Stream.Encode(ms, New LZ4EncoderSettings() With {.CompressionLevel = K4os.Compression.LZ4.LZ4Level.L12_MAX}, leaveOpen:=True)
                    lz.Write(data, 0, data.Length)
                End Using
                Return ms.ToArray()
            End Using
        End Function


        ' ZLIB wrapper sobre DEFLATE raw de .NET
        ' Compat: overload sin preset (mantiene el comportamiento actual).
        Friend Shared Function CompressZlib(data As Byte()) As Byte()
            Return CompressZlib(data, ZlibPreset.Default)
        End Function

        ' ZLIB con preset (CMF/FLG típicos) y Adler-32 del original
        Friend Shared Function CompressZlib(data As Byte(), preset As ZlibPreset) As Byte()
            If data Is Nothing OrElse data.Length = 0 Then Return Array.Empty(Of Byte)()

            ' 1) Comprimir RAW DEFLATE con DeflateStream
            Dim deflate As Byte()
            Using ms As New MemoryStream()
                Dim level As IO.Compression.CompressionLevel
                Select Case preset
                    Case ZlibPreset.Fast
                        level = Compression.CompressionLevel.Fastest
                    Case ZlibPreset.Max
                        level = Compression.CompressionLevel.SmallestSize
                    Case Else
                        level = Compression.CompressionLevel.Optimal
                End Select

                Using ds As New IO.Compression.DeflateStream(ms, level, True)
                    ds.Write(data, 0, data.Length)
                End Using
                deflate = ms.ToArray()
                End Using

            ' 2) Cabecera ZLIB (CMF/FLG) según preset
            ' CMF = 0x78 (CINFO=7 → 32KB, CM=8 → DEFLATE). FLG ya preajustado para %31==0.
            Dim cmf As Byte = &H78
            Dim flg As Byte
            Select Case preset
                Case ZlibPreset.Fast : flg = &H1      ' 0x78 0x01
                Case ZlibPreset.Max : flg = &HDA     ' 0x78 0xDA
                Case Else : flg = &H9C     ' 0x78 0x9C (Default)
            End Select

            ' 3) Adler-32 del ORIGINAL (no del comprimido), big-endian
            Dim adler As UInteger = Adler32(data)

            ' 4) Armar ZLIB: [CMF][FLG][DEFLATE...][Adler32 big-endian]
            Using outMs As New MemoryStream()
                outMs.WriteByte(cmf)
                outMs.WriteByte(flg)
                outMs.Write(deflate, 0, deflate.Length)
                ' Adler-32 big-endian
                outMs.WriteByte(CByte((adler >> 24) And &HFFUI))
                outMs.WriteByte(CByte((adler >> 16) And &HFFUI))
                outMs.WriteByte(CByte((adler >> 8) And &HFFUI))
                outMs.WriteByte(CByte(adler And &HFFUI))
                Return outMs.ToArray()
            End Using
        End Function

        Private Shared Function Adler32(data As Byte()) As UInteger
            Const MOD_ADLER As UInteger = 65521UI
            Dim a As UInteger = 1UI
            Dim b As UInteger = 0UI
            For i As Integer = 0 To data.Length - 1
                a = (a + data(i)) Mod MOD_ADLER
                b = (b + a) Mod MOD_ADLER
            Next
            Return (b << 16) Or a
        End Function
        ' ====== Helper para escribir el "compression_format" v3 ======
        Friend Shared Sub WriteCompressionFormatV3(s As Stream, fmt As CompressionFormat)
            ' BA2 v3: 0 = zip, 3 = lz4
            Dim code As UInteger = If(fmt = CompressionFormat.Lz4, 3UI, 0UI)
            WriteU32(s, code)
        End Sub

    End Class

    ''' <summary>
    ''' Escritor de BA2 (Fallout 4) DX10 (texturas).
    ''' Estructura v3 (compression=3=LZ4).
    ''' </summary>
    ''' 
    Public NotInheritable Class Ba2WriterDX10
        Public Shared Event Writed()

        Public NotInheritable Class Options
            Public Property Encoding As Encoding = Encoding.UTF8
            ' FO4 admite v1, v7 y v8; default NG:
            Public Property Version As UInteger = 8UI
            ' FO4 no usa CompressionCode en DX10; se deja por compat pero se IGNORA.
            Public Property IncludeStrings As Boolean = True
            Public Property ZlibPreset As Ba2WriterCommon.ZlibPreset = Ba2WriterCommon.ZlibPreset.Default
            Public Property CompressionFormat As Ba2WriterCommon.CompressionFormat = Ba2WriterCommon.CompressionFormat.Zip

        End Class

        ''' <summary>
        ''' Construye el BA2 (DX10) completo en memoria.
        ''' </summary>
        Public Shared Function Build(entries As IEnumerable(Of VirtualEntry), opts As Options) As Byte()
            Using ms As New MemoryStream()
                Write(ms, entries, opts)
                Return ms.ToArray()
            End Using
        End Function

        ''' <summary>
        ''' Escribe un BA2 DX10 (FO4 v1/v7/v8) en el stream de salida.
        ''' Cada VirtualEntry debe tener: Data, DxgiFormat (0..255), Width, Height, MipCount, Faces, IsCubemap.
        ''' Compresión: ZLIB (RFC1950). Si comp >= raw => store (CompressedSize = 0).
        ''' </summary>
        Public Shared Sub Write(output As Stream, entries As IEnumerable(Of VirtualEntry), opts As Options)
            If output Is Nothing OrElse Not output.CanWrite Then Throw New ArgumentException("Stream inválido.")
            If entries Is Nothing Then Throw New ArgumentNullException(NameOf(entries))
            If opts Is Nothing Then opts = New Options()

            ' Validación versión/compresión (DX10 soporta v1,7,8 y también v2/v3 según C++; LZ4 solo v3)
            If opts.Version <> 1UI AndAlso opts.Version <> 2UI AndAlso opts.Version <> 3UI AndAlso opts.Version <> 7UI AndAlso opts.Version <> 8UI Then
                Throw New InvalidDataException("DX10 admite Version=1,2,3,7,8.")
            End If
            If opts.CompressionFormat = Ba2WriterCommon.CompressionFormat.Lz4 AndAlso opts.Version <> 3UI Then
                Throw New InvalidDataException("LZ4 solo válido con BA2 Version=3.")
            End If

            Dim enc = If(opts.Encoding, Encoding.UTF8)
            Dim list = entries.ToList()
            If list.Count = 0 Then Throw New ArgumentException("No hay entradas.")

            ' Normalizar rutas lógicas
            For Each ve In list
                ve.Directory = PathUtil.NormalizeSlash(ve.Directory)
            Next

            ' ====== Preparar metadata + payloads ======
            Dim perFile = New List(Of FileMeta)(list.Count)
            Dim iFile As Integer = 0
            For Each ve In list
                ' Validar metadata obligatoria
                If ve.Width <= 0 OrElse ve.Height <= 0 Then Throw New InvalidDataException("DX10: Width/Height inválidos.")
                If ve.MipCount <= 0 Then Throw New InvalidDataException("DX10: MipCount inválido.")
                If ve.Faces <= 0 Then Throw New InvalidDataException("DX10: Faces inválido.")
                If ve.DxgiFormat < 0 OrElse ve.DxgiFormat > 255 Then Throw New InvalidDataException("DX10: DxgiFormat inválido (0..255).")

                ' Hashes
                Dim relNorm As String = PathUtil.JoinDirFile(ve.Directory, ve.FileName)
                Dim norm = PathUtil.NormalizeSlash(relNorm)
                Dim parent As String = "", stem As String = "", extNoDot As String = ""
                Ba2WriterCommon.Fo4SplitPath(norm, parent, stem, extNoDot)

                Dim fm As New FileMeta()
                fm.Index = iFile
                fm.HashFile = Ba2WriterCommon.Crc32Ascii(stem)
                fm.HashDir = Ba2WriterCommon.Crc32Ascii(parent)
                fm.HashExt = Ba2WriterCommon.PackExt(extNoDot)
                fm.ModIndex = 0
                fm.ChunkCount = 1
                fm.ChunkHeaderSize = CUShort(24) ' DX10: 24 bytes (incluye sentinel)

                ' Header DX10 por archivo
                fm.Dx10_Width = CUShort(ve.Width)
                fm.Dx10_Height = CUShort(ve.Height)
                fm.Dx10_MipCount = CByte(ve.MipCount)
                fm.Dx10_DxgiFormatU8 = CByte(ve.DxgiFormat And &HFF)
                fm.Dx10_Flags = If(ve.IsCubemap, CByte(1), CByte(0))
                fm.Dx10_TileMode = CByte(8) ' como en el C++

                ' Datos
                Dim raw As Byte() = If(ve.Data, Array.Empty(Of Byte)())
                If raw.Length > Integer.MaxValue Then Throw New InvalidDataException("DX10: archivo demasiado grande.")
                fm.Chunk_DecompSize = CUInt(raw.Length)
                fm.Chunk_MipFirst = 0US
                fm.Chunk_MipLast = CUShort(ve.MipCount - 1)

                ' ZLIB bien formado (CMF/FLG + Adler-32). Si no mejora => store.
                ' Compresión según Version/CompressionFormat:
                ' - v3 + LZ4 => LZ4 raw
                ' - resto => ZLIB
                Dim comp As Byte()
                If opts.Version = 3UI AndAlso opts.CompressionFormat = Ba2WriterCommon.CompressionFormat.Lz4 Then
                    comp = Ba2WriterCommon.CompressLz4(raw)
                Else
                    comp = Ba2WriterCommon.CompressZlib(raw, opts.ZlibPreset)
                End If
                Dim useComp As Boolean = comp IsNot Nothing AndAlso comp.Length > 0 AndAlso comp.Length < raw.Length
                fm.Chunk_CompSize = CUInt(If(useComp, comp.Length, 0))
                fm.Chunk_CompData = If(useComp, comp, raw)

                fm.FileName = ve.FileName
                fm.Directory = ve.Directory

                perFile.Add(fm)
                iFile += 1
                RaiseEvent Writed()
            Next

            ' ====== Header BA2 (v1/v2/v3/v7/v8) ======
            Ba2WriterCommon.WriteAscii(output, "BTDX")
            Ba2WriterCommon.WriteU32(output, CUInt(opts.Version)) ' V1/V7/V8
            Ba2WriterCommon.WriteAscii(output, "DX10")
            Ba2WriterCommon.WriteU32(output, CUInt(perFile.Count))
            Dim posStringTableOffset As Long = output.Position
            Ba2WriterCommon.WriteU64(output, 0UL)                 ' NameTableOffset (patch)
            ' v2/v3: campo extra u64(1)
            If opts.Version = 2UI OrElse opts.Version = 3UI Then
                Ba2WriterCommon.WriteU64(output, 1UL)
            End If
            ' v3: u32 compression_format (0=zip, 3=lz4)
            If opts.Version = 3UI Then
                Ba2WriterCommon.WriteCompressionFormatV3(output, opts.CompressionFormat)
            End If

            ' ====== File headers + chunk entries (DX10) ======
            Dim patchOffsetPositions As New List(Of Long)(perFile.Count)
            For Each fm In perFile
                ' 1) Hashes
                Ba2WriterCommon.WriteU32(output, fm.HashFile)
                Ba2WriterCommon.WriteU32(output, fm.HashExt)
                Ba2WriterCommon.WriteU32(output, fm.HashDir)

                ' 2) Mini header por archivo
                output.WriteByte(fm.ModIndex)                     ' 0
                output.WriteByte(fm.ChunkCount)                   ' 1
                Ba2WriterCommon.WriteU16(output, fm.ChunkHeaderSize) ' 24 (DX10)

                ' header_t DX10 (orden FO4: h, w, mip_count, format, flags, tile_mode)
                Ba2WriterCommon.WriteU16(output, fm.Dx10_Height)
                Ba2WriterCommon.WriteU16(output, fm.Dx10_Width)
                output.WriteByte(fm.Dx10_MipCount)
                output.WriteByte(fm.Dx10_DxgiFormatU8)
                output.WriteByte(fm.Dx10_Flags)
                output.WriteByte(fm.Dx10_TileMode)

                ' 3) Chunk header (24 bytes, sentinel incluido)
                patchOffsetPositions.Add(output.Position)         ' apunta al u64 Offset
                Ba2WriterCommon.WriteU64(output, 0UL)             ' Offset absoluto (patch)
                Ba2WriterCommon.WriteU32(output, fm.Chunk_CompSize) ' 0 si store
                Ba2WriterCommon.WriteU32(output, fm.Chunk_DecompSize)
                Ba2WriterCommon.WriteU16(output, fm.Chunk_MipFirst)
                Ba2WriterCommon.WriteU16(output, fm.Chunk_MipLast)
                Ba2WriterCommon.WriteU32(output, &HBAADF00DUI)    ' sentinel dentro de los 24
            Next

            ' ====== Data chunks + parcheo offsets ======
            For i = 0 To perFile.Count - 1
                Dim fm = perFile(i)
                Dim off As ULong = CULng(output.Position)

                ' Patch del offset absoluto
                Dim save = output.Position
                output.Position = patchOffsetPositions(i)
                Ba2WriterCommon.WriteU64(output, off)
                output.Position = save

                ' Escribir payload (comp o raw según política)
                Dim cd = fm.Chunk_CompData
                If cd IsNot Nothing AndAlso cd.Length > 0 Then
                    output.Write(cd, 0, cd.Length)
                End If
            Next

            ' ====== String table (UTF-8, u16 length) opcional ======
            If opts.IncludeStrings Then
                Dim stringTableOffset As ULong = CULng(output.Position)
                For Each fm In perFile
                    Dim full = PathUtil.JoinDirFile(fm.Directory, fm.FileName)
                    Dim rawName = enc.GetBytes(full)
                    If rawName.Length > UShort.MaxValue Then Throw New InvalidDataException("Nombre demasiado largo para string table (u16).")
                    Ba2WriterCommon.WriteU16(output, CUShort(rawName.Length))
                    output.Write(rawName, 0, rawName.Length)
                Next
                Dim savePos As Long = output.Position
                output.Position = posStringTableOffset
                Ba2WriterCommon.WriteU64(output, stringTableOffset)
                output.Position = savePos
            Else
                ' Omitir string table → dejar NameTableOffset=0 (ya escrito) y no escribir nombres.
            End If

        End Sub

        ' ====== estructuras ======
        Private NotInheritable Class FileMeta
            Public Index As Integer
            Public HashFile As UInteger
            Public HashExt As UInteger
            Public HashDir As UInteger
            Public ModIndex As Byte
            Public ChunkCount As Byte
            Public ChunkHeaderSize As UShort
            Public Dx10_Width As UShort
            Public Dx10_Height As UShort
            Public Dx10_MipCount As Byte
            Public Dx10_DxgiFormatU8 As Byte
            Public Dx10_Flags As Byte
            Public Dx10_TileMode As Byte
            Public Chunk_CompSize As UInteger
            Public Chunk_DecompSize As UInteger
            Public Chunk_MipFirst As UShort
            Public Chunk_MipLast As UShort
            Public Chunk_CompData As Byte()
            Public FileName As String
            Public Directory As String
        End Class
    End Class

    Public NotInheritable Class Ba2WriterGNRL
        Public Shared Event Writed()

        Public NotInheritable Class Options
            Public Property Encoding As Encoding = Encoding.UTF8
            Public Property Version As UInteger = 8UI
            Public Property IncludeStrings As Boolean = True
            Public Property ZlibPreset As Ba2WriterCommon.ZlibPreset = Ba2WriterCommon.ZlibPreset.Default
            Public Property CompressionFormat As Ba2WriterCommon.CompressionFormat = Ba2WriterCommon.CompressionFormat.Zip

        End Class

        ''' <summary>Construye el BA2 GNRL en memoria.</summary>
        Public Shared Function Build(entries As IEnumerable(Of VirtualEntry), opts As Options) As Byte()
            Using ms As New MemoryStream()
                Write(ms, entries, opts)
                Return ms.ToArray()
            End Using
        End Function

        ''' <summary>
        ''' Escribe un BA2 GNRL v8 (FO4 NG) en <paramref name="output"/>.
        ''' Cada VirtualEntry usa .Directory/.FileName/.Data. (No requiere metadata DX10).
        ''' Compresión: zlib/Deflate. Si no comprime mejor, se guarda sin comprimir (CompSize=0).
        ''' </summary>
        Public Shared Sub Write(output As Stream, entries As IEnumerable(Of VirtualEntry), opts As Options)
            If output Is Nothing OrElse Not output.CanWrite Then Throw New ArgumentException("Stream inválido.", NameOf(output))
            If entries Is Nothing Then Throw New ArgumentNullException(NameOf(entries))
            If opts Is Nothing Then opts = New Options()

            Dim enc = If(opts.Encoding, Encoding.UTF8)
            Dim list = entries.ToList()
            If list.Count = 0 Then Throw New InvalidDataException("No hay entradas (GNRL).")

            ' Normalizar paths lógicos
            For Each ve In list
                ve.Directory = NormalizeSlash(ve.Directory)
            Next

            ' ===== Preparar metadatos y compresión por archivo (zlib) =====
            Dim metas As New List(Of FileMeta)(list.Count)
            For i As Integer = 0 To list.Count - 1
                Dim ve = list(i)
                Dim rel = JoinDirFile(ve.Directory, ve.FileName)
                Dim norm = PathUtil.NormalizeSlash(rel)

                Dim parent As String = "", stem As String = "", extNoDot As String = ""
                Ba2WriterCommon.Fo4SplitPath(norm, parent, stem, extNoDot)

                Dim raw As Byte() = If(ve.Data, Array.Empty(Of Byte)())
                If raw.Length > Integer.MaxValue Then Throw New InvalidDataException("Archivo demasiado grande (GNRL).")

                ' Compresión según Version/CompressionFormat:
                ' - v3 + LZ4 => LZ4 raw
                ' - resto => ZLIB
                Dim comp As Byte()
                If opts.Version = 3UI AndAlso opts.CompressionFormat = Ba2WriterCommon.CompressionFormat.Lz4 Then
                    comp = Ba2WriterCommon.CompressLz4(raw)
                Else
                    comp = Ba2WriterCommon.CompressZlib(raw, opts.ZlibPreset)
                End If

                Dim useComp As Boolean = comp IsNot Nothing AndAlso comp.Length > 0 AndAlso comp.Length < raw.Length

                Dim fm As New FileMeta With {
            .Index = i,
            .HashFile = Ba2WriterCommon.Crc32Ascii(stem),
            .HashExt = Ba2WriterCommon.PackExt(extNoDot),
            .HashDir = Ba2WriterCommon.Crc32Ascii(parent),
            .ModIndex = 0,
            .ChunkCount = 1,
            .ChunkHeaderSize = CUShort(16),              ' GNRL: 16 bytes; el sentinel se escribe aparte
            .CompSize = CUInt(If(useComp, comp.Length, 0)),
            .DecompSize = CUInt(raw.Length),
            .Directory = ve.Directory,
            .FileName = ve.FileName,
            .CompData = If(useComp, comp, raw)           ' lo que realmente se escribirá en payload
        }
                metas.Add(fm)
                RaiseEvent Writed()
            Next

            ' ===== Header BA2 (v1/v2/v3/v7/v8) =====
            If opts.Version <> 1UI AndAlso opts.Version <> 2UI AndAlso opts.Version <> 3UI AndAlso opts.Version <> 7UI AndAlso opts.Version <> 8UI Then
                Throw New InvalidDataException("GNRL admite Version=1,2,3,7,8.")
            End If
            If opts.CompressionFormat = Ba2WriterCommon.CompressionFormat.Lz4 AndAlso opts.Version <> 3UI Then
                Throw New InvalidDataException("LZ4 solo válido con BA2 Version=3.")
            End If



            Ba2WriterCommon.WriteAscii(output, "BTDX")
            Ba2WriterCommon.WriteU32(output, opts.Version)
            Ba2WriterCommon.WriteAscii(output, "GNRL")
            Ba2WriterCommon.WriteU32(output, CUInt(metas.Count))
            Dim posNameTable As Long = output.Position
            Ba2WriterCommon.WriteU64(output, 0UL)                 ' NameTableOffset (patch)
            ' v2/v3: campo extra u64(1)
            If opts.Version = 2UI OrElse opts.Version = 3UI Then
                Ba2WriterCommon.WriteU64(output, 1UL)
            End If
            ' v3: u32 compression_format (0=zip, 3=lz4)
            If opts.Version = 3UI Then
                Ba2WriterCommon.WriteCompressionFormatV3(output, opts.CompressionFormat)
            End If

            ' ===== File headers + chunk headers =====
            Dim patchOffsets As New List(Of Long)(metas.Count)
            For Each fm In metas
                ' per-file header (hashes + mini header GNRL)
                Ba2WriterCommon.WriteU32(output, fm.HashFile)
                Ba2WriterCommon.WriteU32(output, fm.HashExt)
                Ba2WriterCommon.WriteU32(output, fm.HashDir)
                output.WriteByte(fm.ModIndex)                     ' 0
                output.WriteByte(fm.ChunkCount)                   ' 1
                Ba2WriterCommon.WriteU16(output, fm.ChunkHeaderSize)  ' 16

                ' chunk header base (16 bytes) + sentinel (4) inmediatamente después
                patchOffsets.Add(output.Position)
                Ba2WriterCommon.WriteU64(output, 0UL)             ' Offset absoluto (patch)
                Ba2WriterCommon.WriteU32(output, fm.CompSize)     ' CompressedSize (0 si sin comprimir)
                Ba2WriterCommon.WriteU32(output, fm.DecompSize)   ' DecompressedSize
                Ba2WriterCommon.WriteU32(output, &HBAADF00DUI)    ' Sentinel (no cuenta dentro de los 16)
            Next

            ' ===== Payloads + parcheo de offsets absolutos =====
            For i = 0 To metas.Count - 1
                Dim fm = metas(i)
                Dim absOff As ULong = CULng(output.Position)

                ' Parchear offset absoluto del chunk
                Dim save = output.Position
                output.Position = patchOffsets(i)
                Ba2WriterCommon.WriteU64(output, absOff)
                output.Position = save

                ' Escribir los datos reales (comp o raw según política)
                If fm.CompData IsNot Nothing AndAlso fm.CompData.Length > 0 Then
                    output.Write(fm.CompData, 0, fm.CompData.Length)
                End If
            Next

            ' ===== Name table (UTF-8, u16 length) opcional =====
            If opts.IncludeStrings Then
                Dim nameTableOffset As ULong = CULng(output.Position)
                For Each fm In metas
                    Dim full = JoinDirFile(fm.Directory, fm.FileName)
                    Dim nb = enc.GetBytes(full)
                    If nb.Length > UShort.MaxValue Then Throw New InvalidDataException("Nombre demasiado largo (u16).")
                    Ba2WriterCommon.WriteU16(output, CUShort(nb.Length))
                    output.Write(nb, 0, nb.Length)
                Next
                Dim afterNames As Long = output.Position
                output.Position = posNameTable
                Ba2WriterCommon.WriteU64(output, nameTableOffset)
                output.Position = afterNames
            Else
                ' Omitir string table → mantener NameTableOffset=0 y no emitir nombres.
            End If

        End Sub

        ' ===== estructuras internas =====
        Private NotInheritable Class FileMeta
            Public Index As Integer
            Public HashFile As UInteger
            Public HashExt As UInteger
            Public HashDir As UInteger
            Public ModIndex As Byte
            Public ChunkCount As Byte
            Public ChunkHeaderSize As UShort
            Public CompSize As UInteger
            Public DecompSize As UInteger
            Public Directory As String
            Public FileName As String
            Public CompData As Byte()
        End Class


    End Class


End Namespace

