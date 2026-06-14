Option Strict On
Imports System.IO
Imports DirectXTexWrapperCLI   ' Loader.GetDdsMetadata, DdsMetadata

Namespace BethesdaArchive.Core

    ''' <summary>
    ''' Imports DDS files into VirtualEntry instances suitable for BA2 DX10 archives.
    '''
    ''' Contract for BA2 DX10 (FO4): VirtualEntry.Data carries the STRIPPED payload (mip data
    ''' concatenated, no DDS header), and the explicit DX10 metadata (Width / Height / MipCount /
    ''' DxgiFormat / IsCubemap / Faces) is populated on the entry. The writer NEVER parses the
    ''' DDS header; the reader reconstructs it on extraction via Loader.EncodeDDSHeader.
    '''
    ''' BSA (SSE) does NOT use this importer — it stores DDS files verbatim (header included)
    ''' as opaque bytes; callers should pass the full DDS as VirtualEntry.Data without going
    ''' through SplitDdsBytes / FromDdsBytes.
    ''' </summary>
    Public NotInheritable Class Dx10Importer

        ''' <summary>Returns True if the buffer starts with the DDS magic ("DDS ").</summary>
        Public Shared Function HasDdsMagic(bytes As Byte()) As Boolean
            Return bytes IsNot Nothing AndAlso
                   bytes.Length >= 4 AndAlso
                   bytes(0) = AscW("D"c) AndAlso
                   bytes(1) = AscW("D"c) AndAlso
                   bytes(2) = AscW("S"c) AndAlso
                   bytes(3) = AscW(" "c)
        End Function

        ''' <summary>
        ''' Header-only DDS parse via DirectXTex's GetMetadataFromDDSMemory (μs-cost, no pixel
        ''' decode, no allocs beyond the returned managed objects). Returns the DX10 metadata
        ''' and the payload (DDS file bytes minus the header) ready for BA2 DX10 packaging.
        ''' </summary>
        Public Shared Function SplitDdsBytes(ddsBytes As Byte()) As (Metadata As DdsMetadata, Payload As Byte())
            If ddsBytes Is Nothing OrElse ddsBytes.Length = 0 Then
                Throw New ArgumentException("DDS vacío.", NameOf(ddsBytes))
            End If
            If Not HasDdsMagic(ddsBytes) Then
                Throw New InvalidDataException("DDS inválido: falta magic 'DDS '.")
            End If

            Dim md = Loader.GetDdsMetadata(ddsBytes)
            If md Is Nothing OrElse Not md.Loaded Then
                Throw New InvalidDataException("DDS inválido o formato no soportado.")
            End If
            If md.HeaderSize <= 0 OrElse md.HeaderSize > ddsBytes.Length Then
                Throw New InvalidDataException("DDS truncado o header tamaño inválido.")
            End If

            Dim payloadLen As Integer = ddsBytes.Length - md.HeaderSize
            Dim payload As Byte()
            If payloadLen > 0 Then
                payload = New Byte(payloadLen - 1) {}
                Buffer.BlockCopy(ddsBytes, md.HeaderSize, payload, 0, payloadLen)
            Else
                payload = Array.Empty(Of Byte)()
            End If
            Return (md, payload)
        End Function

        ''' <summary>
        ''' Builds a VirtualEntry for BA2 DX10 from a DDS buffer + relative path. The returned
        ''' entry has Data = stripped payload and the DX10 metadata fields populated.
        ''' </summary>
        Public Shared Function FromDdsBytes(ddsBytes As Byte(), relativePath As String) As VirtualEntry
            If String.IsNullOrWhiteSpace(relativePath) Then
                Throw New ArgumentException("relativePath requerido.", NameOf(relativePath))
            End If

            Dim split = SplitDdsBytes(ddsBytes)
            Dim md = split.Metadata
            Dim payload = split.Payload

            Dim dir As String = "", fileName As String = ""
            PathUtil.SplitDirFile(relativePath, dir, fileName)
            If String.IsNullOrWhiteSpace(fileName) Then
                Throw New InvalidDataException("Nombre de archivo relativo inválido.")
            End If

            Return New VirtualEntry() With {
                .Directory = dir,
                .FileName = fileName,
                .Data = payload,                ' contract: stripped payload, no DDS header
                .PreferCompress = True,         ' sin efecto en DX10, mantenido por compatibilidad
                .DxgiFormat = md.DxgiFormat,
                .Width = md.Width,
                .Height = md.Height,
                .MipCount = md.MipCount,
                .IsCubemap = md.IsCubemap,
                .Faces = md.Faces
            }
        End Function

        ''' <summary>Convenience wrapper that reads a DDS file from disk and forwards to FromDdsBytes.</summary>
        Public Shared Function FromDdsFile(ddsPath As String, dataRoot As String) As VirtualEntry
            If String.IsNullOrWhiteSpace(ddsPath) Then Throw New ArgumentException("ddsPath requerido.")
            If Not File.Exists(ddsPath) Then Throw New FileNotFoundException("DDS no encontrado.", ddsPath)

            Dim bytes = File.ReadAllBytes(ddsPath)
            ' Signature is MakeRelativeUnderDataRoot(absPath, dataRoot): the absolute DDS path
            ' first, the data root second. (Previously these were swapped.)
            Dim rel As String = PathUtil.MakeRelativeUnderDataRoot(ddsPath, dataRoot)
            Return FromDdsBytes(bytes, rel)
        End Function

    End Class

End Namespace
