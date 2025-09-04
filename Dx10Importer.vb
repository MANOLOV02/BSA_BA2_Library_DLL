Option Strict On
Imports System.IO
Imports System.Text
Imports DirectXTexWrapperCLI   ' Loader.LoadTextures, TextureLoaded, TextureLevel

Namespace BethesdaArchive.Core

    ''' <summary>
    ''' Importa DDS y arma VirtualEntry para BA2 DX10 usando el wrapper DirectXTex.
    ''' - Preserva compresión (useCompress:=True)
    ''' - No fuerza OpenGL (forceOpenGL:=False)
    ''' - Soporta cubemap (Faces=6, IsCubemap=True)
    ''' - Concatena niveles en el orden del wrapper: [mip0_face0..N][mip1_face0..N]...
    ''' </summary>
    Public NotInheritable Class Dx10Importer
        Public Shared Function Concat(levels As IEnumerable(Of Byte())) As Byte()
            If levels Is Nothing Then Throw New ArgumentNullException(NameOf(levels))
            Dim total As Long = 0
            For Each a In levels
                If a Is Nothing Then Throw New InvalidDataException("Nivel nulo.")
                total += a.Length
            Next
            If total > Integer.MaxValue Then Throw New InvalidDataException("Textura demasiado grande (overflow Int32).")
            Dim blob As Byte() = New Byte(CInt(total) - 1) {}
            Dim cur As Integer = 0
            For Each a In levels
                Buffer.BlockCopy(a, 0, blob, cur, a.Length)
                cur += a.Length
            Next
            Return blob
        End Function

        ''' <summary>
        ''' Crea un VirtualEntry DX10 a partir de un buffer DDS y una ruta relativa (dentro del BA2).
        ''' </summary>
        Public Shared Function FromDdsBytes(ddsBytes As Byte(), relativePath As String) As VirtualEntry
            If ddsBytes Is Nothing OrElse ddsBytes.Length = 0 Then
                Throw New ArgumentException("DDS vacío.", NameOf(ddsBytes))
            End If
            If String.IsNullOrWhiteSpace(relativePath) Then
                Throw New ArgumentException("relativePath requerido.", NameOf(relativePath))
            End If

            ' 1) Cargar DDS con DirectXTex (preservando compresión)
            Dim lst = Loader.LoadTextures(New Byte()() {ddsBytes}, useCompress:=True, forceOpenGL:=False)
            If lst Is Nothing OrElse lst.Count <> 1 OrElse lst(0) Is Nothing OrElse Not lst(0).Loaded Then
                Throw New InvalidDataException("No se pudo cargar DDS con DirectXTex (Loaded=False).")
            End If
            Dim tex = lst(0)

            ' 2) Validaciones de metadata
            If tex.Miplevels <= 0 Then Throw New InvalidDataException("DDS sin mips válidos.")
            If tex.Faces <= 0 Then Throw New InvalidDataException("DDS sin cantidad de caras válida.")
            If tex.Levels Is Nothing OrElse tex.Levels.Count <> tex.Miplevels * tex.Faces Then
                Throw New InvalidDataException("Cantidad de niveles inconsistente (mips*caras).")
            End If
            Dim level0 = tex.Levels(0)
            If level0 Is Nothing OrElse level0.Data Is Nothing Then
                Throw New InvalidDataException("Nivel 0 inválido.")
            End If
            If tex.IsCubemap AndAlso tex.Faces <> 6 Then
                Throw New InvalidDataException("Cubemap inválido: Faces debe ser 6.")
            End If

            ' 3) Concatenar datos: [mip0_face0..Faces-1][mip1_face0..]...
            Dim blob As Byte() = Dx10Importer.Concat(tex.Levels.Select(Function(l) l.Data))

            ' 4) Separar directorio/archivo relativos (con '/')
            Dim dir As String = "", fileName As String = ""
            PathUtil.SplitDirFile(relativePath, dir, fileName)
            If String.IsNullOrWhiteSpace(fileName) Then
                Throw New InvalidDataException("Nombre de archivo relativo inválido.")
            End If

            ' 5) Armar VirtualEntry DX10
            Dim ve As New VirtualEntry() With {
              .Directory = dir,
              .FileName = fileName,
              .Data = blob,
              .PreferCompress = True,          ' sin efecto en DX10, lo dejamos en True
              .DxgiFormat = tex.DxgiCodeFinal, ' formato final reportado por DirectXTex
              .Width = level0.Width,
              .Height = level0.Height,
              .MipCount = tex.Miplevels,
              .IsCubemap = tex.IsCubemap,
              .Faces = tex.Faces
            }
            Return ve
        End Function

        ''' <summary>
        ''' Crea un VirtualEntry DX10 a partir de un archivo DDS en disco y un DataRoot (para la ruta relativa).
        ''' </summary>
        Public Shared Function FromDdsFile(ddsPath As String, dataRoot As String) As VirtualEntry
            If String.IsNullOrWhiteSpace(ddsPath) Then Throw New ArgumentException("ddsPath requerido.")
            If Not File.Exists(ddsPath) Then Throw New FileNotFoundException("DDS no encontrado.", ddsPath)

            Dim bytes = File.ReadAllBytes(ddsPath)
            Dim rel As String = PathUtil.MakeRelativeUnderDataRoot(dataRoot, ddsPath)
            Return FromDdsBytes(bytes, rel)
        End Function

    End Class

End Namespace

