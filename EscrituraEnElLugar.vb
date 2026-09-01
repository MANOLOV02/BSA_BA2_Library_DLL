Imports System.IO

''' <summary>Escritura de archivos que funciona adentro de los dos gestores de mods.
''' <para>LEY: se sobrescribe el destino EN EL LUGAR. No hay temporal, no hay rename y no se borra el
''' destino. Bajo MO2 el VFS resuelve la apertura de un archivo existente al archivo real del mod que lo
''' aporta; bajo Vortex la escritura viaja por el hardlink. Un rename o un Delete+Create rompen las dos
''' cosas a la vez. Derivacion y mediciones: nota de memoria 10-stack-escritura-bajo-mo2-y-vortex.</para>
''' <para>NO ES ATOMICO, y tampoco tolera un lector que no comparta escritura: no existe primitivo que sea
''' a la vez virtualizado por el VFS, respetuoso del hardlink y atomico. La red la pone GuardarConCopia,
''' que NO escribe si no pudo dejar antes una copia verificada del contenido que va a pisar.</para>
''' <para>ESTA CLASE VA EN EL NAMESPACE RAIZ del proyecto (no en BethesdaArchive.Core): los consumidores
''' la importan como BSA_BA2_Library_DLL. Moverla de namespace rompe a los siete a la vez.</para></summary>
Public NotInheritable Class EscrituraEnElLugar

    Private Sub New()
    End Sub

    ''' <summary>Sufijo de la copia. NO es `.bak` (lo usa ArchivePackager) ni `.npcm.bak`
    ''' (LoadOrderActivator lo deja en disco desde 1.4.1 y nunca lo borra).</summary>
    Public Const SufijoCopia As String = ".npcm.prev"

    ''' <summary>Nucleo: la unica implementacion de la ley. <paramref name="seToco"/> sale en True desde
    ''' el instante en que el destino pudo abrirse — o sea, desde que puede quedar a medias. Los
    ''' llamadores lo usan para distinguir "no pude ni empezar" (destino intacto) de "fallo escribiendo".</summary>
    Private Shared Sub EscribirNucleo(destino As String, cuerpo As Action(Of Stream), ByRef seToco As Boolean)
        seToco = False
        Dim carpeta = IO.Path.GetDirectoryName(destino)
        If Not String.IsNullOrEmpty(carpeta) Then Directory.CreateDirectory(carpeta)

        ' OpenOrCreate SIN truncar: si el destino esta tomado, es de solo lectura o esta oculto, se falla
        ' ACA y el archivo queda intacto. (CREATE_ALWAYS sobre un archivo oculto da ACCESS_DENIED.)
        Using fs As New FileStream(destino, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
            seToco = True
            fs.SetLength(0)
            cuerpo(fs)
        End Using
    End Sub

    ''' <summary>Sobrescribe <paramref name="destino"/>. Sin red: es el camino de la salida regenerable
    ''' (horneado, build, texturas, materiales, cache, .pex).</summary>
    Public Shared Sub Escribir(destino As String, cuerpo As Action(Of Stream))
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If cuerpo Is Nothing Then Throw New ArgumentNullException(NameOf(cuerpo))

        Dim existia = File.Exists(destino)
        Dim seToco As Boolean = False
        Try
            EscribirNucleo(destino, cuerpo, seToco)
        Catch
            ' No dejar un archivo de 0 bytes donde antes no habia NADA: el juego y xEdit levantan un
            ' .esp vacio como corrupto.
            If seToco AndAlso Not existia Then Borrar(destino)
            Throw
        End Try
    End Sub

    ''' <summary>Igual, con red: copia el contenido actual a &lt;destino&gt;.npcm.prev antes de tocarlo y, si
    ''' la escritura falla DESPUES de haber empezado, restaura el destino solo. Es el camino de los datos
    ''' del usuario (plugin, ini de BodyGen, sidecars, proyecto de WM). NO se usa en horneado ni en build:
    ''' ahi son miles de archivos por corrida y se regeneran solos.
    ''' <para>⛔ LA COPIA NO ES OPCIONAL. Si hay contenido que perder y no se pudo dejar una copia
    ''' VERIFICADA (el tamano tiene que coincidir: File.Copy puede dejarla a medias), NO se escribe. Antes
    ''' esto seguia igual "porque cancelar el guardado seria peor", y el resultado era truncar el unico
    ''' ejemplar que el usuario tenia sin nada para volver.</para>
    ''' <para>Una copia HEREDADA de una caida anterior no se pisa —no sabemos de que version es— pero
    ''' tampoco deja sin red: la de esta corrida se toma con otro nombre. Un guardado que sale bien limpia
    ''' las dos, que es lo que saca al archivo del estado degradado.</para></summary>
    Public Shared Sub GuardarConCopia(destino As String, cuerpo As Action(Of Stream))
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If cuerpo Is Nothing Then Throw New ArgumentNullException(NameOf(cuerpo))

        ' La copia HEREDADA (de una caida anterior) no se toca: no sabemos de que version es y es lo
        ' unico que le queda al usuario de ese momento. Pero tampoco puede dejarnos sin red: si esta,
        ' la de esta corrida se toma con OTRO nombre. Se limpian las dos cuando el guardado sale bien.
        Dim heredada = destino & SufijoCopia
        Dim hayHeredada As Boolean = File.Exists(heredada)
        Dim copia = If(hayHeredada, destino & SufijoCopia & "2", heredada)

        Dim tieneContenido As Boolean = File.Exists(destino) AndAlso New FileInfo(destino).Length > 0
        Dim copiaHecha As Boolean = False
        If tieneContenido Then
            Try
                Borrar(copia)                       ' un resto de otra corrida no se reusa NUNCA
                File.Copy(destino, copia, overwrite:=True)
                LimpiarAtributos(copia)
                ' ⛔ EXISTIR NO ES ESTAR COMPLETA. File.Copy puede dejar el destino a medias (disco lleno,
                ' error de I/O) y ahi lo que queda es un archivo que PARECE una copia. Se compara el
                ' tamano contra el original: si no coinciden, no hay copia.
                copiaHecha = (New FileInfo(copia).Length = New FileInfo(destino).Length)
                If Not copiaHecha Then Borrar(copia)
            Catch
                copiaHecha = False
                Borrar(copia)
            End Try

            ' ⛔ SIN RED NO SE TRUNCA. Si hay contenido que perder y no se pudo respaldar, no se toca el
            ' original: escribir igual seria destruir la unica copia que existe sin nada para volver. El
            ' usuario ve por que fallo (disco lleno, permisos, el archivo tomado) y su dato sigue entero.
            If Not copiaHecha Then
                Throw New IOException(
                    $"'{IO.Path.GetFileName(destino)}' was NOT modified: its backup could not be created" &
                    $" ('{IO.Path.GetFileName(copia)}'). Free up disk space or close whatever is holding " &
                    "that file, then save again.")
            End If
        End If

        Dim existia = File.Exists(destino)
        Dim seToco As Boolean = False
        Try
            EscribirNucleo(destino, cuerpo, seToco)
        Catch ex As Exception
            If Not seToco Then
                ' No se pudo ni abrir: el destino esta INTACTO. No hay nada que restaurar y no hay que
                ' decirle al usuario que quedo a medias.
                If copiaHecha Then Borrar(copia)
                Throw
            End If
            If seToco AndAlso Not existia Then Borrar(destino)
            If copiaHecha Then
                ' Solo se restaura desde la copia que tomo ESTA corrida: es la unica de la que sabemos
                ' que contenido tiene. La heredada se deja donde esta.
                Try
                    File.Copy(copia, destino, overwrite:=True)
                Catch
                    Throw New IOException(MensajeCopiaViva(destino, copia), ex)
                End Try
                Borrar(copia)           ' fuera del Try de arriba: que falle el borrado NO es "no pude restaurar"
            End If
            Throw
        End Try

        ' Salio bien: el archivo en disco es el bueno, asi que se limpian las dos copias — tambien la
        ' heredada de una caida anterior, que es lo que saca al archivo del estado degradado.
        Borrar(heredada)
        Borrar(destino & SufijoCopia & "2")
    End Sub

    ''' <summary>Vuelca un archivo ya escrito encima del destino, en el lugar. Lo usan los que no pueden
    ''' escribir directo: el packer (lee el archive viejo mientras produce el nuevo) y el exportador de
    ''' FOMOD (tiene que poder cancelar sin romper el zip anterior).
    ''' <para>Se abre el ORIGEN primero: si no esta, o esta tomado, el destino no se toca. El reintento
    ''' cubre SOLO el fallo de apertura —alguien tiene el archivo abierto y lo suelta en un rato—; si ya
    ''' se empezo a escribir no se reintenta, porque repetir tapa el error real.</para></summary>
    Public Shared Sub VolcarEncima(origen As String, destino As String,
                                   Optional reintentos As Integer = 10,
                                   Optional esperaMs As Integer = 200)
        If String.IsNullOrEmpty(origen) Then Throw New ArgumentException("Empty path.", NameOf(origen))
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If reintentos < 1 Then Throw New ArgumentOutOfRangeException(NameOf(reintentos))

        For intento = 1 To reintentos
            Dim seToco As Boolean = False
            Try
                Using src As New FileStream(origen, FileMode.Open, FileAccess.Read,
                                            FileShare.Read Or FileShare.Delete)
                    If src.Length = 0 Then
                        Throw New InvalidDataException(
                            $"'{IO.Path.GetFileName(origen)}' is empty: refusing to overwrite " &
                            $"'{IO.Path.GetFileName(destino)}' with nothing.")
                    End If
                    Dim existiaDestino = File.Exists(destino)
                    Try
                        EscribirNucleo(destino, Sub(fs) src.CopyTo(fs), seToco)
                    Catch
                        ' Misma ley que Escribir: donde no habia nada, no queda un archivo de 0 bytes.
                        If seToco AndAlso Not existiaDestino Then Borrar(destino)
                        Throw
                    End Try
                End Using
                Return
            Catch ex As Exception When (TypeOf ex Is IOException OrElse
                                        TypeOf ex Is UnauthorizedAccessException) AndAlso
                                       Not seToco AndAlso intento < reintentos
                Threading.Thread.Sleep(esperaMs)
            End Try
        Next
    End Sub

    Private Shared Sub Borrar(ruta As String)
        Try
            File.Delete(ruta)
        Catch
            ' Best-effort en todos los usos: un archivo huerfano no rompe nada.
        End Try
    End Sub

    Private Shared Sub LimpiarAtributos(ruta As String)
        Try
            Dim attr = File.GetAttributes(ruta)
            Dim malos = FileAttributes.ReadOnly Or FileAttributes.Hidden
            If (attr And malos) <> 0 Then File.SetAttributes(ruta, attr And Not malos)
        Catch
        End Try
    End Sub

    Private Shared Function MensajeCopiaViva(destino As String, copia As String) As String
        Return $"'{IO.Path.GetFileName(destino)}' was left half-written and could not be restored " &
               $"automatically. The previous version is in '{IO.Path.GetFileName(copia)}', next to it " &
               "(with Mod Organizer, in the Overwrite folder)."
    End Function
End Class
