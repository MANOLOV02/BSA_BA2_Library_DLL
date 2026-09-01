Imports System.IO

''' <summary>Escritura de archivos que funciona adentro de los dos gestores de mods.
''' <para>LEY: se sobrescribe el destino EN EL LUGAR. No hay temporal, no hay rename y no se borra el
''' destino. Bajo MO2 el VFS resuelve la apertura de un archivo existente al archivo real del mod que lo
''' aporta; bajo Vortex la escritura viaja por el hardlink. Un rename o un Delete+Create rompen las dos
''' cosas a la vez.</para>
''' <para>⛔ BASE Y ESTADO DE ESA AFIRMACION — es sobre TERCEROS, y esta cabecera es su UNICA casa (los
''' demas sitios apuntan aca, no la repiten): (1) el VFS de MO2 (usvfs) engancha abrir/borrar/renombrar
''' pero NO <c>ReplaceFileW</c>, asi que <c>File.Replace</c> corre contra el disco real, fuera del
''' perfil; (2) Vortex despliega por hardlink, y borrar o renombrar el destino corta el vinculo — el mod
''' se queda con la version vieja, y bajo MO2 lo escrito con nombre nuevo cae en <c>overwrite</c>.
''' Las dos afirmaciones vienen del cambio que introdujo esta clase (2.0.5); su derivacion NO quedo
''' registrada y NO estan re-medidas contra los gestores en este arbol. Si alguien las re-mide, la
''' medicion va ACA. Hasta entonces: NO volver al <c>.tmp</c>+<c>File.Replace</c> sin esa medicion —
''' el sintoma que este diseño evita es el archivo aterrizando en el Data real o en <c>overwrite</c>,
''' fuera de su mod.</para>
''' <para>NO ES ATOMICO, y tampoco tolera un lector que no comparta escritura: no existe primitivo que sea
''' a la vez virtualizado por el VFS, respetuoso del hardlink y atomico. La red la pone GuardarConCopia,
''' que NO escribe si no pudo dejar antes una copia verificada del contenido que va a pisar.</para>
''' <para>ESTA CLASE VA EN EL NAMESPACE RAIZ del proyecto (no en BethesdaArchive.Core): los consumidores
''' la importan como BSA_BA2_Library_DLL. Moverla de namespace rompe a los siete a la vez.</para></summary>
Public NotInheritable Class EscrituraEnElLugar

    Private Sub New()
    End Sub

    ''' <summary>Sufijo de la copia. NO es `.bak` (lo usa ArchivePackager) ni `.npcm.bak` (ese es el
    ''' respaldo de LoadOrderActivator, que vive hasta que su escritura queda CONFIRMADA por el verify
    ''' de relectura y recien ahi se borra — misma ley que esta, con el punto de confirmacion corrido).
    ''' <para>Los slots son `.npcm.prev`, `.npcm.prev2`, `.npcm.prev3`… La copia de cada corrida va al
    ''' primer slot LIBRE, asi que ninguna copia anterior se pisa nunca.</para></summary>
    Public Const SufijoCopia As String = ".npcm.prev"

    ''' <summary>Nucleo: la unica implementacion de la ley. <paramref name="seToco"/> sale en True desde
    ''' el instante en que el destino pudo abrirse — o sea, desde que puede quedar a medias. Los
    ''' llamadores lo usan para distinguir "no pude ni empezar" (destino intacto) de "fallo escribiendo".
    ''' <para><paramref name="sincronizar"/> agrega un <c>FlushFileBuffers</c> antes de cerrar: los bytes
    ''' llegan al plato y no se quedan en la cache del sistema. Ver <see cref="Escribir"/> para el costo
    ''' MEDIDO y por que el default es False.</para></summary>
    Private Shared Sub EscribirNucleo(destino As String, cuerpo As Action(Of Stream), ByRef seToco As Boolean,
                                      Optional sincronizar As Boolean = False)
        seToco = False
        Dim carpeta = IO.Path.GetDirectoryName(destino)
        If Not String.IsNullOrEmpty(carpeta) Then Directory.CreateDirectory(carpeta)

        ' OpenOrCreate SIN truncar: si el destino esta tomado, es de solo lectura o esta oculto, se falla
        ' ACA y el archivo queda intacto. (CREATE_ALWAYS sobre un archivo oculto da ACCESS_DENIED.)
        Using fs As New FileStream(destino, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
            seToco = True
            fs.SetLength(0)
            cuerpo(fs)
            ' Adentro del Using: cerrar NO sincroniza. Sin esto los bytes viven en la cache del sistema y
            ' un corte de luz se los lleva aunque el guardado haya dicho que salio bien.
            If sincronizar Then fs.Flush(True)
        End Using
    End Sub

    ''' <summary>Sobrescribe <paramref name="destino"/>. Sin red: es el camino de la salida regenerable
    ''' (horneado, build, texturas, materiales, cache, .pex).
    ''' <para>⛔ <paramref name="sincronizar"/> DEFAULTEA EN FALSE Y ESE DEFAULT ESTA MEDIDO. Con la forma
    ''' exacta de <see cref="EscribirNucleo"/> (OpenOrCreate → SetLength(0) → Write → Dispose), 300
    ''' archivos preexistentes reescritos, dos corridas alternadas por configuracion:</para>
    ''' <list type="bullet">
    ''' <item>disco del sistema, 300 KB: 0,654 ms/archivo sin sincronizar → 5,181 ms con. <b>+4,53 ms, x7,9.</b></item>
    ''' <item>disco del sistema, 2 KB: 0,580 → 3,242 ms. <b>+2,66 ms.</b> El costo es POR LLAMADA, no por byte.</item>
    ''' <item>segundo disco, 300 KB: 0,827 → 1,896 ms. +1,07 ms.</item>
    ''' </list>
    ''' <para>Sobre el horneado eso es +4,5 s cada 1.000 NIFs (+45 s a los 10.000) encima de un paso que
    ''' hoy tarda 0,654 ms. Por eso la salida REGENERABLE no sincroniza: si se pierde, se rehornea. El
    ''' dato del usuario si — <see cref="GuardarConCopia"/> sincroniza siempre y no pregunta.
    ''' ⛔ No dar vuelta este default sin volver a medir.</para></summary>
    Public Shared Sub Escribir(destino As String, cuerpo As Action(Of Stream),
                               Optional sincronizar As Boolean = False)
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If cuerpo Is Nothing Then Throw New ArgumentNullException(NameOf(cuerpo))

        Dim existia = File.Exists(destino)
        Dim seToco As Boolean = False
        Try
            EscribirNucleo(destino, cuerpo, seToco, sincronizar)
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
    ''' <para>⛔ LA LEY DE LA COPIA, y por que dice esto y no otra cosa. Las copias HEREDADAS (las que
    ''' dejo una caida anterior) NO se pisan y NO se borran a la ligera: cada una es un ejemplar del dato
    ''' del usuario que no existe en ningun otro lado. La de esta corrida va al primer slot LIBRE
    ''' (`.npcm.prev`, `.npcm.prev2`, `.npcm.prev3`…), asi que tomar red nunca destruye red.</para>
    ''' <para><b>Al salir bien se borra la copia que tomo ESTA corrida. Una heredada se borra SOLO cuando
    ''' esta corrida PROBO que no guarda nada propio: cuando la copia que tomamos —los bytes que habia en
    ''' el destino al empezar— es BYTE POR BYTE la misma.</b> No hay umbral, no hay tolerancia y no se
    ''' mira el CONTENIDO para adivinar si un archivo "parece" valido: es una comparacion exacta.</para>
    ''' <para>Antes decia «un guardado que sale bien limpia las dos» y eso BORRABA EL ULTIMO EJEMPLAR.
    ''' Medido: { destino = 0 B, .prev = 3000 B } → { destino = 5 B, .prev = AUSENTE }. El 0 B es
    ''' exactamente lo que deja un corte a mitad de <see cref="EscribirNucleo"/>, porque `SetLength(0)`
    ''' corre antes del cuerpo; y con el corte cayendo un poco despues el destino queda A MEDIAS con
    ''' longitud > 0 y el daño era identico. Gate: Tools\EscrituraEnElLugarGate (G3, G4, G5, G12).</para>
    ''' <para>⚠️ EL COSTO, dicho y no escondido: una heredada con contenido propio QUEDA EN DISCO al lado
    ''' del archivo del usuario (bajo Mod Organizer, en la carpeta Overwrite) hasta que el usuario la
    ''' borre o hasta que un guardado la pruebe redundante. Eso es deliberado: es su dato, y el unico
    ''' que lo tiene. La alternativa —borrarla igual— es el defecto que este parrafo documenta.</para></summary>
    Public Shared Sub GuardarConCopia(destino As String, cuerpo As Action(Of Stream))
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If cuerpo Is Nothing Then Throw New ArgumentNullException(NameOf(cuerpo))

        ' Las heredadas ocupan slots desde el 1; la de esta corrida va al primero LIBRE. Que el slot este
        ' libre es lo que hace seguro el `Borrar(copia)` del camino de error: nunca puede caer sobre una
        ' copia que no tomo esta corrida. (Antes el nombre era el fijo `.prev2` y el `Borrar` previo
        ' destruia la copia de una segunda caida — el mismo daño un nivel mas adentro.)
        Dim heredadas = SlotsOcupados(destino, SufijoCopia)
        Dim copia = Slot(destino, SufijoCopia, heredadas.Count + 1)

        Dim tieneContenido As Boolean = File.Exists(destino) AndAlso New FileInfo(destino).Length > 0
        Dim copiaHecha As Boolean = False
        If tieneContenido Then
            Try
                File.Copy(destino, copia, overwrite:=True)
                LimpiarAtributos(copia)
                ' ⛔ EXISTIR NO ES ESTAR COMPLETA. File.Copy puede dejar el destino a medias (disco lleno,
                ' error de I/O) y ahi lo que queda es un archivo que PARECE una copia. Se compara el
                ' tamano contra el original: si no coinciden, no hay copia.
                copiaHecha = (New FileInfo(copia).Length = New FileInfo(destino).Length)
                If Not copiaHecha Then
                    Borrar(copia)
                Else
                    ' La red solo es red si llego al plato: sin esto, un corte de luz puede llevarse la
                    ' copia Y el destino, que es justo el estado que toda esta clase existe para evitar.
                    Sincronizar(copia)
                End If
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
            ' sincronizar:=True y sin preguntar: ESTE es, por contrato, el camino del dato del usuario
            ' (no se usa en horneado ni en build). El costo medido esta en el docstring de `Escribir`:
            ' +2,66 a +4,53 ms POR ARCHIVO, y aca es un archivo por accion del usuario.
            EscribirNucleo(destino, cuerpo, seToco, sincronizar:=True)
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

        ' ⛔ SALIO BIEN. Aca vivia el defecto: `Borrar(heredada)` a secas.
        '
        ' Se borra la copia que tomo ESTA corrida — es un duplicado de lo que el usuario acaba de
        ' reemplazar a proposito, y eso es la mitad de la ley vieja que SI era cierta.
        '
        ' Una heredada se borra SOLO si esta corrida la probo redundante. La prueba es una implicacion
        ' cerrada, no un juicio: `copia` son los bytes que habia en el destino al empezar y, por la ley
        ' de arriba, es descartable; si una heredada es BYTE POR BYTE igual a `copia`, entonces no
        ' guarda nada que `copia` no guarde, y es descartable por el mismo motivo. Si difiere, es un
        ' ejemplar unico del dato del usuario y se queda.
        '
        ' Se compara ANTES de borrar `copia`, que es contra quien se compara.
        Dim redundantes As New List(Of String)
        If copiaHecha Then
            For Each h In heredadas
                If MismoContenido(h, copia) Then redundantes.Add(h)
            Next
        End If
        If copiaHecha Then Borrar(copia)
        For Each h In redundantes
            Borrar(h)
        Next
    End Sub

    ''' <summary>Vuelca un archivo ya escrito encima del destino, en el lugar. Lo usan los que no pueden
    ''' escribir directo: el packer (lee el archive viejo mientras produce el nuevo) y el exportador de
    ''' FOMOD (tiene que poder cancelar sin romper el zip anterior).
    ''' <para>Se abre el ORIGEN primero: si no esta, o esta tomado, el destino no se toca. El reintento
    ''' cubre SOLO el fallo de apertura —alguien tiene el archivo abierto y lo suelta en un rato—; si ya
    ''' se empezo a escribir no se reintenta, porque repetir tapa el error real.</para>
    ''' <para>⛔ CONTRATO de <paramref name="alTocarElDestino"/>: <b>si este callback NUNCA corrio y
    ''' VolcarEncima tiro, el destino esta byte-identico a como estaba.</b> Se dispara ADENTRO del cuerpo,
    ''' despues de abrir Y truncar y antes de copiar el primer byte del origen: es el instante exacto en
    ''' que el destino ya esta probadamente destruido. Existe porque los llamadores no podian distinguir
    ''' "no pude ni empezar" de "lo rompi", y sin esa distincion o avisan de una perdida que no paso, o
    ''' no intentan una recuperacion que si hacia falta.</para>
    ''' <para>El unico hueco, y se dice: si el <c>SetLength(0)</c> de <see cref="EscribirNucleo"/> tirara,
    ''' el callback no corre y el destino queda en estado desconocido. El <c>seToco</c> interno si cubre
    ''' ese instante; la senal publica no. No hay medicion de que ese fallo ocurra.</para></summary>
    Public Shared Sub VolcarEncima(origen As String, destino As String,
                                   Optional reintentos As Integer = 10,
                                   Optional esperaMs As Integer = 200,
                                   Optional alTocarElDestino As Action = Nothing)
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
                        ' El aviso va ADENTRO del cuerpo: EscribirNucleo ya abrio y ya trunco, y todavia
                        ' no se copio un byte. Ese es el instante en que el destino esta destruido y el
                        ' llamador lo tiene que saber. Afuera no sirve: antes de la llamada todavia no
                        ' paso nada, y despues ya no se distingue de "no pude ni empezar".
                        EscribirNucleo(destino,
                                       Sub(fs)
                                           alTocarElDestino?.Invoke()
                                           src.CopyTo(fs)
                                       End Sub,
                                       seToco)
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

    ''' <summary>Nombre del slot <paramref name="n"/> de respaldo: 1 =&gt; `&lt;destino&gt;&lt;sufijo&gt;`,
    ''' 2 =&gt; `…&lt;sufijo&gt;2`, 3 =&gt; `…&lt;sufijo&gt;3`… El 1 no lleva numero para no romper los
    ''' archivos que ya estan en disco de las versiones anteriores.</summary>
    Private Shared Function Slot(destino As String, sufijo As String, n As Integer) As String
        Return destino & sufijo &
               If(n = 1, "", n.ToString(Globalization.CultureInfo.InvariantCulture))
    End Function

    ''' <summary>Los slots de respaldo OCUPADOS, en orden desde el 1. Cada uno es una copia HEREDADA: la
    ''' dejo una corrida que nunca confirmo su escritura (el proceso murio, o la restauracion automatica
    ''' fallo y el mensaje le dijo al usuario donde estaba su version anterior).
    ''' <para>Corta en el primer hueco a proposito. Si el usuario borro `.npcm.prev` a mano y dejo
    ''' `.npcm.prev2`, esta funcion no ve la segunda: no la cuenta como heredada, no la compara y no la
    ''' borra NUNCA. Un huerfano que sobrevive de mas es el lado seguro del error.</para></summary>
    Private Shared Function SlotsOcupados(destino As String, sufijo As String) As List(Of String)
        Dim ocupados As New List(Of String)
        Dim n As Integer = 1
        While File.Exists(Slot(destino, sufijo, n))
            ocupados.Add(Slot(destino, sufijo, n))
            n += 1
        End While
        Return ocupados
    End Function

    ''' <summary>El primer nombre de respaldo LIBRE para <paramref name="destino"/> con el sufijo dado.
    ''' <para>⛔ ESTA ES LA LEY DEL NOMBRE, Y VIVE ACA SOLA. La usa <see cref="GuardarConCopia"/> con
    ''' <see cref="SufijoCopia"/> y la usa <c>LoadOrderActivator.WriteEntries</c> con su `.npcm.bak`, que
    ''' no puede usar GuardarConCopia porque su respaldo tiene que sobrevivir a la escritura (lo consume
    ''' el rollback del verify de relectura, que corre despues). Dos artefactos con vidas distintas, una
    ''' sola ley: <b>un respaldo que ya esta en disco no se pisa jamas</b>.</para></summary>
    Public Shared Function PrimerSlotLibre(destino As String, sufijo As String) As String
        If String.IsNullOrEmpty(destino) Then Throw New ArgumentException("Empty path.", NameOf(destino))
        If String.IsNullOrEmpty(sufijo) Then Throw New ArgumentException("Empty suffix.", NameOf(sufijo))
        Return Slot(destino, sufijo, SlotsOcupados(destino, sufijo).Count + 1)
    End Function

    ''' <summary>Comparacion BYTE POR BYTE de dos archivos. Corta por tamano primero.
    ''' <para>⛔ NO alcanza con comparar el tamano: es la misma trampa que documenta el `copiaHecha` de
    ''' <see cref="GuardarConCopia"/> — dos archivos distintos con la misma longitud existen, y aca el
    ''' precio de equivocarse es borrar el unico ejemplar del dato del usuario.</para>
    ''' <para>Costo: cero en el camino sano, porque sin copias heredadas no se llama. Cuando se llama,
    ''' son dos lecturas secuenciales de un archivo que esta misma corrida YA copio entero unas lineas
    ''' antes — es estrictamente menos que el `File.Copy` que el metodo ya paga.</para></summary>
    Private Shared Function MismoContenido(a As String, b As String) As Boolean
        Try
            Dim fa As New FileInfo(a), fb As New FileInfo(b)
            If Not fa.Exists OrElse Not fb.Exists Then Return False
            If fa.Length <> fb.Length Then Return False
            Using sa As Stream = File.OpenRead(a), sb As Stream = File.OpenRead(b)
                Dim ba(65535) As Byte, bb(65535) As Byte
                Do
                    Dim na = LeerBloque(sa, ba)
                    Dim nb = LeerBloque(sb, bb)
                    If na <> nb Then Return False
                    If na = 0 Then Return True
                    For i = 0 To na - 1
                        If ba(i) <> bb(i) Then Return False
                    Next
                Loop
            End Using
        Catch
            ' ⛔ Si no se pudo comparar, la respuesta segura es NO. Una heredada que no pudimos PROBAR
            ' redundante no se borra: el fallo del comparador no puede costarle el dato al usuario.
            Return False
        End Try
        Return False
    End Function

    ''' <summary>Lee hasta llenar el buffer o agotar el stream. `Stream.Read` puede devolver menos de lo
    ''' pedido sin estar en el final, y comparar bloques de largo distinto daria "difieren" sobre dos
    ''' archivos identicos.</summary>
    Private Shared Function LeerBloque(s As Stream, buf As Byte()) As Integer
        Dim total As Integer = 0
        While total < buf.Length
            Dim n = s.Read(buf, total, buf.Length - total)
            If n = 0 Then Exit While
            total += n
        End While
        Return total
    End Function

    ''' <summary>FlushFileBuffers sobre un archivo ya cerrado, para el dato del usuario. Best-effort: no
    ''' poder sincronizar no invalida una copia que ya esta escrita y verificada por tamano.</summary>
    Private Shared Sub Sincronizar(ruta As String)
        Try
            Using fs As New FileStream(ruta, FileMode.Open, FileAccess.Write, FileShare.Read)
                fs.Flush(True)
            End Using
        Catch
        End Try
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
