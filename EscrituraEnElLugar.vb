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

    ''' <summary>El cuerpo CERRO el stream, rompiendo el contrato de <see cref="EscribirNucleo"/>.
    ''' <para>⛔ ES UN TIPO PROPIO PORQUE HAY UNA DECISION COLGADA DE EL, no para clasificar mejor. Cuando
    ''' esto se tira, la escritura del cuerpo YA TERMINO: el wrapper cerro y vacio, y los bytes que hay en
    ''' el destino son exactamente los que el cuerpo produjo. Lo unico que la violacion se llevo puesto es
    ''' la GARANTIA DE DURABILIDAD (el <c>Flush(True)</c> de esta clase no llego a correr).</para>
    ''' <para><b>Por eso un destino NUEVO no se borra en este camino, y ese es el cambio.</b> El
    ''' <c>Borrar(destino)</c> de <see cref="Escribir"/> existe para no dejar "un archivo de 0 bytes donde
    ''' antes no habia NADA" — un archivo del que no sabemos nada porque el cuerpo murio a mitad. Acá el
    ''' cuerpo NO murio: volvio normalmente. Y con un cuerpo bien portado esos mismos bytes se aceptan sin
    ''' verificar nada —<see cref="Escribir"/> no mira un solo byte antes de devolver "guardado"—, asi que
    ''' aplicarles un estandar mas duro por el solo hecho de que el cuerpo uso un wrapper es una asimetria
    ''' que no se sostiene: no sabemos MENOS del archivo en un caso que en el otro.</para>
    ''' <para>⛔ Y SE SIGUE TIRANDO IGUAL. Que el archivo se conserve no vuelve buena a la corrida: el
    ''' llamador tiene que arreglar el <c>leaveOpen:=True</c> antes de confiar en la durabilidad de lo que
    ''' acaba de escribir. Conservar el archivo y avisar son cosas distintas.</para>
    ''' <para>⚠️ NO cambia el otro brazo: si el CUERPO TIRO, los bytes del destino si son desconocidos
    ''' (pudo morir a mitad) y el borrado del destino nuevo se queda como estaba. Tampoco cambia la
    ''' restauracion desde la copia de <see cref="GuardarConCopia"/>: alli el destino EXISTIA y devolverle
    ''' al usuario su version anterior es lo correcto pase lo que pase.</para>
    ''' <para>Hereda de <see cref="InvalidOperationException"/> a proposito: es lo que los llamadores y el
    ''' gate ya atrapan, asi que nadie que hoy lo maneje deja de manejarlo.</para></summary>
    Public NotInheritable Class ContratoDelCuerpoException
        Inherits InvalidOperationException

        Public Sub New(mensaje As String)
            MyBase.New(mensaje)
        End Sub
    End Class

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
    ''' MEDIDO y por que el default es False.</para>
    ''' <para>⛔⛔ CONTRATO DEL CUERPO: <b>el cuerpo escribe en el stream y NO LO CIERRA.</b> El dueño del
    ''' stream es esta clase, que lo abre, lo trunca, lo sincroniza y lo cierra. Un cuerpo que envuelve el
    ''' stream en un `BinaryWriter` / `StreamWriter` / `XmlWriter` / `GZipStream` dentro de un `Using`
    ''' <b>tiene que pasar `leaveOpen:=True`</b>, porque si no el `End Using` del wrapper cierra el
    ''' FileStream de abajo.</para>
    ''' <para>Esto no es teorico y el costo fue una app cerrada: `OSD_Class.Save_As` envolvia en
    ''' `New IO.BinaryWriter(stream)` sin `leaveOpen`. Mientras no habia nada despues de `cuerpo(fs)` el
    ''' defecto era invisible —el stream se cerraba dos veces y a nadie le importaba—; en cuanto el
    ''' `Flush(True)` se puso DESPUES del cuerpo, el mismo codigo empezo a tirar
    ''' `ObjectDisposedException: Cannot access a closed file` en produccion. Por eso la guarda de abajo
    ''' existe: el proximo violador tiene que fallar diciendo QUE contrato rompio y sobre QUE archivo, no
    ''' con un ObjectDisposed criptico desde las entrañas del framework.</para></summary>
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

            ' ⛔ GUARDA DEL CONTRATO. Si el cuerpo cerro el stream, `CanWrite` da False y CUALQUIER cosa
            ' que hagamos despues revienta con un mensaje que no nombra ni la causa ni el archivo. Se
            ' falla ACA, diciendo las dos cosas.
            ' ⛔ Y NO se saltea el Flush en silencio: un cuerpo que cierra el stream ya se llevo puesta la
            ' garantia de durabilidad (el wrapper cerro y sincronizo lo que quiso, cuando quiso), asi que
            ' seguir de largo seria devolver "guardado" sobre una promesa que no podemos sostener.
            ' ⛔ TIPO PROPIO, no un InvalidOperationException pelado: los llamadores tienen que poder
            ' distinguir ESTE fallo de cualquier otro, porque de el cuelga si el destino nuevo se borra o
            ' se conserva. Ver ContratoDelCuerpoException. Deriva de InvalidOperationException, asi que
            ' quien ya lo atrapaba lo sigue atrapando.
            If Not fs.CanWrite Then
                Throw New ContratoDelCuerpoException(
                    $"The write body closed the stream for '{IO.Path.GetFileName(destino)}'. Bodies passed to " &
                    "EscrituraEnElLugar must write and NOT close: this class owns the stream (it opens, " &
                    "truncates, syncs and closes it). If the body wraps the stream in a BinaryWriter, " &
                    "StreamWriter, XmlWriter or similar inside a Using, pass leaveOpen:=True. " &
                    "The file was left with the bytes the body wrote, but WITHOUT the durability flush " &
                    "this class guarantees.")
            End If

            ' Adentro del Using: cerrar NO sincroniza. Sin esto los bytes viven en la cache del sistema y
            ' un corte de luz se los lleva aunque el guardado haya dicho que salio bien.
            If sincronizar Then fs.Flush(True)
        End Using
    End Sub

    ''' <summary>Sobrescribe <paramref name="destino"/>. Sin red: es el camino de la salida regenerable
    ''' (horneado, build, texturas, materiales, cache, .pex).
    ''' <para>⛔ QUE SIGNIFICA "materiales" EN ESA LISTA, porque la ambiguedad ya costo una revision: es el
    ''' camino de CLONE (<c>Clone_Materials_class.WriteMaterialJob</c>, que escribe a <c>ManoloCloned\</c> —
    ''' una carpeta que la app posee y regenera sola en la corrida siguiente). NO es el guardado de un
    ''' material desde el editor de Wardrobe Manager: ese pisa un suelto YA INSTALADO que la app no
    ''' produjo y no puede rehacer, es un archivo por click, y va por <see cref="GuardarConCopia"/> —
    ''' igual que <c>NifContent_Class.Save_As_Manolo_ConCopia</c> frente a <c>Save_As_Manolo</c>.</para>
    ''' <para>⛔ LA LISTA NO SE LEE POR TIPO DE ARCHIVO, SE LEE POR LLAMADOR. El NIF esta en las DOS
    ''' listas y el material tambien: lo que decide es quien escribe y si lo escrito se puede rehacer.</para>
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
        Catch ex As ContratoDelCuerpoException
            ' ⛔ ACA NO SE BORRA, y el motivo entero esta en ContratoDelCuerpoException: el cuerpo VOLVIO
            ' NORMALMENTE y despues fallo la guarda, asi que los bytes del destino son exactamente los que
            ' el cuerpo produjo — los mismos que, viniendo de un cuerpo bien portado, esta funcion acepta
            ' sin verificar nada. Lo que se perdio es el Flush, no el archivo. Se tira igual.
            Throw
        Catch
            ' No dejar un archivo de 0 bytes donde antes no habia NADA: el juego y xEdit levantan un
            ' .esp vacio como corrupto. Este brazo es el del cuerpo que TIRO: ahi los bytes SI son
            ' desconocidos (pudo morir a mitad) y el borrado se queda como estaba.
            If seToco AndAlso Not existia Then Borrar(destino)
            Throw
        End Try
    End Sub

    ''' <summary>Igual, con red: copia el contenido actual a &lt;destino&gt;.npcm.prev antes de tocarlo y, si
    ''' la escritura falla DESPUES de haber empezado, restaura el destino solo. Es el camino de los datos
    ''' del usuario (plugin, ini de BodyGen, sidecars, proyecto de WM). NO se usa en horneado ni en build:
    ''' ahi son miles de archivos por corrida y se regeneran solos.
    ''' <para>⛔ "RESTAURA EL DESTINO SOLO" ES UNA PROMESA, Y ESTUVO ROTA. La restauracion era un
    ''' <c>File.Copy</c> —CREATE_ALWAYS— mientras la escritura ya usaba <c>OpenOrCreate</c>: sobre un
    ''' destino OCULTO, que es justo el caso para el que existe todo este diseño, el camino de ida
    ''' funcionaba y el de vuelta tiraba, dejando el destino PARCIAL. Hoy la vuelta usa la MISMA primitiva
    ''' que la ida. Testigos: caso <c>M</c> (la medicion de que primitiva aguanta que atributo) y
    ''' <c>G21</c> (destino oculto restaurado byte a byte y todavia oculto) de
    ''' <c>Tools\EscrituraEnElLugarGate</c>; <c>G8</c> es el mismo camino con destino normal, que es por lo
    ''' que el defecto podia vivir sin que nada se pusiera rojo.</para>
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
        Dim causaCopia As String = ""
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
            Catch ex As Exception
                ' ⛔ LA CAUSA REAL NO SE TIRA A LA BASURA. Acá habia un `Catch` pelado, asi que una
                ' `PathTooLongException`, una `DirectoryNotFoundException` o una `NotSupportedException`
                ' salian todas disfrazadas del mensaje fijo de abajo —"disco lleno o el archivo tomado"— y
                ' el usuario se quedaba buscando espacio en disco por un problema de ruta.
                copiaHecha = False
                causaCopia = ex.Message
                Borrar(copia)
            End Try

            ' ⛔ SIN RED NO SE TRUNCA. Si hay contenido que perder y no se pudo respaldar, no se toca el
            ' original: escribir igual seria destruir la unica copia que existe sin nada para volver. El
            ' usuario ve por que fallo (disco lleno, permisos, el archivo tomado) y su dato sigue entero.
            If Not copiaHecha Then
                ' El `causaCopia` va EMBEBIDO en el texto y no como InnerException: los dialogos de la app
                ' muestran solo `.Message` (FomodExport_Form.vb:355 lo documenta para el caso gemelo), asi
                ' que un inner es una causa que nadie ve. Cuando la copia fallo por tamaño —no hubo
                ' excepcion, solo no coincidio— no hay causa que agregar y el mensaje queda como estaba.
                Throw New IOException(
                    $"'{IO.Path.GetFileName(destino)}' was NOT modified: its backup could not be created" &
                    $" ('{IO.Path.GetFileName(copia)}'). Free up disk space or close whatever is holding " &
                    "that file, then save again." &
                    If(causaCopia = "", "", Environment.NewLine & causaCopia))
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
            ' ⛔ MISMA EXCEPCION QUE EN `Escribir`, MISMA LEY: un destino NUEVO no se borra cuando lo unico
            ' que fallo fue la guarda del contrato — el cuerpo volvio normalmente y esos bytes son los
            ' suyos. (En la practica este brazo solo se alcanza con un destino que NO existia, porque si
            ' existia con contenido hay `copia` y manda la restauracion de abajo, que NO se toca: alli
            ' devolverle al usuario su version anterior es lo correcto pase lo que pase.)
            If seToco AndAlso Not existia AndAlso Not (TypeOf ex Is ContratoDelCuerpoException) Then Borrar(destino)
            If copiaHecha Then
                ' Solo se restaura desde la copia que tomo ESTA corrida: es la unica de la que sabemos
                ' que contenido tiene. La heredada se deja donde esta.
                '
                ' ⛔⛔ LA RESTAURACION USA LA MISMA PRIMITIVA QUE LA ESCRITURA, Y ESTO ERA UN DEFECTO REAL.
                ' Acá habia `File.Copy(copia, destino, overwrite:=True)`. `File.Copy` pide CREATE_ALWAYS,
                ' y CREATE_ALWAYS sobre un destino OCULTO da ERROR_ACCESS_DENIED — la MISMA trampa que la
                ' cabecera de esta clase documenta para la escritura y que `EscribirNucleo` resuelve con
                ' OpenOrCreate. O sea que el camino de ida estaba blindado y el de vuelta no:
                ' <b>exactamente sobre los archivos ocultos para los que este diseño existe, la
                ' restauracion automatica que promete el docstring NO funcionaba.</b> El destino quedaba
                ' con lo que el cuerpo alcanzo a escribir (PARCIAL) y el usuario recibia el mensaje de
                ' "no se pudo restaurar" — su version anterior sobrevivia en la copia, pero la promesa
                ' era falsa. Afecta a todo GuardarConCopia: ESP, NIF, materiales, OSP/OSD, BodyGen y
                ' sidecars. Lo dejan OneDrive y los desempaquetadores, asi que no es raro.
                '
                ' MEDIDO en net8.0.30 (gate: caso M de EscrituraEnElLugarGate, que lo remide en cada
                ' corrida en vez de dejarlo escrito en un comentario que puede envejecer):
                '   File.Copy(-> destino OCULTO, overwrite:=True) .... UnauthorizedAccessException
                '   OpenOrCreate + SetLength(0) sobre OCULTO ......... OK, y CONSERVA el atributo
                '
                ' Se abre la COPIA primero y recien despues se toca el destino: es la misma forma que usa
                ' VolcarEncima, y evita que un fallo abriendo la copia deje el destino ya truncado.
                ' sincronizar:=True por el mismo motivo que el resto de este metodo: es el dato del
                ' usuario volviendo a su lugar.
                Try
                    Dim seTocoRestaurando As Boolean = False
                    Using fsCopia As New FileStream(copia, FileMode.Open, FileAccess.Read, FileShare.Read)
                        EscribirNucleo(destino,
                                       Sub(fsDestino) fsCopia.CopyTo(fsDestino),
                                       seTocoRestaurando,
                                       sincronizar:=True)
                    End Using
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
    ''' ese instante; la senal publica no. No hay medicion de que ese fallo ocurra.</para>
    ''' <para>⛔⛔ <paramref name="sincronizar"/> DEFAULTEA EN <b>True</b>, al reves que en
    ''' <see cref="Escribir"/>, y el motivo es el patron de los llamadores: <b>los tres que borran la otra
    ''' copia inmediatamente despues</b> — <c>ArchivePackager</c> (borra el `.new` y su sello),
    ''' <c>FomodExporter</c> (borra el `.tmp`) y el "move project" de Wardrobe Manager (borra el `.osp`
    ''' ORIGEN). En los tres, apenas vuelve esta llamada el destino pasa a ser <b>el unico ejemplar</b>. Sin
    ''' el flush, el borrado de la otra copia puede llegar al plato ANTES que los bytes que la reemplazan y
    ''' un corte de luz se lleva las dos — que es exactamente el estado que toda esta clase existe para
    ''' evitar, y el mismo que documenta el «la red solo es red si llego al plato» de
    ''' <see cref="GuardarConCopia"/>.</para>
    ''' <para>Va por DEFAULT y no por pedido porque lo peligroso tiene que ser lo que exige justificacion
    ''' explicita, no al reves: un llamador nuevo que vuelque y borre hereda la garantia sin saber que
    ''' existe. El opt-out es para el camino MEDIDO que no puede pagarlo — hoy uno solo:
    ''' <c>OSP_Clases.CommitTextureJobs</c>, cientos de archivos por corrida, que ademas NO borra el origen
    ''' (el origen es el suelto del juego) y por lo tanto no necesita la garantia. Costo del flush, medido
    ''' en <see cref="Escribir"/>: +2,66 a +4,53 ms POR LLAMADA — despreciable en un volcado de archive de
    ''' GiB, caro en un lote de cientos de texturas.</para></summary>
    Public Shared Sub VolcarEncima(origen As String, destino As String,
                                   Optional reintentos As Integer = 10,
                                   Optional esperaMs As Integer = 200,
                                   Optional alTocarElDestino As Action = Nothing,
                                   Optional sincronizar As Boolean = True)
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
                                       seToco,
                                       sincronizar)
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
    ''' antes — es estrictamente menos que el `File.Copy` que el metodo ya paga.</para>
    ''' <para>⛔ ES PUBLICO A PROPOSITO, y por eso el predicado vive ACA y no se reimplementa: la misma
    ''' ley de retencion —"un respaldo se descarta SOLO cuando se PROBO byte por byte que no guarda nada
    ''' propio"— la necesita <c>FomodExporter</c> para sus <c>.recovered</c>. Dos implementaciones del
    ''' mismo predicado es como se empiezan a borrar ejemplares unicos por un lado y no por el otro.</para></summary>
    Public Shared Function MismoContenido(a As String, b As String) As Boolean
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

    ''' <summary>Le saca a un RESPALDO recien copiado los atributos que lo esconden o lo congelan.
    ''' <para>⛔ HACE FALTA PORQUE <c>File.Copy</c> PROPAGA LOS ATRIBUTOS DEL ORIGEN. MEDIDO en net8.0.30:
    ''' copiar un archivo OCULTO deja la copia OCULTA, y uno de SOLO LECTURA la deja de SOLO LECTURA. Sin
    ''' esto, el respaldo del dato del usuario le queda INVISIBLE justo cuando lo necesita, y ademas un
    ''' respaldo de solo lectura no se puede borrar despues.</para>
    ''' <para>⛔ SOLO SE APLICA A COPIAS QUE ESTA CLASE (o el packer) ACABA DE CREAR. NUNCA al archivo del
    ''' usuario: sacarle el oculto a un archivo suyo seria una ley nueva, y la cabecera de
    ''' <see cref="EscribirNucleo"/> es explicita en que la escritura CONSERVA el atributo.</para>
    ''' <para>Es publico para que el `.bak.unpack` de <c>ArchivePackager.Unpack</c> —el otro respaldo del
    ''' arbol que nace de un <c>File.Copy</c>— use ESTA ley y no una segunda copiada al lado.</para></summary>
    Public Shared Sub LimpiarAtributos(ruta As String)
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
