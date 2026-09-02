Imports System.IO
Imports System.Text.Json

''' <summary>El DIARIO de un lote: lo único que sobrevive a que el proceso muera a mitad del guardado.
'''
''' <para>⛔ POR QUE EXISTE. `LoteConCopias` deshace ante una EXCEPCIÓN, pero si el proceso se termina de
''' golpe —cierre, corte de luz, kill— no corre ningún `Finally`: quedan las copias en disco y NADIE sabe
''' que pertenecían al mismo guardado, ni cuáles se habían escrito, ni si el lote había llegado a
''' confirmarse. `CopiasPendientes` puede nombrar archivos sueltos, pero no puede contestar ninguna de esas
''' tres preguntas. El diario sí, porque se escribe ANTES de tocar el destino y se sincroniza.</para>
'''
''' <para>⛔ NO VIVE EN `Data`. Es metadata de la aplicación, no del mod: va en la carpeta local de la app.
''' Meterlo al lado de los archivos del usuario lo pondría bajo el VFS de MO2 y ensuciaría el mod.</para>
'''
''' <para>⛔ Y NO RESTAURA SOLO. Después de un reinicio no hay forma de saber si el usuario quiere volver a
''' la versión anterior o conservar lo nuevo — la decisión es suya. El diario DETECTA y DESCRIBE; el
''' diálogo del arranque decide. Ver <see cref="EscrituraEnElLugar.DiariosPendientes"/>.</para></summary>
Public NotInheritable Class DiarioDeEscritura

    ''' <summary>Una operación del lote. <c>Operacion</c> es "crear", "reemplazar" o "borrar": son las tres
    ''' que el rollback distingue, y sin ese dato una copia suelta no dice qué había que hacer con ella.</summary>
    Public Class Entrada
        Public Property Destino As String
        Public Property Copia As String
        Public Property Operacion As String
        Public Property Atributos As Integer
    End Class

    Public Const SufijoDiario As String = ".npcm-txn.json"

    Private ReadOnly _ruta As String
    Private ReadOnly _entradas As New List(Of Entrada)

    ''' <summary>Ruta del archivo de diario en disco.</summary>
    Public ReadOnly Property Ruta As String
        Get
            Return _ruta
        End Get
    End Property

    Public ReadOnly Property Entradas As IReadOnlyList(Of Entrada)
        Get
            Return _entradas
        End Get
    End Property

    ''' <summary>Abre un diario nuevo y lo persiste vacío: que el archivo EXISTA antes de la primera
    ''' operación es lo que permite detectar un lote que murió entre la primera copia y su registro.</summary>
    Public Sub New(carpetaDiarios As String)
        Directory.CreateDirectory(carpetaDiarios)
        _ruta = Path.Combine(carpetaDiarios, Guid.NewGuid().ToString("N") & SufijoDiario)
        Persistir()
    End Sub

    ''' <summary>Lee un diario que quedó de una corrida anterior. Devuelve Nothing si el archivo no se puede
    ''' interpretar — un diario ilegible NO es una transacción recuperable y no se puede inventar.</summary>
    Public Shared Function Leer(ruta As String) As DiarioDeEscritura
        Try
            Dim json = File.ReadAllText(ruta)
            Dim ent = JsonSerializer.Deserialize(Of List(Of Entrada))(json)
            If ent Is Nothing Then Return Nothing
            Return New DiarioDeEscritura(ruta, ent)
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub New(ruta As String, entradas As List(Of Entrada))
        _ruta = ruta
        _entradas = entradas
    End Sub

    ''' <summary>⛔ SE LLAMA DESPUES DE ASEGURAR LA COPIA Y ANTES DE TOCAR EL DESTINO. Al revés el diario
    ''' describiría un estado que todavía no ocurrió, o —peor— el destino cambiaría sin que nadie lo haya
    ''' anotado. Para una CREACIÓN se registra antes de crear el archivo; para un BORRADO, antes del
    ''' `File.Delete` y con los atributos originales adentro.</summary>
    Public Sub Registrar(destino As String, copia As String, operacion As String, atributos As FileAttributes)
        _entradas.Add(New Entrada With {
            .Destino = destino,
            .Copia = copia,
            .Operacion = operacion,
            .Atributos = CInt(atributos)
        })
        Persistir()
    End Sub

    ''' <summary>El lote terminó (confirmado, o deshecho por completo): el diario ya no describe nada
    ''' pendiente y se va. Best-effort a propósito: un diario huérfano que no se pudo borrar hace que el
    ''' arranque siguiente OFREZCA una recuperación de más, que es molesto pero inofensivo — perder el
    ''' diario de un lote que sí quedó a medias sería lo contrario.</summary>
    Public Sub Cerrar()
        Try
            If File.Exists(_ruta) Then File.Delete(_ruta)
        Catch
        End Try
    End Sub

    ''' <summary>Escribe y SINCRONIZA. Sin el flush el diario puede no existir justo en el corte que vino a
    ''' documentar, que es el único momento en que sirve.</summary>
    Private Sub Persistir()
        Dim bytes = JsonSerializer.SerializeToUtf8Bytes(_entradas)
        Using fs As New FileStream(_ruta, FileMode.Create, FileAccess.Write, FileShare.Read)
            fs.Write(bytes, 0, bytes.Length)
            fs.Flush(flushToDisk:=True)
        End Using
    End Sub
End Class
