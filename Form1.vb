Imports ADODB
Imports System.Threading
Imports System.Collections.Generic

Public Class Form1

    Private Class ContactoAgenda
        Public Id As Integer
        Public Nombre As String
        Public Puesto As String
        Public Mail As String
        Public TipoAudiencia As String   ' TI / No-TI
        Public RolNivel As String        ' CEO / C-Level / TI / Gerencia / Staff
        Public Segmento As String        ' segmentoavital de la empresa
    End Class


    Private Function ObtenerHoraPorRol(rol As String) As String
        Select Case rol
            Case "CEO"
                Return "07:00"
            Case "C-Level"
                Return "09:30"
            Case "TI"
                Return "11:00"
            Case "Gerencia"
                Return "12:30"
            Case Else
                Return "13:30" ' Staff u otros
        End Select
    End Function


    ' ⬇️ AQUÍ, justo después de abrir la clase Form1
    Private Class ContactoEmpresaInfo
        Public Property Id As Integer
        Public Property Nombre As String
        Public Property Puesto As String
        Public Property Mail As String
        Public Property Segmento As String
        Public Property Rol As String
        Public Property TipoAudiencia As String
    End Class

    ' ⬇️ Después de esto, ya van tus botones, funciones, timers, etc.


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btniniciar.Click
        Timer1.Enabled = True
        lblestado.Text = "🟢 MOTOR EN MARCHA - Cada 2 min"
        EscribirLog("=== SISTEMA INICIADO ===")
        EjecutarCicloCompleto()


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configurar conexión
        rutabd = "provider=microsoft.ACE.oledb.12.0; data source=C:\Aplicacionesvb\AvitalMarketing\prospectos.accdb; persist security info=false;jet oledb:database password=rana147852"
        ba = New ADODB.Connection

        ' Configurar timer
        Timer1.Interval = 120000 ' 2 minutos
        Timer1.Enabled = False

        lblestado.Text = "✅ Sistema listo - Click en INICIAR"

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnLimpiarLog.Click
        txtlog.Clear()
    End Sub

    Private Sub btnDetener_Click(sender As Object, e As EventArgs) Handles btnDetener.Click
        Timer1.Enabled = False
        lblestado.Text = "⏸️ SISTEMA DETENIDO"
        EscribirLog("=== SISTEMA DETENIDO ===")
    End Sub

    Private Sub btnCicloManual_Click(sender As Object, e As EventArgs) Handles btnCicloManual.Click
        EscribirLog("=== PRUEBA CONTROLADA INICIADA ===")

        ' 1. Verificar conexión
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
                EscribirLog("✅ BD Conectada")
            Else
                EscribirLog("ℹ️ BD Ya estaba conectada")
            End If
        Catch ex As Exception
            EscribirLog("❌ Error BD: " & ex.Message)
            Return
        End Try

        ' 2. Verificar tablas
        EscribirLog("📋 Verificando tablas...")
        VerificarTablas()

        ' 3. Buscar UNA empresa de prueba
        EscribirLog("🔍 Buscando empresa de prueba...")
        BuscarUnaEmpresaPrueba()

        ' 4. Mostrar estado final
        EscribirLog("=== PRUEBA COMPLETADA ===")
    End Sub

    Private Sub btnVerEstado_Click(sender As Object, e As EventArgs) Handles btnVerEstado.Click
        Dim estado = ""
        estado &= "📊 ESTADO DEL SISTEMA" & vbCrLf
        estado &= "=====================" & vbCrLf
        estado &= "Hora actual: " & Date.Now.ToString("HH:mm:ss") & vbCrLf
        estado &= "Día de semana: " & Date.Now.DayOfWeek.ToString & vbCrLf
        estado &= "Horario ENVIAR permitido ahora: " & IIf(EsHorarioPermitidoEnvio(), "SÍ", "NO") & vbCrLf
        estado &= "Horario PROGRAMAR Mensaje 1 permitido ahora: " & IIf(EsHorarioPermitidoMensaje1(), "SÍ", "NO") & vbCrLf
        estado &= "Límite diario: " & limite_diario_envios & vbCrLf
        estado &= "Timer activo: " & IIf(Timer1.Enabled, "SÍ", "NO") & vbCrLf

        Try
            If ba.State <> 1 Then ba.Open(rutabd)
            Dim rs As New Recordset
            rs.Open("SELECT COUNT(*) as total FROM envios_programados WHERE enviado = False", ba)
            If Not rs.EOF Then estado &= "Envíos pendientes: " & rs.Fields("total").Value & vbCrLf
            rs.Close()

            rs.Open("SELECT COUNT(*) as total FROM prospectos WHERE en_secuencia = True", ba)
            If Not rs.EOF Then estado &= "Contactos en secuencia: " & rs.Fields("total").Value & vbCrLf
            rs.Close()

            ba.Close()
        Catch ex As Exception
            estado &= "Error BD: " & ex.Message & vbCrLf
        End Try

        MessageBox.Show(estado, "Estado del Sistema")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        EjecutarCicloCompleto()
    End Sub


    Private Sub EjecutarCicloCompleto()
        ' 1) Conectar BD
        If ba.State <> 1 Then
            Try
                ba.Open(rutabd)
            Catch ex As Exception
                EscribirLog("❌ Error conectando BD: " & ex.Message)
                Return
            End Try
        End If

        ' 2) Programar NUEVOS (Mensaje 1) SOLO dentro de ventana permitida de Mensaje 1
        If EsHorarioPermitidoMensaje1() Then
            ProgramarHoyYBuffer_Mensaje1()
        End If

        ' 3) Enviar 1 correo pendiente (sin tandas, sin sleeps)
        If EsHorarioPermitidoEnvio() Then
            EnviarUnPendiente_1x1()
        End If

        ' 4) Estado
        Dim pHoy As Integer = ContarProgramadosPorDia(DateTime.Now.Date)
        lblestado.Text = "🟢 Activo | Programados hoy: " & pHoy & " | Hora: " & DateTime.Now.ToString("HH:mm")
    End Sub

    ' ==========================
    ' Horarios permitidos
    ' ==========================
    Private Function EsHorarioPermitidoMensaje1() As Boolean
        ' Mensaje 1: SOLO Mar-Jue 06:40–14:00
        Dim ahora As DateTime = DateTime.Now
        Dim dia As DayOfWeek = ahora.DayOfWeek
        Dim h As String = ahora.ToString("HH:mm")

        If dia < DayOfWeek.Tuesday OrElse dia > DayOfWeek.Thursday Then Return False
        If h < hora_inicio_envios OrElse h > hora_fin_envios Then Return False
        Return True
    End Function

    Private Function EsHorarioPermitidoEnvio() As Boolean
        ' Ventana general de envíos:
        ' - Mar-Jue 06:40–14:00
        ' - Lunes 12:00–15:00 (para mensajes 2-5 NO CEO; los CEO no deberían estar programados ahí)
        Dim ahora As DateTime = DateTime.Now
        Dim dia As DayOfWeek = ahora.DayOfWeek
        Dim h As String = ahora.ToString("HH:mm")

        If dia >= DayOfWeek.Tuesday AndAlso dia <= DayOfWeek.Thursday Then
            Return (h >= hora_inicio_envios AndAlso h <= hora_fin_envios)
        End If

        If dia = DayOfWeek.Monday Then
            Return (h >= hora_inicio_lunes AndAlso h <= hora_fin_lunes)
        End If

        Return False
    End Function

    ' ========== BUSCAR NUEVOS CONTACTOS ==========
    Private Sub BuscarYProgramarNuevosContactos(capacidad As Integer)
        EscribirLog("🔍 Buscando nuevos contactos...")

        Try
            ' Buscar empresas elegibles
            Dim rsEmpresas As New Recordset
            Dim sqlEmpresas As String =
                "SELECT TOP 3 e.cia " &
                "FROM listaempresas e " &
                "INNER JOIN (" &
                "   SELECT cia, COUNT(*) as total_contactos " &
                "   FROM prospectos " &
                "   WHERE en_secuencia = False AND mail IS NOT NULL AND mailincorrecto = False " &
                "   GROUP BY cia " &
                ") p ON e.cia = p.cia " &
                "WHERE e.empleados < 900 " &
                "AND e.megalopolis = True " &
                "AND e.enterprise = False " &
                "AND e.conerror = False " &
                "AND e.con_contactos_en_secuencia = False " &
                "AND p.total_contactos >= 3 " &
                "ORDER BY e.empleados DESC"

            rsEmpresas.Open(sqlEmpresas, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            Dim empresasProcesadas As Integer = 0

            If Not rsEmpresas.EOF Then
                rsEmpresas.MoveFirst()
                While Not rsEmpresas.EOF And empresasProcesadas < 3
                    Dim empresa As String = rsEmpresas.Fields("cia").Value.ToString()
                    EscribirLog("🏢 Procesando empresa: " & empresa)

                    ' Procesar contactos de esta empresa
                    ProcesarContactosEmpresa(empresa)

                    empresasProcesadas += 1
                    rsEmpresas.MoveNext()
                End While
            Else
                EscribirLog("ℹ️ No hay empresas nuevas que cumplan criterios")
            End If

            rsEmpresas.Close()

        Catch ex As Exception
            EscribirLog("❌ Error buscando contactos: " & ex.Message)
        End Try
    End Sub

    ' ========== PROCESAR CONTACTOS DE UNA EMPRESA ==========
    Private Sub ProcesarContactosEmpresa(empresa As String)
        Try
            ' Buscar contactos de esta empresa (máximo 3)
            Dim rsContactos As New Recordset
            Dim sqlContactos As String =
                    "SELECT p.id, p.nombre, p.puesto, p.mail, e.segmentoavital " &
                    "FROM prospectos p " &
                    "INNER JOIN listaempresas e ON p.cia = e.cia " &
                    "WHERE p.cia = '" & empresa.Replace("'", "''") & "' " &
                    "AND p.en_secuencia = False " &
                    "AND p.mail IS NOT NULL " &
                    "AND p.mailincorrecto = False"


            rsContactos.Open(sqlContactos, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If Not rsContactos.EOF Then
                rsContactos.MoveFirst()
                While Not rsContactos.EOF
                    Dim contactoId As Integer = rsContactos.Fields("id").Value
                    Dim nombre As String = rsContactos.Fields("nombre").Value.ToString()
                    Dim puesto As String = rsContactos.Fields("puesto").Value.ToString()
                    Dim email As String = rsContactos.Fields("mail").Value.ToString()
                    Dim segmento As String = rsContactos.Fields("segmentoavital").Value.ToString()

                    ' Determinar rol
                    Dim rolInfo = DeterminarRolYHorario(puesto)
                    Dim tipoAudiencia As String = IIf(rolInfo.Contains("TI"), "TI", "No-TI")

                    EscribirLog("   👤 " & nombre & " - " & puesto & " (" & rolInfo & ")")

                    ' Buscar secuencia
                    Dim secuenciaId As Integer = ObtenerSecuenciaId(segmento, tipoAudiencia)

                    If secuenciaId > 0 Then
                        ' Programar 5 correos
                        ProgramarSecuencia(contactoId, secuenciaId, rolInfo)

                        ' Marcar como en secuencia
                        MarcarContactoEnSecuencia(contactoId, secuenciaId)

                        EscribirLog("     ✅ Programado (Secuencia ID: " & secuenciaId & ")")
                    Else
                        EscribirLog("     ⚠️ Sin secuencia para " & segmento & " - " & tipoAudiencia)
                    End If

                    rsContactos.MoveNext()
                End While

                ' Marcar empresa como activada
                MarcarEmpresaActivada(empresa)
            End If

            rsContactos.Close()

        Catch ex As Exception
            EscribirLog("❌ Error procesando empresa " & empresa & ": " & ex.Message)
        End Try
    End Sub

    ' ========== DETERMINAR ROL ==========
    Private Function DeterminarRolYHorario(puesto As String) As String
        Dim puestoLower As String = puesto.ToLower()

        If puestoLower.Contains("ceo") Or puestoLower.Contains("founder") Or puestoLower.Contains("dueño") Then
            Return "CEO"
        ElseIf puestoLower.Contains("director") Or puestoLower.Contains("c-level") Or puestoLower.Contains("vp") Then
            Return "C-Level"
        ElseIf puestoLower.Contains("ti") Or puestoLower.Contains("tecnolog") Or puestoLower.Contains("cto") Or puestoLower.Contains("it") Or puestoLower.Contains("sistemas") Then
            Return "TI"
        ElseIf puestoLower.Contains("gerente") Or puestoLower.Contains("jefe") Or puestoLower.Contains("manager") Then
            Return "Gerencia"
        Else
            Return "Staff"
        End If
    End Function

    ' ========== OBTENER SECUENCIA ID ==========
    Private Function ObtenerSecuenciaId(segmento As String, tipoAudiencia As String, Optional rolNivel As String = "") As Integer
        Try
            ' Buscar el segmento_id basado en el nombre del segmento
            Dim segmentoId As Integer = 0
            Dim rsSeg As New Recordset

            ' El segmento viene como "Clínicas / Hospitales / Laboratorios"
            ' Buscamos en la tabla segmentacion
            Dim sqlSeg As String = "SELECT id FROM segmentacion WHERE segmento = '" & segmento & "'"
            rsSeg.Open(sqlSeg, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If Not rsSeg.EOF Then
                segmentoId = rsSeg.Fields("id").Value
                EscribirLog("   Segmento ID encontrado: " & segmentoId)
            Else
                EscribirLog("   ⚠️ Segmento no encontrado: " & segmento)
                rsSeg.Close()
                Return 0
            End If
            rsSeg.Close()

            ' Buscar secuencia para este segmento y tipo de audiencia
            Dim rsSec As New Recordset
            Dim sqlSec As String =
            "SELECT id FROM secuencias_maestra " &
            "WHERE segmento_id = " & segmentoId & " " &
            "AND tipo_audiencia = '" & tipoAudiencia.Replace("'", "''") & "' " &
            IIf(String.IsNullOrWhiteSpace(rolNivel), "", "AND rol_nivel = '" & rolNivel.Replace("'", "''") & "' ") &
            "AND activa = True"

            rsSec.Open(sqlSec, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If Not rsSec.EOF Then
                Dim idEncontrado As Integer = rsSec.Fields("id").Value
                EscribirLog("   ✅ Secuencia ID encontrada: " & idEncontrado)
                rsSec.Close()
                Return idEncontrado
            Else
                EscribirLog("   ⚠️ No hay secuencia para: Segmento=" & segmentoId & ", Audiencia=" & tipoAudiencia)
            End If

            rsSec.Close()

        Catch ex As Exception
            EscribirLog("❌ Error en ObtenerSecuenciaId: " & ex.Message)
        End Try

        Return 0
    End Function

    ' ========== PROGRAMAR SOLO DÍA 1 ==========
    Private Sub ProgramarSecuencia(contactoId As Integer, secuenciaId As Integer, rol As String)
        Try
            ' SOLO DÍA 1 - NO PROGRAMAR DÍAS 2-5
            Dim diaNumero As Integer = 1

            ' Obtener mensaje para día 1
            Dim rsMsg As New Recordset
            rsMsg.Open("SELECT id FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId & " AND dia_numero = " & diaNumero, ba)

            If Not rsMsg.EOF Then
                Dim mensajeId As Integer = rsMsg.Fields("id").Value
                rsMsg.Close()

                ' Obtener hora según rol (con aleatoriedad ya incluida en ProgramarPrimerCorreo)
                Dim horaEnvio As String = "09:00" ' Default
                Select Case rol
                    Case "CEO" : horaEnvio = "07:00"
                    Case "C-Level" : horaEnvio = "09:30"
                    Case "TI" : horaEnvio = "11:00"
                    Case "Gerencia" : horaEnvio = "12:30"
                    Case "Staff" : horaEnvio = "13:30"
                End Select

                ' Calcular fecha base (hoy, ajustado a martes-jueves)
                Dim fechaBase As DateTime = DateTime.Today
                While fechaBase.DayOfWeek < DayOfWeek.Tuesday Or fechaBase.DayOfWeek > DayOfWeek.Thursday
                    fechaBase = fechaBase.AddDays(1)
                End While

                ' Combinar fecha con hora
                Dim fechaCompleta As DateTime = DateTime.Parse(fechaBase.ToString("yyyy-MM-dd") & " " & horaEnvio)

                ' Insertar en programación
                Dim rsEnvio As New Recordset
                rsEnvio.Open("envios_programados", ba, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic)

                rsEnvio.AddNew()
                rsEnvio.Fields("contacto_id").Value = contactoId
                rsEnvio.Fields("secuencia_id").Value = secuenciaId
                rsEnvio.Fields("mensaje_id").Value = mensajeId
                rsEnvio.Fields("fecha_programada").Value = fechaCompleta
                rsEnvio.Fields("enviado").Value = False
                rsEnvio.Update()
                rsEnvio.Close()

                EscribirLog("   ✅ Programado SOLO DÍA 1 para contacto " & contactoId)
            Else
                rsMsg.Close()
            End If

        Catch ex As Exception
            EscribirLog("❌ Error programando secuencia: " & ex.Message)
        End Try
    End Sub

    ' ========== MARCAR CONTACTO EN SECUENCIA ==========
    Private Sub MarcarContactoEnSecuencia(contactoId As Integer, secuenciaId As Integer)
        Try
            Dim rs As New Recordset
            rs.Open("SELECT * FROM prospectos WHERE id = " & contactoId, ba, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic)

            If Not rs.EOF Then
                rs.Fields("en_secuencia").Value = True
                rs.Fields("fecha_entrada_secuencia").Value = DateTime.Now
                rs.Fields("secuencia_activa_id").Value = secuenciaId
                rs.Update()
            End If

            rs.Close()
        Catch ex As Exception
        End Try
    End Sub

    ' ========== MARCAR EMPRESA ACTIVADA ==========
    Private Sub MarcarEmpresaActivada(empresa As String)
        Try
            Dim rs As New Recordset
            rs.Open("SELECT * FROM listaempresas WHERE cia = '" & empresa.Replace("'", "''") & "'", ba, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic)

            If Not rs.EOF Then
                rs.Fields("con_contactos_en_secuencia").Value = True
                rs.Fields("fecha_activacion_secuencia").Value = DateTime.Now
                rs.Update()
            End If

            rs.Close()
        Catch ex As Exception
        End Try
    End Sub


    ' ==========================================================
    ' NUEVO MOTOR (Abraham):
    ' - Programa SOLO Mensaje 1 respetando 110/día + buffer 25
    ' - Selección por listaempresas (personas/mixto/enviarsinmixto)
    ' - Programación por rol con roles_horarios
    ' - Gaps humanos 40-300s (se refleja en fecha_programada)
    ' ==========================================================

    Private Function ContarProgramadosPorDia(fechaDia As DateTime) As Integer
        Try
            Dim inicio As DateTime = fechaDia.Date
            Dim fin As DateTime = fechaDia.Date.AddDays(1)
            Dim rs As New ADODB.Recordset
            Dim sql As String =
                "SELECT COUNT(*) as total FROM envios_programados " &
                "WHERE fecha_programada >= #" & inicio.ToString("yyyy-MM-dd HH:mm:ss") & "# " &
                "AND fecha_programada < #" & fin.ToString("yyyy-MM-dd HH:mm:ss") & "#;"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Dim total As Integer = 0
            If Not rs.EOF Then total = CInt(rs.Fields("total").Value)
            rs.Close()
            Return total
        Catch
            Return 0
        End Try
    End Function

    Private Function EsDiaPermitidoMensaje1(fechaDia As DateTime) As Boolean
        Dim d As DayOfWeek = fechaDia.DayOfWeek
        Return (d = DayOfWeek.Tuesday OrElse d = DayOfWeek.Wednesday OrElse d = DayOfWeek.Thursday)
    End Function

    Private Function SiguienteDiaPermitidoMensaje1(desde As DateTime) As DateTime
        Dim f As DateTime = desde.Date
        For i As Integer = 0 To 14
            If EsDiaPermitidoMensaje1(f) Then Return f
            f = f.AddDays(1)
        Next
        Return desde.Date
    End Function

    Private Function ObtenerHorarioRol(rol As String) As (HoraInicio As TimeSpan, HoraFin As TimeSpan, Activo As Boolean)
        ' 1) Intento exacto
        Dim rs As New ADODB.Recordset
        Dim r As String = rol.Replace("'", "''")
        Dim sql As String = "SELECT TOP 1 hora_inicio, hora_fin, activo FROM roles_horarios WHERE rol = '" & r & "'"
        rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rs.EOF Then
            Dim hi As TimeSpan = TimeSpan.Parse(CStr(rs.Fields("hora_inicio").Value))
            Dim hf As TimeSpan = TimeSpan.Parse(CStr(rs.Fields("hora_fin").Value))
            Dim act As Boolean = CBool(rs.Fields("activo").Value)
            rs.Close()
            Return (hi, hf, act)
        End If
        rs.Close()

        ' 2) Intento flexible (por prefijo)
        Dim pref As String = rol
        If pref.Contains("/") Then pref = pref.Split("/"c)(0).Trim()
        If pref.Length > 12 Then pref = pref.Substring(0, 12)
        pref = pref.Replace("'", "''")

        sql = "SELECT TOP 1 hora_inicio, hora_fin, activo FROM roles_horarios WHERE rol LIKE '" & pref & "%'"
        rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rs.EOF Then
            Dim hi As TimeSpan = TimeSpan.Parse(CStr(rs.Fields("hora_inicio").Value))
            Dim hf As TimeSpan = TimeSpan.Parse(CStr(rs.Fields("hora_fin").Value))
            Dim act As Boolean = CBool(rs.Fields("activo").Value)
            rs.Close()
            Return (hi, hf, act)
        End If
        rs.Close()

        Return (TimeSpan.Parse("09:00"), TimeSpan.Parse("10:00"), False)
    End Function

    Private Function SiguienteFechaConGap(ultima As DateTime, inicio As DateTime, fin As DateTime, rnd As Random) As DateTime
        Dim gap As Integer = rnd.Next(gap_min_seg, gap_max_seg + 1)
        Dim candidata As DateTime

        If ultima < inicio Then
            candidata = inicio.AddSeconds(rnd.Next(gap_min_seg, gap_max_seg + 1))
        Else
            candidata = ultima.AddSeconds(gap)
        End If

        If candidata > fin Then candidata = fin.AddSeconds(-5)
        Return candidata
    End Function

    Private Function ObtenerEmpresasElegiblesMensaje1() As List(Of String)
        Dim empresas As New List(Of String)

        Dim sql As String =
            "SELECT e.cia " &
            "FROM listaempresas e " &
            "WHERE e.megalopolis = True " &
            "AND e.enterprise = False " &
            "AND e.conerror = False " &
            "AND e.con_contactos_en_secuencia = False " &
            "AND (e.enviarsinmixto = True OR (e.enviarsinmixto = False AND e.mixto = True AND e.personas >= 3)) " &
            "ORDER BY e.personas DESC;"

        Dim rs As New ADODB.Recordset
        rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        While Not rs.EOF
            empresas.Add(CStr(rs.Fields("cia").Value))
            rs.MoveNext()
        End While
        rs.Close()
        Return empresas
    End Function

    Private Function ObtenerProspectosElegiblesDeEmpresa(cia As String) As List(Of ContactoAgenda)
        Dim lista As New List(Of ContactoAgenda)

        Dim sql As String =
            "SELECT p.id, p.nombre, p.puesto, p.mail, e.segmentoavital " &
            "FROM prospectos p " &
            "INNER JOIN listaempresas e ON p.cia = e.cia " &
            "WHERE p.cia = '" & cia.Replace("'", "''") & "' " &
            "AND p.en_secuencia = False " &
            "AND p.mail IS NOT NULL " &
            "AND p.mailincorrecto = False;"

        Dim rs As New ADODB.Recordset
        rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        While Not rs.EOF
            Dim c As New ContactoAgenda
            c.Id = CInt(rs.Fields("id").Value)
            c.Nombre = If(IsDBNull(rs.Fields("nombre").Value), "", CStr(rs.Fields("nombre").Value))
            c.Puesto = If(IsDBNull(rs.Fields("puesto").Value), "", CStr(rs.Fields("puesto").Value))
            c.Mail = If(IsDBNull(rs.Fields("mail").Value), "", CStr(rs.Fields("mail").Value))
            c.Segmento = If(IsDBNull(rs.Fields("segmentoavital").Value), "", CStr(rs.Fields("segmentoavital").Value))
            lista.Add(c)
            rs.MoveNext()
        End While

        rs.Close()
        Return lista
    End Function

    Private Function ObtenerSegmentacionDesdeTabla(puesto As String, ByRef tipoAudiencia As String, ByRef rolNivel As String) As Boolean
        tipoAudiencia = ""
        rolNivel = ""

        If String.IsNullOrWhiteSpace(puesto) Then Return False

        Dim p As String = puesto.Trim().Replace("'", "''")
        Dim rs As New ADODB.Recordset
        Dim sql As String =
            "SELECT TOP 1 tipoaudiencia, rol_nivel FROM segmentacion_personas WHERE cargo = '" & p & "'"

        rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rs.EOF Then
            tipoAudiencia = If(IsDBNull(rs.Fields("tipoaudiencia").Value), "", CStr(rs.Fields("tipoaudiencia").Value))
            rolNivel = If(IsDBNull(rs.Fields("rol_nivel").Value), "", CStr(rs.Fields("rol_nivel").Value))
            rs.Close()
            Return (Not String.IsNullOrWhiteSpace(tipoAudiencia) AndAlso Not String.IsNullOrWhiteSpace(rolNivel))
        End If
        rs.Close()
        Return False
    End Function

    Private Sub ProgramarHoyYBuffer_Mensaje1()
        ' Programa SOLO Mensaje 1.
        ' HOY: llena hasta 110
        ' BUFFER: 25 para el siguiente día permitido (Mensaje 1)
        Try
            Dim hoy As DateTime = DateTime.Now.Date
            If Not EsDiaPermitidoMensaje1(hoy) Then Exit Sub

            Dim siguientePermitido As DateTime = SiguienteDiaPermitidoMensaje1(hoy.AddDays(1))

            Dim programadosHoy As Integer = ContarProgramadosPorDia(hoy)
            Dim programadosSiguiente As Integer = ContarProgramadosPorDia(siguientePermitido)

            Dim capHoy As Integer = limite_diario_envios - programadosHoy
            Dim capBuffer As Integer = buffer_manana - programadosSiguiente

            If capHoy <= 0 AndAlso capBuffer <= 0 Then Exit Sub

            EscribirLog($"📌 Programación: HOY cap={capHoy} | Buffer({siguientePermitido:yyyy-MM-dd}) cap={capBuffer}")

            Dim empresas As List(Of String) = ObtenerEmpresasElegiblesMensaje1()

            Dim ultimasHoy As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
            Dim ultimasBuf As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
            Dim rnd As New Random()

            For Each cia As String In empresas
                Dim listaPros As List(Of ContactoAgenda) = ObtenerProspectosElegiblesDeEmpresa(cia)
                If listaPros.Count = 0 Then Continue For

                Dim n As Integer = listaPros.Count

                ' Intentar HOY
                If capHoy > 0 AndAlso n <= capHoy Then
                    ProgramarEmpresaMensaje1EnDia(cia, listaPros, hoy, ultimasHoy, rnd)
                    capHoy -= n
                    MarcarEmpresaActivada(cia)
                    If capHoy <= 0 AndAlso capBuffer <= 0 Then Exit For
                    Continue For
                End If

                ' Intentar BUFFER (siguiente día permitido)
                If capBuffer > 0 AndAlso n <= capBuffer Then
                    ProgramarEmpresaMensaje1EnDia(cia, listaPros, siguientePermitido, ultimasBuf, rnd)
                    capBuffer -= n
                    MarcarEmpresaActivada(cia)
                    If capHoy <= 0 AndAlso capBuffer <= 0 Then Exit For
                    Continue For
                End If

                ' No cabe: seguimos con otra empresa (posible más pequeña)
            Next

        Catch ex As Exception
            EscribirLog("❌ Error ProgramarHoyYBuffer_Mensaje1: " & ex.Message)
        End Try
    End Sub

    Private Sub ProgramarEmpresaMensaje1EnDia(cia As String, prospectos As List(Of ContactoAgenda), fechaDia As DateTime,
                                             ByRef ultimasPorRol As Dictionary(Of String, DateTime), rnd As Random)

        EscribirLog($"🏢 Programando empresa {cia} en {fechaDia:yyyy-MM-dd} | Prospectos={prospectos.Count}")

        For Each c As ContactoAgenda In prospectos
            Dim tipo As String = ""
            Dim rol As String = ""
            If Not ObtenerSegmentacionDesdeTabla(c.Puesto, tipo, rol) Then
                EscribirLog($"   ⚠️ SIN MATCH en segmentacion_personas: puesto='{c.Puesto}' | contacto={c.Nombre}({c.Id})")
                ' Fallback mínimo: NO detenemos, pero lo marcamos con rol Staff (si existe) y NoTI
                tipo = "NoTI"
                rol = "Staff operativo"
            End If

            c.TipoAudiencia = tipo
            c.RolNivel = rol

            Dim secuenciaId As Integer = ObtenerSecuenciaId(c.Segmento, c.TipoAudiencia, c.RolNivel)
            If secuenciaId <= 0 Then
                EscribirLog($"   ⚠️ Sin secuencia: seg='{c.Segmento}' aud='{c.TipoAudiencia}' rol='{c.RolNivel}'")
                Continue For
            End If

            Dim h = ObtenerHorarioRol(c.RolNivel)
            If Not h.Activo Then
                EscribirLog($"   ⚠️ Rol sin horario activo en roles_horarios: '{c.RolNivel}'")
                Continue For
            End If

            Dim inicio As DateTime = fechaDia.Date.Add(h.HoraInicio)
            Dim fin As DateTime = fechaDia.Date.Add(h.HoraFin)

            Dim ultima As DateTime = DateTime.MinValue
            If ultimasPorRol.ContainsKey(c.RolNivel) Then ultima = ultimasPorRol(c.RolNivel)

            Dim fechaProg As DateTime = SiguienteFechaConGap(ultima, inicio, fin, rnd)
            ultimasPorRol(c.RolNivel) = fechaProg

            ProgramarMensaje1EnFecha(c.Id, secuenciaId, fechaProg)
            MarcarContactoEnSecuencia(c.Id, secuenciaId)
        Next
    End Sub

    Private Sub ProgramarMensaje1EnFecha(contactoId As Integer, secuenciaId As Integer, fechaProgramada As DateTime)
        Try
            Dim rsMsg As New ADODB.Recordset
            rsMsg.Open("SELECT TOP 1 id FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId & " AND dia_numero = 1",
                       ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsMsg.EOF Then
                rsMsg.Close()
                EscribirLog("   ⚠️ No hay mensaje Día 1 para secuencia " & secuenciaId)
                Exit Sub
            End If

            Dim mensajeId As Integer = CInt(rsMsg.Fields("id").Value)
            rsMsg.Close()

            Dim rsEnv As New ADODB.Recordset
            rsEnv.Open("envios_programados", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsEnv.AddNew()
            rsEnv.Fields("contacto_id").Value = contactoId
            rsEnv.Fields("secuencia_id").Value = secuenciaId
            rsEnv.Fields("mensaje_id").Value = mensajeId
            rsEnv.Fields("fecha_programada").Value = fechaProgramada
            rsEnv.Fields("enviado").Value = False
            rsEnv.Update()
            rsEnv.Close()

            EscribirLog("   ✅ Programado Día1 contacto " & contactoId & " | " & fechaProgramada.ToString("HH:mm:ss"))

        Catch ex As Exception
            EscribirLog("❌ Error ProgramarMensaje1EnFecha: " & ex.Message)
        End Try
    End Sub

    ' ==========================================================
    ' ENVÍO 1x1 (sin tandas, sin sleep)
    ' ==========================================================
    Private Sub EnviarUnPendiente_1x1()
        Try
            Dim ahora As DateTime = DateTime.Now

            Dim rs As New ADODB.Recordset
            Dim sql As String =
                "SELECT TOP 1 ep.id, ep.contacto_id, ep.mensaje_id, " &
                "p.nombrecompleto, p.mail, e.cia, m.asunto, m.cuerpo " &
                "FROM ((envios_programados ep " &
                "INNER JOIN prospectos p ON ep.contacto_id = p.id) " &
                "INNER JOIN listaempresas e ON p.cia = e.cia) " &
                "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
                "WHERE ep.fecha_programada <= #" & ahora.ToString("yyyy-MM-dd HH:mm:ss") & "# " &
                "AND ep.enviado = False " &
                "ORDER BY ep.fecha_programada;"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rs.EOF Then
                rs.Close()
                Exit Sub
            End If

            Dim contactoId As Integer = CInt(rs.Fields("contacto_id").Value)
            Dim mensajeId As Integer = CInt(rs.Fields("mensaje_id").Value)
            Dim nombre As String = If(IsDBNull(rs.Fields("nombrecompleto").Value), "", CStr(rs.Fields("nombrecompleto").Value))
            Dim mailReal As String = If(IsDBNull(rs.Fields("mail").Value), "", CStr(rs.Fields("mail").Value))
            Dim empresa As String = If(IsDBNull(rs.Fields("cia").Value), "", CStr(rs.Fields("cia").Value))
            Dim asuntoBase As String = If(IsDBNull(rs.Fields("asunto").Value), "", CStr(rs.Fields("asunto").Value))
            Dim cuerpoBase As String = If(IsDBNull(rs.Fields("cuerpo").Value), "", CStr(rs.Fields("cuerpo").Value))
            rs.Close()

            If String.IsNullOrWhiteSpace(mailReal) Then Exit Sub

            ' Cuenta de envío
            Dim rsCuenta As New ADODB.Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital WHERE limite > 0", ba,
                          ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCuenta.EOF Then
                rsCuenta.Close()
                Exit Sub
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa)
            Dim cuerpo As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            Dim mail As New Net.Mail.MailMessage
            mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
            mail.To.Add(mailReal)
            mail.Subject = asunto
            mail.Body = cuerpo
            mail.IsBodyHtml = True

            smtpClient.Send(mail)

            MarcarCorreoEnviado(contactoId, mensajeId)
            EscribirLog("✅ Enviado 1x1 a " & nombre & " (" & mailReal & ")")

        Catch ex As Exception
            EscribirLog("⚠️ Error EnviarUnPendiente_1x1: " & ex.Message)
        End Try
    End Sub


    ' ========== EJECUTAR ENVÍOS PENDIENTES ==========
    Private Sub EjecutarEnviosPendientes()
        ' Envío 1x1 (sin tandas, sin sleeps). La "humanidad" ya viene de fecha_programada.
        If Not EsHorarioPermitidoEnvio() Then Exit Sub
        EnviarUnPendiente_1x1()
    End Sub

    ' ========== ESCRIBIR EN LOG ==========
    Private Sub EscribirLog(mensaje As String)
        txtlog.AppendText(DateTime.Now.ToString("HH:mm:ss") & " - " & mensaje & vbCrLf)
        txtlog.ScrollToCaret()
    End Sub


    Private Sub VerificarTablas()
        Try
            Dim tablas() As String = {"secuencias_maestra", "mensajes", "envios_programados", "roles_horarios"}

            For Each tabla In tablas
                Dim rs As New Recordset
                Try
                    rs.Open("SELECT COUNT(*) as total FROM " & tabla, ba)
                    If Not rs.EOF Then
                        EscribirLog("   " & tabla & ": " & rs.Fields("total").Value & " registros")
                    End If
                    rs.Close()
                Catch ex As Exception
                    EscribirLog("   ❌ " & tabla & ": NO EXISTE o error")
                End Try
            Next

        Catch ex As Exception
            EscribirLog("⚠️ Error verificando tablas")
        End Try
    End Sub

    Private Sub BuscarUnaEmpresaPrueba()
        Try
            ' Buscar UNA empresa que cumpla criterios
            Dim rs As New Recordset
            Dim sql As String =
            "SELECT TOP 1 e.cia, COUNT(p.id) as contactos " &
            "FROM listaempresas e " &
            "INNER JOIN prospectos p ON e.cia = p.cia " &
            "WHERE e.empleados < 900 " &
            "AND e.megalopolis = True " &
            "AND e.enterprise = False " &
            "AND e.conerror = False " &
            "AND p.en_secuencia = False " &
            "GROUP BY e.cia " &
            "HAVING COUNT(p.id) >= 1 " &
            "ORDER BY e.cia"

            rs.Open(sql, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If Not rs.EOF Then
                Dim empresa As String = rs.Fields("cia").Value.ToString()
                Dim contactos As Integer = rs.Fields("contactos").Value

                EscribirLog("✅ Empresa encontrada: " & empresa)
                EscribirLog("   Contactos disponibles: " & contactos)

                ' Procesar SOLO esta empresa
                ProcesarContactosEmpresa(empresa)
            Else
                EscribirLog("ℹ️ No se encontraron empresas de prueba")
            End If

            rs.Close()

        Catch ex As Exception
            EscribirLog("❌ Error buscando empresa: " & ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        EscribirLog("=== PRUEBA DIRECTA A ABRAHAM ===")

        ' Conectar BD si no está conectada
        If ba.State <> 1 Then
            Try
                ba.Open(rutabd)
                EscribirLog("✅ BD Conectada para prueba")
            Catch ex As Exception
                EscribirLog("❌ Error BD: " & ex.Message)
                Return
            End Try
        End If

        ' Obtener UNA cuenta de mailsavital para prueba
        Try
            Dim rsCuenta As New Recordset
            rsCuenta.Open("SELECT TOP 1 * FROM mailsavital WHERE limite > 0", ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                EscribirLog("❌ No hay cuentas disponibles en mailsavital")
                Return
            End If

            Dim correoEnvio = rsCuenta.Fields("correo").Value.ToString
            Dim passwordEnvio = rsCuenta.Fields("password").Value.ToString
            Dim nombreRemitente = rsCuenta.Fields("nombre").Value.ToString

            rsCuenta.Close()

            ' Configurar SMTP (Mailersend como en tu código)
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            ' Crear correo de prueba
            Dim mail As New Net.Mail.MailMessage
            mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
            mail.To.Add("abraham@avital.mx")
            mail.Subject = "PRUEBA Sistema Automático - " & Date.Now.ToString("HH:mm:ss")

            Dim cuerpoHTML = "<html><body>" &
                "<h3>✅ Prueba Exitosa</h3>" &
                "<p>Este correo fue enviado desde el <strong>nuevo sistema automático de secuencias</strong>.</p>" &
                "<p>Fecha/Hora: " & Date.Now.ToString("dd/MM/yyyy HH:mm:ss") & "</p>" &
                "<p>Cuenta usada: " & correoEnvio & "</p>" &
                "<p>Si recibes esto, el sistema está funcionando correctamente.</p>" &
                "</body></html>"

            mail.Body = cuerpoHTML
            mail.IsBodyHtml = True

            ' Enviar
            smtpClient.Send(mail)

            EscribirLog("✅ Correo de prueba ENVIADO a Abraham")
            EscribirLog("   Desde: " & correoEnvio)
            EscribirLog("   Asunto: " & mail.Subject)

            MessageBox.Show("✅ Correo de prueba enviado a Abraham", "Prueba Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            EscribirLog("❌ Error en prueba: " & ex.Message)
            MessageBox.Show("❌ Error: " & ex.Message, "Error en Prueba", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        EscribirLog("=== PRUEBA COMPLETADA ===")
    End Sub



    ' ========== ENVIAR CORREO PLANEADO A ABRAHAM PARA REVISIÓN ==========
    Private Sub EnviarCorreoPlaneadoParaRevision(contactoId As Integer, diaNumero As Integer)
        Try
            ' 1. Obtener datos del contacto y su secuencia programada
            Dim rs As New Recordset
            Dim sql As String = "SELECT " &
                               "p.nombre, p.puesto, p.mail, " &
                               "e.cia, e.segmentoavital, " &
                               "ep.secuencia_id, ep.mensaje_id, ep.fecha_programada, " &
                               "m.asunto, m.cuerpo " &
                               "FROM envios_programados ep " &
                               "INNER JOIN prospectos p ON ep.contacto_id = p.id " &
                               "INNER JOIN listaempresas e ON p.cia = e.cia " &
                               "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
                               "WHERE ep.contacto_id = " & contactoId & " " &
                               "AND ep.enviado = False " &
                               "ORDER BY ep.fecha_programada"

            ' Si no encuentra por ID específico, buscar el PRIMER contacto programado
            If contactoId = 0 Then
                sql = "SELECT TOP 1 " &
                     "p.nombre, p.puesto, p.mail, " &
                     "e.cia, e.segmentoavital, " &
                     "ep.secuencia_id, ep.mensaje_id, ep.fecha_programada, " &
                     "m.asunto, m.cuerpo, ep.contacto_id " &
                     "FROM envios_programados ep " &
                     "INNER JOIN prospectos p ON ep.contacto_id = p.id " &
                     "INNER JOIN listaempresas e ON p.cia = e.cia " &
                     "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
                     "WHERE ep.enviado = False " &
                     "ORDER BY ep.fecha_programada"
            End If

            rs.Open(sql, ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If rs.EOF Then
                EscribirLog("⚠️ No hay correos planeados pendientes")
                rs.Close()
                Return
            End If

            ' 2. Extraer datos
            Dim nombre As String = rs.Fields("nombre").Value.ToString()
            Dim puesto As String = rs.Fields("puesto").Value.ToString()
            Dim emailReal As String = rs.Fields("mail").Value.ToString()
            Dim empresa As String = rs.Fields("cia").Value.ToString()
            Dim segmento As String = rs.Fields("segmentoavital").Value.ToString()
            Dim secuenciaId As Integer = rs.Fields("secuencia_id").Value
            Dim mensajeId As Integer = rs.Fields("mensaje_id").Value
            Dim fechaProgramada As DateTime = CDate(rs.Fields("fecha_programada").Value)
            Dim asuntoBase As String = rs.Fields("asunto").Value.ToString()
            Dim cuerpoBase As String = rs.Fields("cuerpo").Value.ToString()

            If contactoId = 0 Then contactoId = rs.Fields("contacto_id").Value
            rs.Close()

            ' 3. Personalizar mensaje usando la función de formateo
            Dim asuntoPersonalizado As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa)
            Dim cuerpoPersonalizado As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

            ' 4. Determinar día de la secuencia
            Dim diaSecuencia As Integer = 1
            If diaNumero > 0 Then
                diaSecuencia = diaNumero
            Else
                ' Intentar detectar el día basado en el mensaje_id
                Dim rsDia As New Recordset
                rsDia.Open("SELECT dia_numero FROM mensajes WHERE id = " & mensajeId, ba)
                If Not rsDia.EOF Then
                    diaSecuencia = rsDia.Fields("dia_numero").Value
                End If
                rsDia.Close()
            End If

            ' 5. Agregar encabezado de PRUEBA/REVISIÓN
            asuntoPersonalizado = "[REVISIÓN] " & asuntoPersonalizado & " - Día " & diaSecuencia

            ' 6. Agregar panel de información de diagnóstico
            Dim panelDiagnostico As String = "<div style='background-color: #f8f9fa; border-left: 4px solid #007bff; padding: 15px; margin: 20px 0; font-size: 13px;'>" &
                                            "<h4 style='margin-top: 0; color: #0056b3;'>🔍 INFORMACIÓN DE PRUEBA</h4>" &
                                            "<table style='width: 100%; border-collapse: collapse;'>" &
                                            "<tr><td style='padding: 5px;'><strong>Contacto Real:</strong></td><td style='padding: 5px;'>" & nombre & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Email Real:</strong></td><td style='padding: 5px;'>" & emailReal & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Empresa:</strong></td><td style='padding: 5px;'>" & empresa & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Puesto:</strong></td><td style='padding: 5px;'>" & puesto & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Segmento:</strong></td><td style='padding: 5px;'>" & segmento & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Secuencia ID:</strong></td><td style='padding: 5px;'>" & secuenciaId & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Día de secuencia:</strong></td><td style='padding: 5px;'>" & diaSecuencia & " de 5</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Programado para:</strong></td><td style='padding: 5px;'>" & fechaProgramada.ToString("dd/MM/yyyy HH:mm") & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>ID Mensaje:</strong></td><td style='padding: 5px;'>" & mensajeId & "</td></tr>" &
                                            "<tr><td style='padding: 5px;'><strong>Enviado a:</strong></td><td style='padding: 5px; color: #28a745;'><strong>Abraham (MODO REVISIÓN)</strong></td></tr>" &
                                            "</table></div>"

            ' 7. Insertar panel después del saludo
            Dim cuerpoFinal As String = cuerpoPersonalizado
            If cuerpoFinal.Contains("</p>") Then
                Dim primeraParte As String = cuerpoFinal.Substring(0, cuerpoFinal.IndexOf("</p>") + 4)
                Dim segundaParte As String = cuerpoFinal.Substring(cuerpoFinal.IndexOf("</p>") + 4)
                cuerpoFinal = primeraParte & panelDiagnostico & segundaParte
            Else
                cuerpoFinal = panelDiagnostico & cuerpoFinal
            End If

            ' 8. Agregar footer de prueba
            cuerpoFinal &= "<hr style='margin: 30px 0; border: 1px dashed #ccc;'/>" &
                          "<p style='text-align: center; color: #6c757d; font-size: 12px;'>" &
                          "⚠️ <strong>MODO REVISIÓN</strong> - Este correo fue enviado a Abraham para revisión. " &
                          "En producción iría a: <strong>" & emailReal & "</strong></p>"

            ' 9. Obtener cuenta de envío
            Dim rsCuenta As New Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital", ba, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                rsCuenta.Close()
                EscribirLog("❌ No hay cuentas de envío disponibles")
                Return
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            ' 10. Configurar SMTP2Go
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            ' 11. Crear y enviar correo a Abraham
            Dim mail As New Net.Mail.MailMessage
            mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
            mail.To.Add("abraham@avital.mx")  ' ← ENVÍO DE REVISIÓN A ABRAHAM
            mail.Subject = asuntoPersonalizado
            mail.Body = cuerpoFinal
            mail.IsBodyHtml = True

            smtpClient.Send(mail)

            ' 12. Log del envío
            EscribirLog("✅ Correo de REVISIÓN enviado a Abraham:")
            EscribirLog("   Contacto: " & nombre & " (" & emailReal & ")")
            EscribirLog("   Empresa: " & empresa)
            EscribirLog("   Asunto: " & asuntoPersonalizado)
            EscribirLog("   Día secuencia: " & diaSecuencia)
            EscribirLog("   Programado original para: " & fechaProgramada.ToString("dd/MM HH:mm"))

        Catch ex As Exception
            EscribirLog("❌ Error enviando correo de revisión: " & ex.Message)
        End Try




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        EscribirLog("=== REVISIÓN: PRIMER CORREO PLANEADO ===")
        EnviarCorreoPlaneadoParaRevision(0, 1) ' 0 = primer contacto, 1 = día 1
        MessageBox.Show("✅ Correo de revisión (Día 1) enviado a Abraham", "Revisión", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            CargarSecuenciaRestaurantes()

            MessageBox.Show("✅ Secuencias de RESTAURANTES cargadas correctamente.", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            EscribirLog("✅ Secuencias de RESTAURANTES (TI y CEO) cargadas en BD.")
        Catch ex As Exception
            MessageBox.Show("❌ Error cargando secuencias: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EscribirLog("❌ Error cargando secuencias: " & ex.Message)
        Finally
            If ba.State = 1 Then ba.Close()
        End Try

    End Sub


    Private Sub CargarSecuenciaRestaurantes()
        Dim segmentoId As Integer = 0
        Dim rsSeg As New ADODB.Recordset

        ' 1. Buscar ID de segmento "Restaurantes / Cadenas Alimenticias"
        Try
            rsSeg.Open("SELECT id FROM segmentacion WHERE segmento = 'Restaurantes / Cadenas Alimenticias'", ba,
                   ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsSeg.EOF Then
                EscribirLog("❌ Segmento 'Restaurantes / Cadenas Alimenticias' NO existe en tabla segmentacion.")
                rsSeg.Close()
                Exit Sub
            End If

            segmentoId = CInt(rsSeg.Fields("id").Value)
            rsSeg.Close()
        Catch ex As Exception
            EscribirLog("❌ Error obteniendo segmento RESTAURANTES: " & ex.Message)
            Exit Sub
        End Try

        ' =========================================================
        ' ========== SECUENCIA RESTAURANTES - TI - CCTV ===========
        ' =========================================================
        Dim secuenciaId_TI As Integer = 0
        Dim rsSec As New ADODB.Recordset

        ' Verificar si ya existe una secuencia TI para este segmento
        rsSec.Open("SELECT id FROM secuencias_maestra WHERE segmento_id = " & segmentoId & " AND tipo_audiencia = 'TI'", ba,
               ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsSec.EOF Then
            secuenciaId_TI = CInt(rsSec.Fields("id").Value)
            EscribirLog("ℹ️ Secuencia RESTAURANTES - TI ya existía. Se sobreescribirán los mensajes.")
            rsSec.Close()
        Else
            rsSec.Close()
            ' Crear nueva secuencia TI
            rsSec.Open("secuencias_maestra", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsSec.AddNew()
            rsSec.Fields("segmento_id").Value = segmentoId
            rsSec.Fields("tipo_audiencia").Value = "TI"
            rsSec.Fields("nombre").Value = "Restaurantes / Cadenas Alimenticias - TI CCTV"
            rsSec.Fields("activa").Value = True
            rsSec.Update()
            secuenciaId_TI = CInt(rsSec.Fields("id").Value)
            rsSec.Close()
            EscribirLog("✅ Secuencia RESTAURANTES - TI creada (ID " & secuenciaId_TI & ").")
        End If

        ' Borrar mensajes anteriores de esa secuencia para no duplicar
        Try
            ba.Execute("DELETE FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId_TI)
        Catch ex As Exception
            EscribirLog("⚠️ No se pudieron borrar mensajes anteriores TI: " & ex.Message)
        End Try

        ' Insertar mensajes TI
        Dim rsMsg As New ADODB.Recordset

        ' TI - Día 1 (Primer mail)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
        rsMsg.Fields("dia_numero").Value = 1
        rsMsg.Fields("asunto").Value = "El CCTV no avisa. El problema te estalla a ti."
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Si un cliente se accidenta, hay un robo o un reclamo laboral… y tus cámaras no grabaron, el problema recae en el área de TI.</p>" &
        "<p>Necesitas tranquilidad real: saber que todo está grabando siempre.</p>" &
        "<p>En Avital revisamos: grabación real, estado de las cámaras y del NVR, alertas y retención.</p>" &
        "<p>Sin costo y sin compromiso.</p>" &
        "<p>No vendemos cámaras. Vendemos tranquilidad.</p>" &
        "<p>¿Lo reviso esta semana?</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' TI - Día 2 (Seguimiento corto)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
        rsMsg.Fields("dia_numero").Value = 2
        rsMsg.Fields("asunto").Value = "Te escribo por tu CCTV"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Solo para confirmar si recibiste mi mensaje anterior.</p>" &
        "<p>El riesgo sigue presente: el CCTV no avisa cuando deja de grabar… y el problema cae en TI.</p>" &
        "<p>Puedo revisarlo sin costo esta semana.</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' TI - Día 3 (Recordatorio suave, dolor reptiliano)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
        rsMsg.Fields("dia_numero").Value = 3
        rsMsg.Fields("asunto").Value = "Las fallas del CCTV no esperan"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Te escribo porque muchos restaurantes nos contactan después de un incidente… cuando descubren que sus cámaras no grabaron.</p>" &
        "<p>Prefiero ayudarte antes de que eso pase.</p>" &
        "<p>¿Agendamos revisión sin costo?</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' TI - Día 4 (Más directo)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
        rsMsg.Fields("dia_numero").Value = 4
        rsMsg.Fields("asunto").Value = "Si hoy falla tu CCTV… te cae encima a ti"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Si hoy ocurre un robo, reclamo laboral o incidente, y tus cámaras no están grabando, TI queda expuesto.</p>" &
        "<p>Puedo decirte en 20 minutos si tu sistema está grabando correctamente.</p>" &
        "<p>Sin costo. Sin compromiso.</p>" &
        "<p>¿Te lo reviso hoy?</p>" &
        "<p>👉 https://wa.me/525542359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' TI - Día 5 (Último correo – cierre suave)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
        rsMsg.Fields("dia_numero").Value = 5
        rsMsg.Fields("asunto").Value = "Entiendo que ahora no es el momento de atenderme"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Entiendo que ahora no es el mejor momento para revisar el tema del CCTV.</p>" &
        "<p>Solo quiero dejarte mi contacto por si en algún momento quieres validar que tu sistema está realmente grabando bien.</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        EscribirLog("✅ Mensajes RESTAURANTES - TI creados (5 correos).")

        ' =========================================================
        ' ========== SECUENCIA RESTAURANTES - CEOs - CCTV =========
        ' =========================================================
        Dim secuenciaId_CEO As Integer = 0

        ' Verificar si ya existe una secuencia CEO para este segmento
        rsSec.Open("SELECT id FROM secuencias_maestra WHERE segmento_id = " & segmentoId & " AND tipo_audiencia = 'CEO'", ba,
               ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsSec.EOF Then
            secuenciaId_CEO = CInt(rsSec.Fields("id").Value)
            EscribirLog("ℹ️ Secuencia RESTAURANTES - CEO ya existía. Se sobreescribirán los mensajes.")
            rsSec.Close()
        Else
            rsSec.Close()
            ' Crear nueva secuencia CEO
            rsSec.Open("secuencias_maestra", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsSec.AddNew()
            rsSec.Fields("segmento_id").Value = segmentoId
            rsSec.Fields("tipo_audiencia").Value = "CEO"
            rsSec.Fields("nombre").Value = "Restaurantes / Cadenas Alimenticias - CEO CCTV"
            rsSec.Fields("activa").Value = True
            rsSec.Update()
            secuenciaId_CEO = CInt(rsSec.Fields("id").Value)
            rsSec.Close()
            EscribirLog("✅ Secuencia RESTAURANTES - CEO creada (ID " & secuenciaId_CEO & ").")
        End If

        ' Borrar mensajes anteriores de esa secuencia para no duplicar
        Try
            ba.Execute("DELETE FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId_CEO)
        Catch ex As Exception
            EscribirLog("⚠️ No se pudieron borrar mensajes anteriores CEO: " & ex.Message)
        End Try

        ' CEO - Día 1 (Primer mail)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
        rsMsg.Fields("dia_numero").Value = 1
        rsMsg.Fields("asunto").Value = "Si hoy pasa algo… ¿tienes el video?"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>La mayoría de los restaurantes descubre que sus cámaras NO grabaron justo después de un robo, un incidente o un reclamo.</p>" &
        "<p>Lo que realmente necesitas es tranquilidad: saber que tienes evidencia cuando la necesites.</p>" &
        "<p>En Avital auditamos tu CCTV: grabación efectiva, discos duros, ángulos críticos, fallas silenciosas.</p>" &
        "<p>Revisión sin costo.</p>" &
        "<p>En Avital no vendemos cámaras. Vendemos tranquilidad. ¿Quieres que lo revisemos?</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' CEO - Día 2 (Seguimiento inmediato)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
        rsMsg.Fields("dia_numero").Value = 2
        rsMsg.Fields("asunto").Value = "Te escribo sobre tus cámaras"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Solo para confirmar si recibiste mi mensaje.</p>" &
        "<p>La pregunta sigue siendo la misma: si hoy pasa algo… ¿tienes el video?</p>" &
        "<p>Puedo revisarlo sin costo.</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' CEO - Día 3 (Recordatorio más emocional)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
        rsMsg.Fields("dia_numero").Value = 3
        rsMsg.Fields("asunto").Value = "Cuando lo necesitas, el video debe existir"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Muchos dueños descubren que sus cámaras NO grabaron después de un incidente, un robo o un reclamo.</p>" &
        "<p>Ese momento no se puede retroceder.</p>" &
        "<p>Si quieres, reviso tu CCTV esta semana sin costo.</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' CEO - Día 4 (Seguimiento final directo)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
        rsMsg.Fields("dia_numero").Value = 4
        rsMsg.Fields("asunto").Value = "Si hoy pasa algo… ¿estás protegido?"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Si hoy ocurre un incidente y tu CCTV no grabó, no hay evidencia, no hay defensa y no hay tranquilidad.</p>" &
        "<p>Puedo decirte en minutos si tu sistema está realmente grabando.</p>" &
        "<p>Revisión sin costo.</p>" &
        "<p>¿Lo vemos hoy?</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        ' CEO - Día 5 (Cierre suave)
        rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        rsMsg.AddNew()
        rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
        rsMsg.Fields("dia_numero").Value = 5
        rsMsg.Fields("asunto").Value = "Si ahora no es el momento, lo entiendo"
        rsMsg.Fields("cuerpo").Value =
        "<p>Hola, {nombre}.</p>" &
        "<p>Si ahora no es buen momento para revisar el CCTV, lo entiendo totalmente.</p>" &
        "<p>Cuando quieras validar que tu sistema está grabando bien y que tienes evidencia cuando la necesites, puedes escribirme directo.</p>" &
        "<p>👉 https://wa.me/525542842359</p>"
        rsMsg.Update()
        rsMsg.Close()

        EscribirLog("✅ Mensajes RESTAURANTES - CEO creados (5 correos).")
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            ' Lista de segmentos (EXCEPTO Restaurantes, que ya cargaste aparte)
            Dim segmentos = {
                "Clínicas / Hospitales / Laboratorios",
                "Despachos / Oficinas / Corporativos",
                "Manufactura / Industria Pesada",
                "Hoteles / Hospitalidad",
                "Inmobiliario / Administración de Edificios",
                "Constructoras / Arquitectura / Ingeniería",
                "Retail / Tiendas / Franquicias",
                "Tecnología / Software / Startups",
                "Logística / Transporte / Almacenes",
                "Educación",
                "Energía / Medio Ambiente",
                "Medios / Agencias / Creativos",
                "Gobierno / Asociaciones Civiles"
            }

            For Each seg In segmentos
                CargarSecuenciaGenericaSegmento(seg)
            Next

            MessageBox.Show("✅ Secuencias genéricas cargadas para TODOS los segmentos (excepto Restaurantes).", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            EscribirLog("✅ Secuencias genéricas cargadas para todos los segmentos.")
        Catch ex As Exception
            MessageBox.Show("❌ Error cargando secuencias genéricas: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EscribirLog("❌ Error cargando secuencias genéricas: " & ex.Message)
        Finally
            If ba.State = 1 Then ba.Close()
        End Try

    End Sub

    Private Sub CargarSecuenciaGenericaSegmento(segmentoNombre As String)
        Dim segmentoId As Integer = 0
        Dim rsSeg As New ADODB.Recordset

        ' 1. Buscar ID del segmento
        Try
            rsSeg.Open("SELECT id FROM segmentacion WHERE segmento = '" & segmentoNombre.Replace("'", "''") & "'", ba,
                   ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsSeg.EOF Then
                EscribirLog("❌ Segmento no encontrado en tabla segmentacion: " & segmentoNombre)
                rsSeg.Close()
                Exit Sub
            End If

            segmentoId = CInt(rsSeg.Fields("id").Value)
            rsSeg.Close()
        Catch ex As Exception
            EscribirLog("❌ Error obteniendo segmento [" & segmentoNombre & "]: " & ex.Message)
            Exit Sub
        End Try

        ' =========================================================
        ' ========== CREAR / OBTENER SECUENCIA TI =================
        ' =========================================================
        Dim secuenciaId_TI As Integer = 0
        Dim rsSec As New ADODB.Recordset

        ' Buscar si ya existe secuencia TI
        rsSec.Open("SELECT id FROM secuencias_maestra WHERE segmento_id = " & segmentoId & " AND tipo_audiencia = 'TI'", ba,
               ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsSec.EOF Then
            secuenciaId_TI = CInt(rsSec.Fields("id").Value)
            EscribirLog("ℹ️ Secuencia TI ya existía para segmento: " & segmentoNombre & ". Se sobreescriben mensajes.")
            rsSec.Close()
        Else
            rsSec.Close()
            ' Crear nueva secuencia TI
            rsSec.Open("secuencias_maestra", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsSec.AddNew()
            rsSec.Fields("segmento_id").Value = segmentoId
            rsSec.Fields("tipo_audiencia").Value = "TI"
            rsSec.Fields("nombre").Value = segmentoNombre & " - TI CCTV"
            rsSec.Fields("activa").Value = True
            rsSec.Update()
            secuenciaId_TI = CInt(rsSec.Fields("id").Value)
            rsSec.Close()
            EscribirLog("✅ Secuencia TI creada para segmento: " & segmentoNombre & " (ID " & secuenciaId_TI & ").")
        End If

        ' Borrar mensajes previos TI
        Try
            ba.Execute("DELETE FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId_TI)
        Catch ex As Exception
            EscribirLog("⚠️ No se pudieron borrar mensajes TI previos para " & segmentoNombre & ": " & ex.Message)
        End Try

        ' Insertar 5 mensajes TI
        Dim rsMsg As New ADODB.Recordset
        Dim asunto As String
        Dim cuerpo As String

        For dia As Integer = 1 To 5
            asunto = ""
            cuerpo = ""

            Select Case dia
                Case 1
                    asunto = "El CCTV no avisa. El problema te estalla a ti."
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Si hoy hay un incidente, un robo o un reclamo laboral en {empresa} y las cámaras no grabaron, el problema recae directo en TI.</p>" &
                    "<p>El CCTV no avisa cuando deja de grabar.</p>" &
                    "<p>En Avital revisamos: grabación real, estado de cámaras y NVR, discos, alertas y retención.</p>" &
                    "<p>Revisión sin costo y sin compromiso.</p>" &
                    "<p>¿Te ayudo a validar que todo está grabando bien?</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 2
                    asunto = "Te escribo por el CCTV de {empresa}"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Solo para confirmar si recibiste mi correo anterior sobre el CCTV de {empresa}.</p>" &
                    "<p>El riesgo es silencioso: el sistema puede dejar de grabar y nadie se entera hasta que ya es tarde.</p>" &
                    "<p>Puedo revisarlo esta semana sin costo.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 3
                    asunto = "Las fallas del CCTV no esperan"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Muchas empresas nos buscan después de un problema: descubren que sus cámaras no estaban grabando justo cuando más lo necesitaban.</p>" &
                    "<p>Prefiero ayudarte antes de que eso pase en {empresa}.</p>" &
                    "<p>¿Agendamos una revisión rápida sin costo?</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 4
                    asunto = "Si hoy falla el CCTV… te cae encima a ti"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Si hoy pasa algo serio en {empresa} y el CCTV no grabó, TI queda expuesto.</p>" &
                    "<p>En menos de 20 minutos puedo decirte si tu sistema está grabando correctamente.</p>" &
                    "<p>Sin costo, sin compromiso, solo claridad.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 5
                    asunto = "Entiendo que ahora no es el momento, pero el riesgo sigue"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Entiendo perfecto que tienes mil temas encima y ahora quizá no sea el mejor momento para revisar el CCTV.</p>" &
                    "<p>Solo quiero dejarte mi contacto por si en algún momento quieres validar que en {empresa} sí se está grabando todo lo importante.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"
            End Select

            rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsMsg.AddNew()
            rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_TI
            rsMsg.Fields("dia_numero").Value = dia
            rsMsg.Fields("asunto").Value = asunto
            rsMsg.Fields("cuerpo").Value = cuerpo
            rsMsg.Update()
            rsMsg.Close()
        Next

        EscribirLog("✅ Mensajes TI creados (5 correos) para segmento: " & segmentoNombre)

        ' =========================================================
        ' ========== CREAR / OBTENER SECUENCIA CEO ================
        ' =========================================================
        Dim secuenciaId_CEO As Integer = 0

        ' Buscar si ya existe secuencia CEO
        rsSec.Open("SELECT id FROM secuencias_maestra WHERE segmento_id = " & segmentoId & " AND tipo_audiencia = 'CEO'", ba,
               ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsSec.EOF Then
            secuenciaId_CEO = CInt(rsSec.Fields("id").Value)
            EscribirLog("ℹ️ Secuencia CEO ya existía para segmento: " & segmentoNombre & ". Se sobreescriben mensajes.")
            rsSec.Close()
        Else
            rsSec.Close()
            ' Crear nueva secuencia CEO
            rsSec.Open("secuencias_maestra", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsSec.AddNew()
            rsSec.Fields("segmento_id").Value = segmentoId
            rsSec.Fields("tipo_audiencia").Value = "CEO"
            rsSec.Fields("nombre").Value = segmentoNombre & " - CEO CCTV"
            rsSec.Fields("activa").Value = True
            rsSec.Update()
            secuenciaId_CEO = CInt(rsSec.Fields("id").Value)
            rsSec.Close()
            EscribirLog("✅ Secuencia CEO creada para segmento: " & segmentoNombre & " (ID " & secuenciaId_CEO & ").")
        End If

        ' Borrar mensajes previos CEO
        Try
            ba.Execute("DELETE FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId_CEO)
        Catch ex As Exception
            EscribirLog("⚠️ No se pudieron borrar mensajes CEO previos para " & segmentoNombre & ": " & ex.Message)
        End Try

        ' Insertar 5 mensajes CEO
        For dia As Integer = 1 To 5
            asunto = ""
            cuerpo = ""

            Select Case dia
                Case 1
                    asunto = "Si hoy pasa algo en {empresa}… ¿tienes el video?"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>La mayoría de las empresas descubre que sus cámaras NO grabaron después de un robo, un incidente o un reclamo.</p>" &
                    "<p>Ahí ya es demasiado tarde.</p>" &
                    "<p>Lo que realmente buscas como dirección es tranquilidad: saber que tendrás evidencia cuando la necesites.</p>" &
                    "<p>En Avital auditamos tu CCTV: grabación efectiva, discos, ángulos críticos y fallas silenciosas.</p>" &
                    "<p>Revisión sin costo.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 2
                    asunto = "Te escribo sobre las cámaras de {empresa}"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Solo para confirmar si viste mi correo anterior.</p>" &
                    "<p>Si hoy pasa algo en {empresa} y las cámaras no grabaron, no hay defensa ni respaldo.</p>" &
                    "<p>Puedo ayudarte a validar que el sistema está funcionando como debería.</p>" &
                    "<p>Revisión sin costo.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 3
                    asunto = "Cuando lo necesitas, el video debe existir"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Muchos directivos nos llaman después de un problema serio, cuando descubren que el video simplemente no existe.</p>" &
                    "<p>Ese momento no se puede regresar.</p>" &
                    "<p>Por eso proponemos revisar el CCTV antes, no después.</p>" &
                    "<p>¿Quieres que revisemos cómo está hoy en {empresa}?</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 4
                    asunto = "Tranquilidad es saber que sí hay evidencia"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>El CCTV no es solo un gasto: es la tranquilidad de que, si algo pasa en {empresa}, tendrás con qué respaldarte.</p>" &
                    "<p>En pocos minutos puedo decirte si tu sistema está realmente grabando.</p>" &
                    "<p>Revisión sin costo, sin compromiso.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"

                Case 5
                    asunto = "Si ahora no es el momento, lo entiendo"
                    cuerpo =
                    "<p>Hola, {nombre}.</p>" &
                    "<p>Si ahora no es buen momento, lo entiendo totalmente.</p>" &
                    "<p>Cuando quieras validar que las cámaras de {empresa} están grabando bien y que hay evidencia cuando la necesites, puedes escribirme directo.</p>" &
                    "<p>👉 https://wa.me/525542842359</p>"
            End Select

            rsMsg.Open("mensajes", ba, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            rsMsg.AddNew()
            rsMsg.Fields("secuencia_maestra_id").Value = secuenciaId_CEO
            rsMsg.Fields("dia_numero").Value = dia
            rsMsg.Fields("asunto").Value = asunto
            rsMsg.Fields("cuerpo").Value = cuerpo
            rsMsg.Update()
            rsMsg.Close()
        Next

        EscribirLog("✅ Mensajes CEO creados (5 correos) para segmento: " & segmentoNombre)
    End Sub

    Private Sub btnGenerarAgendaHoy_Click(sender As Object, e As EventArgs) Handles btnGenerarAgendaHoy.Click
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            GenerarAgendaHoy_NuevasEmpresas()

            MessageBox.Show("✅ Agenda de HOY generada. Revisa la tabla envios_programados.", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("❌ Error generando agenda: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EscribirLog("❌ Error en btnGenerarAgendaHoy: " & ex.Message)
        Finally
            If ba.State = 1 Then ba.Close()
        End Try

    End Sub


    Private Sub btnProbarEnviosHoy_Click(sender As Object, e As EventArgs) Handles btnProbarEnviosHoy.Click
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            EnviarAgendaDeHoy_ModoPrueba()

            MessageBox.Show("✅ Se intentó enviar la agenda de hoy en modo PRUEBA a abraham@avital.mx", "Prueba", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("❌ Error enviando agenda de hoy (prueba): " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EscribirLog("❌ Error en btnProbarEnviosHoy: " & ex.Message)
        Finally
            If ba.State = 1 Then ba.Close()
        End Try
    End Sub

    Private Sub EnviarAgendaDeHoy_ModoPrueba()
        EscribirLog("=== ENVIAR AGENDA DE HOY - MODO PRUEBA ===")

        Dim hoyInicio As Date = Date.Today
        Dim hoyFin As Date = Date.Today.AddDays(1)

        Try
            Dim rs As New ADODB.Recordset

            ' ⚠️ CORRECCIÓN: SQL corregido - faltaba "FROM" y tenía JOIN incorrecto
            Dim sql As String =
        "SELECT ep.id, ep.contacto_id, ep.secuencia_id, ep.mensaje_id, ep.fecha_programada, " &
        "p.nombrecompleto, p.puesto, p.mail, e.cia, e.segmentoavital, m.asunto, m.cuerpo " &
        "FROM ((envios_programados ep " &
        "INNER JOIN prospectos p ON ep.contacto_id = p.id) " &
        "INNER JOIN listaempresas e ON p.cia = e.cia) " &
        "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
        "WHERE ep.enviado = False " &
        "AND ep.fecha_programada >= #" & hoyInicio.ToString("yyyy-MM-dd 00:00:00") & "# " &
        "AND ep.fecha_programada < #" & hoyFin.ToString("yyyy-MM-dd 00:00:00") & "# " &
        "ORDER BY ep.fecha_programada"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.EOF Then
                EscribirLog("ℹ️ No hay envíos programados para hoy.")
                rs.Close()
                Return
            End If

            ' Obtener una cuenta de mailsavital para enviar
            Dim rsCuenta As New ADODB.Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital WHERE limite > 0", ba,
                      ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                EscribirLog("❌ No hay cuentas disponibles en mailsavital.")
                rsCuenta.Close()
                rs.Close()
                Return
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            ' Configurar SMTP2Go
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            Dim enviados As Integer = 0
            Dim total As Integer = 0

            rs.MoveFirst()
            While Not rs.EOF
                total += 1

                ' Solo enviar los PRIMEROS 5 para prueba
                If enviados >= 2 Then
                    EscribirLog("⏸️ Se limitó a 5 correos de prueba")
                    Exit While
                End If

                Dim nombre As String = rs.Fields("nombrecompleto").Value.ToString()
                Dim puesto As String = rs.Fields("puesto").Value.ToString()
                Dim mailReal As String = rs.Fields("mail").Value.ToString()
                Dim empresa As String = rs.Fields("cia").Value.ToString()
                Dim segmento As String = rs.Fields("segmentoavital").Value.ToString()
                Dim asuntoBase As String = rs.Fields("asunto").Value.ToString()
                Dim cuerpoBase As String = rs.Fields("cuerpo").Value.ToString()

                ' Personalizar
                'Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa).Replace("{puesto}", puesto)
                'Dim cuerpo As String = cuerpoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa).Replace("{puesto}", puesto)
                ' Personalizar usando la función de formateo
                Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa).Replace("{puesto}", puesto)
                Dim cuerpo As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

                ' Prefijo modo prueba
                asunto = "[PRUEBA] " & asunto

                ' Panel de diagnóstico
                cuerpo = "<div style='background:#f0f8ff; padding:10px; border-left:4px solid #007bff; margin-bottom:15px;'>" &
                     "<strong>⚠️ MODO PRUEBA</strong><br/>" &
                     "Este correo estaría dirigido a: <strong>" & nombre & "</strong><br/>" &
                     "Email real: " & mailReal & "<br/>" &
                     "Empresa: " & empresa & "<br/>" &
                     "Puesto: " & puesto & "<br/>" &
                     "Segmento: " & segmento & "</div>" &
                     cuerpo

                Dim mail As New Net.Mail.MailMessage
                mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
                mail.To.Add("abraham@avital.mx")   ' SIEMPRE a ti
                mail.Subject = asunto
                mail.Body = cuerpo
                mail.IsBodyHtml = True

                Try
                    smtpClient.Send(mail)
                    enviados += 1
                    EscribirLog("✅ [" & enviados & "] Enviado a Abraham: " & asunto.Substring(0, Math.Min(50, asunto.Length)) & "...")
                Catch ex As Exception
                    EscribirLog("❌ Error enviando correo " & total & ": " & ex.Message)
                End Try

                rs.MoveNext()
            End While

            rs.Close()

            EscribirLog("=== FIN MODO PRUEBA ===")
            EscribirLog("   Correos encontrados para hoy: " & total)
            EscribirLog("   Correos enviados a Abraham: " & enviados & " de 5")

            If enviados = 0 Then
                MessageBox.Show("No se pudo enviar ningún correo. Revisa: 1) Cuenta SMTP, 2) Credenciales, 3) Conexión internet",
                          "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                MessageBox.Show("✅ Se enviaron " & enviados & " correos de prueba a abraham@avital.mx",
                          "Prueba Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            EscribirLog("❌ Error en EnviarAgendaDeHoy_ModoPrueba: " & ex.Message)
            MessageBox.Show("❌ Error SQL/SMTP: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub EnviarAgendaDeHoy_MODO_REAL()
        EscribirLog("=== ENVIAR AGENDA DE HOY - MODO REAL ===")

        Dim hoyInicio As Date = Date.Today
        Dim hoyFin As Date = Date.Today.AddDays(1)

        Try
            Dim rs As New ADODB.Recordset
            Dim sql As String =
        "SELECT ep.id, ep.contacto_id, ep.secuencia_id, ep.mensaje_id, ep.fecha_programada, " &
        "p.nombrecompleto, p.puesto, p.mail, e.cia, e.segmentoavital, m.asunto, m.cuerpo " &
        "FROM ((envios_programados ep " &
        "INNER JOIN prospectos p ON ep.contacto_id = p.id) " &
        "INNER JOIN listaempresas e ON p.cia = e.cia) " &
        "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
        "WHERE ep.enviado = False " &
        "AND ep.fecha_programada >= #" & hoyInicio.ToString("yyyy-MM-dd 00:00:00") & "# " &
        "AND ep.fecha_programada < #" & hoyFin.ToString("yyyy-MM-dd 00:00:00") & "# " &
        "ORDER BY ep.fecha_programada"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.EOF Then
                EscribirLog("ℹ️ No hay envíos programados para hoy.")
                rs.Close()
                Return
            End If

            ' Obtener cuenta de envío
            Dim rsCuenta As New ADODB.Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital WHERE limite > 0", ba,
                  ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                EscribirLog("❌ No hay cuentas disponibles en mailsavital.")
                rsCuenta.Close()
                rs.Close()
                Return
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            ' Configurar SMTP2Go
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            Dim enviados As Integer = 0
            Dim errores As Integer = 0
            Dim total As Integer = 0

            rs.MoveFirst()
            While Not rs.EOF
                total += 1

                Dim nombre As String = rs.Fields("nombrecompleto").Value.ToString()
                Dim puesto As String = rs.Fields("puesto").Value.ToString()
                Dim mailReal As String = rs.Fields("mail").Value.ToString()
                Dim empresa As String = rs.Fields("cia").Value.ToString()
                Dim segmento As String = rs.Fields("segmentoavital").Value.ToString()
                Dim asuntoBase As String = rs.Fields("asunto").Value.ToString()
                Dim cuerpoBase As String = rs.Fields("cuerpo").Value.ToString()
                Dim contactoId As Integer = CInt(rs.Fields("contacto_id").Value)
                Dim mensajeId As Integer = CInt(rs.Fields("mensaje_id").Value)

                Try
                    ' Personalizar usando la función de formateo
                    Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa).Replace("{puesto}", puesto)
                    Dim cuerpo As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

                    ' NO agregar prefijo [PRUEBA] ni panel diagnóstico

                    Dim mail As New Net.Mail.MailMessage
                    mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
                    mail.To.Add(mailReal)   ' ← CORREO REAL del contacto
                    mail.Subject = asunto
                    mail.Body = cuerpo
                    mail.IsBodyHtml = True

                    smtpClient.Send(mail)
                    enviados += 1

                    ' Marcar como enviado en la BD
                    MarcarCorreoEnviado(contactoId, mensajeId)

                    EscribirLog("✅ [" & enviados & "] Enviado REAL a " & nombre & " (" & mailReal & ")")

                Catch ex As Exception
                    errores += 1
                    EscribirLog("❌ Error enviando a " & mailReal & ": " & ex.Message)
                End Try

                rs.MoveNext()
            End While

            rs.Close()

            EscribirLog("=== FIN ENVÍOS REALES ===")
            EscribirLog("   Correos encontrados para hoy: " & total)
            EscribirLog("   Correos enviados REALES: " & enviados)
            EscribirLog("   Errores: " & errores)

            MessageBox.Show("✅ Se enviaron " & enviados & " correos REALES." & vbCrLf &
                      "Errores: " & errores, "Envíos Completados", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            EscribirLog("❌ Error en EnviarAgendaDeHoy_MODO_REAL: " & ex.Message)
            MessageBox.Show("❌ Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub MarcarCorreoEnviado(contactoId As Integer, mensajeId As Integer)
        Try
            ' Marcar en envios_programados
            Dim rs As New ADODB.Recordset
            rs.Open("SELECT * FROM envios_programados WHERE contacto_id = " & contactoId & " AND mensaje_id = " & mensajeId, ba,
                ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            If Not rs.EOF Then
                rs.Fields("enviado").Value = True
                rs.Fields("fecha_envio").Value = DateTime.Now
                rs.Update()
            End If

            rs.Close()

        Catch ex As Exception
            EscribirLog("⚠️ Error marcando correo como enviado: " & ex.Message)
        End Try
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            ' 1) Cargar segmentacion_personas en diccionarios
            Dim dictTipo As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim dictRol As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim rsSeg As New ADODB.Recordset
            rsSeg.Open(
            "SELECT cargo, tipoaudiencia, rol_nivel " &
            "FROM segmentacion_personas " &
            "WHERE cargo IS NOT NULL",
            ba,
            ADODB.CursorTypeEnum.adOpenStatic,
            ADODB.LockTypeEnum.adLockReadOnly
        )

            While Not rsSeg.EOF
                Dim cargo As String = rsSeg.Fields("cargo").Value.ToString().Trim()

                Dim ta As String = ""
                If Not IsDBNull(rsSeg.Fields("tipoaudiencia").Value) Then
                    ta = rsSeg.Fields("tipoaudiencia").Value.ToString().Trim()
                End If

                Dim rol As String = ""
                If Not IsDBNull(rsSeg.Fields("rol_nivel").Value) Then
                    rol = rsSeg.Fields("rol_nivel").Value.ToString().Trim()
                End If

                If cargo <> "" Then
                    If Not dictTipo.ContainsKey(cargo) Then
                        dictTipo.Add(cargo, ta)
                        dictRol.Add(cargo, rol)
                    End If
                End If

                rsSeg.MoveNext()
            End While
            rsSeg.Close()

            EscribirLog("Cargos cargados desde segmentacion_personas: " & dictTipo.Count)

            ' 2) Recorrer SOLO los prospectos que NO tienen tipoaudiencia o rol_nivel
            Dim rsPros As New ADODB.Recordset

            Dim sqlPros As String =
            "SELECT * FROM prospectos " &
            "WHERE (tipoaudiencia IS NULL OR tipoaudiencia = '') " &
            "   OR (rol_nivel IS NULL OR rol_nivel = '')"

            rsPros.Open(
            sqlPros,
            ba,
            ADODB.CursorTypeEnum.adOpenKeyset,
            ADODB.LockTypeEnum.adLockOptimistic
        )

            Dim total As Integer = 0
            Dim clasificados As Integer = 0
            Dim sinMatch As Integer = 0

            While Not rsPros.EOF
                total += 1

                Dim puesto As String = ""
                If Not IsDBNull(rsPros.Fields("puesto").Value) Then
                    puesto = rsPros.Fields("puesto").Value.ToString().Trim()
                End If

                If puesto <> "" AndAlso dictTipo.ContainsKey(puesto) Then
                    rsPros.Fields("tipoaudiencia").Value = dictTipo(puesto)
                    rsPros.Fields("rol_nivel").Value = dictRol(puesto)
                    clasificados += 1
                Else
                    ' No encontrado en segmentacion_personas → lo dejamos en NULL
                    rsPros.Fields("tipoaudiencia").Value = DBNull.Value
                    rsPros.Fields("rol_nivel").Value = DBNull.Value
                    sinMatch += 1
                End If

                rsPros.Update()
                rsPros.MoveNext()
            End While

            rsPros.Close()

            EscribirLog("✅ Clasificación completada.")
            EscribirLog("   Prospectos procesados (sin tipo/rol): " & total)
            EscribirLog("   Clasificados (con match en cargo): " & clasificados)
            EscribirLog("   Sin match en segmentacion_personas: " & sinMatch)

            MessageBox.Show(
            "✅ Clasificación completada." & vbCrLf &
            "Procesados (sin tipo/rol): " & total & vbCrLf &
            "Clasificados: " & clasificados & vbCrLf &
            "Sin match: " & sinMatch & vbCrLf &
            "Revisa en Access si necesitas agregar más cargos a segmentacion_personas.",
            "Listo",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

        Catch ex As Exception
            EscribirLog("❌ Error en clasificación: " & ex.Message)
            MessageBox.Show("❌ Error clasificando prospectos: " & ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
        Finally
            If ba.State = 1 Then ba.Close()
        End Try
    End Sub


    ' Programa SOLO el día 1 de la secuencia para un contacto
    Private Sub ProgramarPrimerCorreo(contacto As ContactoAgenda, secuenciaId As Integer)
        Try
            ' 1) Obtener mensaje del día 1
            Dim rsMsg As New ADODB.Recordset
            rsMsg.Open("SELECT id FROM mensajes WHERE secuencia_maestra_id = " & secuenciaId & " AND dia_numero = 1",
               ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsMsg.EOF Then
                EscribirLog("   ⚠️ No hay mensaje Día 1 para secuencia " & secuenciaId)
                rsMsg.Close()
                Exit Sub
            End If

            Dim mensajeId As Integer = CInt(rsMsg.Fields("id").Value)
            rsMsg.Close()

            ' 2) Calcular fecha base (hoy, ajustado a Mar-Jue)
            Dim hoy As DateTime = DateTime.Today
            Dim fechaBase As DateTime = hoy

            ' Ajustar a martes-jueves
            While fechaBase.DayOfWeek < DayOfWeek.Tuesday OrElse fechaBase.DayOfWeek > DayOfWeek.Thursday
                fechaBase = fechaBase.AddDays(1)
            End While

            ' 3) Obtener hora base según rol
            Dim horaBase As String = ObtenerHoraPorRol(contacto.RolNivel)

            ' 4) AGREGAR RANGO ALEATORIO de 60 a 180 segundos (1-3 minutos)
            Dim rand As New Random()
            Dim segundosAleatorios As Integer = rand.Next(60, 181) ' Entre 1 y 3 minutos

            ' Crear DateTime con hora base + segundos aleatorios
            Dim horaCompleta As DateTime = DateTime.Parse(fechaBase.ToString("yyyy-MM-dd") & " " & horaBase)
            horaCompleta = horaCompleta.AddSeconds(segundosAleatorios)

            ' 5) Si estamos en horario de trabajo y es retrasado, enviar más pronto
            Dim ahora As DateTime = DateTime.Now
            If ahora >= DateTime.Parse(fechaBase.ToString("yyyy-MM-dd") & " " & hora_inicio_envios) AndAlso
           ahora <= DateTime.Parse(fechaBase.ToString("yyyy-MM-dd") & " " & hora_fin_envios) Then

                ' Si la hora programada YA PASÓ, enviar pronto (pero con retraso aleatorio)
                If horaCompleta < ahora Then
                    ' Enviar entre 30 y 120 segundos desde ahora
                    segundosAleatorios = rand.Next(30, 121)
                    horaCompleta = ahora.AddSeconds(segundosAleatorios)
                    EscribirLog("   ⏰ Correo retrasado - Programado para " & horaCompleta.ToString("HH:mm:ss"))
                End If
            End If

            Dim fechaProgramada As DateTime = horaCompleta

            ' 6) Insertar en envios_programados
            Dim rsEnvio As New ADODB.Recordset
            rsEnvio.Open("envios_programados", ba,
                 ADODB.CursorTypeEnum.adOpenKeyset,
                 ADODB.LockTypeEnum.adLockOptimistic)

            rsEnvio.AddNew()
            rsEnvio.Fields("contacto_id").Value = contacto.Id
            rsEnvio.Fields("secuencia_id").Value = secuenciaId
            rsEnvio.Fields("mensaje_id").Value = mensajeId
            rsEnvio.Fields("fecha_programada").Value = fechaProgramada
            rsEnvio.Fields("enviado").Value = False
            rsEnvio.Update()
            rsEnvio.Close()

            ' 7) Marcar contacto en secuencia (día 1)
            MarcarContactoEnSecuencia(contacto.Id, secuenciaId)

            EscribirLog("   ✅ Programado Día 1 para " & contacto.Nombre & " (" & contacto.Mail &
                ") | Rol=" & contacto.RolNivel & " | Hora=" & fechaProgramada.ToString("HH:mm:ss") &
                " (+" & segundosAleatorios & " seg)")

        Catch ex As Exception
            EscribirLog("   ❌ Error en ProgramarPrimerCorreo: " & ex.Message)
        End Try
    End Sub

    Private Sub GenerarAgendaHoy_NuevasEmpresas()
        Try
            If ba.State <> 1 Then
                ba.Open(rutabd)
            End If

            EscribirLog("=== GENERAR AGENDA DE HOY (NUEVAS EMPRESAS) ===")

            ' 1) Contar envíos YA programados para hoy
            Dim hoy As DateTime = DateTime.Today
            Dim manana As DateTime = hoy.AddDays(1)

            Dim rsCount As New ADODB.Recordset
            Dim sqlCount As String =
            "SELECT COUNT(*) AS total " &
            "FROM envios_programados " &
            "WHERE fecha_programada >= #" & hoy.ToString("yyyy-MM-dd") & " 00:00:00# " &
            "AND fecha_programada < #" & manana.ToString("yyyy-MM-dd") & " 00:00:00#"

            rsCount.Open(sqlCount, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim yaProgramadosHoy As Integer = 0
            If Not rsCount.EOF Then
                yaProgramadosHoy = CInt(rsCount.Fields("total").Value)
            End If
            rsCount.Close()

            EscribirLog("📊 Envíos YA programados para hoy: " & yaProgramadosHoy)

            Dim capacidadDisponible As Integer = limite_diario_envios - yaProgramadosHoy
            If capacidadDisponible <= 0 Then
                EscribirLog("⏸️ Sin capacidad disponible para hoy (límite " & limite_diario_envios & ")")
                Exit Sub
            End If

            EscribirLog("📏 Espacio disponible para HOY: " & capacidadDisponible & " correos")

            ' 2) Buscar empresas elegibles (mínimo 3 contactos válidos)
            Dim sqlEmpresas As String =
            "SELECT e.cia, e.empleados, COUNT(p.id) AS contactos_validos " &
            "FROM listaempresas e " &
            "INNER JOIN prospectos p ON e.cia = p.cia " &
            "WHERE e.empleados < 900 " &
            "AND e.megalopolis = True " &
            "AND e.enterprise = False " &
            "AND e.conerror = False " &
            "AND e.con_contactos_en_secuencia = False " &
            "AND p.en_secuencia = False " &
            "AND p.mail IS NOT NULL " &
            "AND p.mailincorrecto = False " &
            "AND p.tipoaudiencia IS NOT NULL " &
            "GROUP BY e.cia, e.empleados " &
            "HAVING COUNT(p.id) >= 3 " &
            "ORDER BY e.empleados DESC"

            Dim rsEmpresas As New ADODB.Recordset
            rsEmpresas.Open(sqlEmpresas, ba,
                        ADODB.CursorTypeEnum.adOpenStatic,
                        ADODB.LockTypeEnum.adLockReadOnly)

            Dim restantes As Integer = capacidadDisponible
            Dim empresasTomadas As Integer = 0

            If rsEmpresas.EOF Then
                EscribirLog("ℹ️ No hay empresas nuevas que cumplan criterios base (mínimo 3 contactos, etc.)")
                rsEmpresas.Close()
                Exit Sub
            End If

            rsEmpresas.MoveFirst()

            While Not rsEmpresas.EOF AndAlso restantes > 0
                Dim cia As String = rsEmpresas.Fields("cia").Value.ToString()
                Dim contactosValidos As Integer = CInt(rsEmpresas.Fields("contactos_validos").Value)

                EscribirLog("🏢 Evaluando empresa: " & cia & " | Contactos válidos: " & contactosValidos)

                ' Si la empresa tiene más contactos de los que caben hoy, la dejamos para otro día
                If contactosValidos > restantes Then
                    EscribirLog("   ⏭️ Omitida hoy: contactos " & contactosValidos &
                            " > espacio disponible " & restantes)
                    rsEmpresas.MoveNext()
                    Continue While
                End If

                ' 3) Cargar contactos de la empresa con tipoaudiencia y rol_nivel
                Dim rsCont As New ADODB.Recordset
                Dim sqlCont As String =
                "SELECT p.id, p.nombrecompleto, p.puesto, p.mail, p.tipoaudiencia, p.rol_nivel, e.segmentoavital " &
                "FROM prospectos p " &
                "INNER JOIN listaempresas e ON p.cia = e.cia " &
                "WHERE p.cia = '" & cia.Replace("'", "''") & "' " &
                "AND p.en_secuencia = False " &
                "AND p.mail IS NOT NULL " &
                "AND p.mailincorrecto = False " &
                "AND p.tipoaudiencia IS NOT NULL"

                rsCont.Open(sqlCont, ba,
                        ADODB.CursorTypeEnum.adOpenStatic,
                        ADODB.LockTypeEnum.adLockReadOnly)

                Dim listaContactos As New List(Of ContactoAgenda)
                Dim numTI As Integer = 0
                Dim numNoTI As Integer = 0

                While Not rsCont.EOF
                    Dim c As New ContactoAgenda()
                    c.Id = CInt(rsCont.Fields("id").Value)
                    c.Nombre = rsCont.Fields("nombrecompleto").Value.ToString()
                    c.Puesto = rsCont.Fields("puesto").Value.ToString()
                    c.Mail = rsCont.Fields("mail").Value.ToString()
                    c.TipoAudiencia = rsCont.Fields("tipoaudiencia").Value.ToString()
                    c.Segmento = rsCont.Fields("segmentoavital").Value.ToString()

                    If Not IsDBNull(rsCont.Fields("rol_nivel").Value) Then
                        c.RolNivel = rsCont.Fields("rol_nivel").Value.ToString()
                    Else
                        c.RolNivel = "Staff"
                    End If

                    listaContactos.Add(c)

                    If c.TipoAudiencia = "TI" Then
                        numTI += 1
                    ElseIf c.TipoAudiencia = "No-TI" Then
                        numNoTI += 1
                    End If

                    rsCont.MoveNext()
                End While
                rsCont.Close()

                ' Validar que sea empresa mixta TI + No-TI
                If numTI = 0 OrElse numNoTI = 0 Then
                    EscribirLog("   ⏭️ Omitida: no es empresa mixta (TI=" & numTI & ", No-TI=" & numNoTI & ")")
                    rsEmpresas.MoveNext()
                    Continue While
                End If

                ' Validar de nuevo que no rebase capacidad por seguridad
                If listaContactos.Count > restantes Then
                    EscribirLog("   ⏭️ Omitida hoy (después de filtrar): contactos " &
                            listaContactos.Count & " > espacio disponible " & restantes)
                    rsEmpresas.MoveNext()
                    Continue While
                End If

                ' 4) Programar Día 1 para TODOS los contactos de la empresa
                Dim contactosProgramados As Integer = 0

                For Each c In listaContactos
                    ' Log de diagnóstico por contacto
                    EscribirLog("   → Contacto: " & c.Nombre &
                " | Puesto=" & c.Puesto &
                " | TipoAudiencia=" & c.TipoAudiencia &
                " | Rol=" & c.RolNivel &
                " | Mail=" & c.Mail)

                    ' Si ya no hay espacio, salimos
                    If restantes <= 0 Then
                        EscribirLog("     ⏸️ Sin espacio restante, ya no se programan más contactos de esta empresa.")
                        Exit For
                    End If

                    Dim secuenciaId As Integer = ObtenerSecuenciaId(c.Segmento, c.TipoAudiencia)
                    EscribirLog("     Secuencia encontrada ID=" & secuenciaId & " para segmento=" &
                c.Segmento & " / " & c.TipoAudiencia)

                    If secuenciaId <= 0 Then
                        EscribirLog("     ⚠️ Sin secuencia ACTIVA para este contacto. No se programa.")
                        Continue For
                    End If

                    ' Aquí sí se programa
                    ProgramarPrimerCorreo(c, secuenciaId)
                    restantes -= 1
                    contactosProgramados += 1

                    EscribirLog("     ✅ Programado Día 1. Restantes hoy: " & restantes)
                Next


                ' 5) Si se programaron todos los contactos, marcar empresa activada
                If contactosProgramados = listaContactos.Count AndAlso listaContactos.Count > 0 Then
                    MarcarEmpresaActivada(cia)
                    empresasTomadas += 1
                    EscribirLog("   ✅ Empresa ACTIVADA. Contactos programados: " & contactosProgramados)
                Else
                    EscribirLog("   ⚠️ Empresa NO activada completamente. Programados: " &
                            contactosProgramados & " de " & listaContactos.Count)
                End If

                rsEmpresas.MoveNext()
            End While

            rsEmpresas.Close()

            EscribirLog("=== AGENDA DE HOY GENERADA ===")
            EscribirLog("   Empresas activadas hoy: " & empresasTomadas)
            EscribirLog("   Espacio restante hoy: " & restantes & " correos")

        Catch ex As Exception
            EscribirLog("❌ Error en GenerarAgendaHoy_NuevasEmpresas: " & ex.Message)
        End Try
    End Sub



    Sub ProbarFormatoMensajes()
        ' Procedimiento para probar el formateo de mensajes
        Dim mensajePrueba As String
        Dim nombreCliente As String
        Dim empresaCliente As String
        Dim mensajeFinal As String

        ' Configurar datos de prueba
        nombreCliente = "Carlos Rodríguez"
        empresaCliente = "Restaurante El Mirador"

        ' Ejemplo 1: Mensaje de tu tabla (como está almacenado)
        mensajePrueba = "Hola, {nombre}." & vbCrLf &
                   "La mayoría de los restaurantes descubre que sus cámaras NO grabaron justo."

        ' Procesar el mensaje
        mensajeFinal = FormatearMensajeParaEnvio(mensajePrueba, nombreCliente, empresaCliente)

        Debug.Print("=== EJEMPLO 1 ===")
        Debug.Print(mensajeFinal)
        Debug.Print("")

        ' Ejemplo 2: Mensaje que ya tiene HTML
        mensajePrueba = "<p>Hola, {nombre}.</p><p>Solo para confirmar si recibiste un mensaje.</p><p>La pregunta sigue siendo:</p>"

        mensajeFinal = FormatearMensajeParaEnvio(mensajePrueba, nombreCliente, empresaCliente)

        Debug.Print("=== EJEMPLO 2 ===")
        Debug.Print(mensajeFinal)
        Debug.Print("")

        ' Ejemplo 3: Mensaje con {empresa}
        mensajePrueba = "Hola, {nombre}." & vbCrLf &
                   "Si hoy pasa algo en {empresa} y las cámaras no grabaron, no hay defensa."

        mensajeFinal = FormatearMensajeParaEnvio(mensajePrueba, nombreCliente, empresaCliente)

        Debug.Print("=== EJEMPLO 3 ===")
        Debug.Print(mensajeFinal)

        MsgBox("Pruebas completadas. Revisa la ventana Immediate (Ctrl+G)")
    End Sub


    'Imports System.Text.RegularExpressions

    Public Function FormatearMensajeParaEnvio(mensajeOriginal As String, nombre As String, Optional empresa As String = "") As String
        ' Esta función toma el mensaje de tu tabla y lo convierte en la versión elegante con HTML
        ' Reemplaza variables y asegura formato correcto con firma profesional

        Dim mensajeFormateado As String = ""
        Dim lineas() As String
        Dim i As Integer

        ' 1. Si el mensaje ya tiene etiquetas <p>, mantenerlas
        ' Si no, agregarlas alrededor de párrafos
        If mensajeOriginal.IndexOf("<p>", StringComparison.OrdinalIgnoreCase) = -1 Then
            ' El mensaje está en texto plano - convertir a HTML
            lineas = mensajeOriginal.Split(New String() {vbCrLf}, StringSplitOptions.None)

            For i = 0 To lineas.Length - 1
                If lineas(i).Trim() <> "" Then
                    mensajeFormateado &= "<p>" & lineas(i).Trim() & "</p>"
                End If
            Next
        Else
            ' Ya tiene HTML, usar tal cual
            mensajeFormateado = mensajeOriginal
        End If

        ' 2. Reemplazar {nombre} con el nombre real
        mensajeFormateado = mensajeFormateado.Replace("{nombre}", nombre)

        ' 3. Si hay variable {empresa} y se proporcionó empresa, reemplazar
        If empresa <> "" Then
            mensajeFormateado = mensajeFormateado.Replace("{empresa}", empresa)
        Else
            ' Si no hay empresa, quitar cualquier {empresa} que quede
            mensajeFormateado = mensajeFormateado.Replace("{empresa}", "tu empresa")
        End If

        ' 4. LIMPIAR ERRORES DE FORMATO
        mensajeFormateado = mensajeFormateado.Replace("<sup>+</sup>", "")

        If empresa <> "" Then
            mensajeFormateado = mensajeFormateado.Replace("(empresario) no se ha sido sufrido", empresa & " tendrás con qué respaldarte")
        Else
            mensajeFormateado = mensajeFormateado.Replace("(empresario) no se ha sido sufrido", "tu empresa tendrás con qué respaldarte")
        End If

        mensajeFormateado = mensajeFormateado.Replace("no se ha sido sufrido", "tendrás evidencia")

        ' 5. CONVERTIR ENLACE DE WHATSAPP A LINK HTML
        ' Buscar cualquier URL que contenga wa.me/
        Dim startIndex As Integer = mensajeFormateado.IndexOf("https://wa.me/", StringComparison.OrdinalIgnoreCase)

        While startIndex >= 0
            ' Encontrar el final del enlace (espacio, </p>, o fin)
            Dim endIndex As Integer = mensajeFormateado.IndexOf(" ", startIndex)
            If endIndex = -1 Then endIndex = mensajeFormateado.IndexOf("</p>", startIndex)
            If endIndex = -1 Then endIndex = mensajeFormateado.Length

            Dim url As String = mensajeFormateado.Substring(startIndex, endIndex - startIndex)

            ' Crear link HTML elegante
            Dim htmlLink As String = "<a href=""" & url & """ style=""color: #25D366; text-decoration: none; font-weight: bold;"">👉 Agendar llamada por WhatsApp</a>"

            ' Reemplazar URL por link HTML
            mensajeFormateado = mensajeFormateado.Replace(url, htmlLink)

            ' Buscar siguiente URL
            startIndex = mensajeFormateado.IndexOf("https://wa.me/", startIndex + htmlLink.Length, StringComparison.OrdinalIgnoreCase)
        End While

        ' 6. AGREGAR FIRMA PROFESIONAL SI NO LA TIENE
        If mensajeFormateado.IndexOf("Saludos,</p>", StringComparison.OrdinalIgnoreCase) = -1 Then
            mensajeFormateado &= "<br/><p>Saludos,</p>" &
                           "<p><strong>Abraham Cattan</strong><br/>" &
                           "Avital IT Services<br/>" &
                           "📞 55 4284 2359<br/>" &
                           "📍 CDMX, México</p>"
        End If

        ' 7. AGREGAR FOOTER DE PRIVACIDAD SI ES MODO REAL (no prueba)
        If Not mensajeFormateado.Contains("[PRUEBA]") AndAlso Not mensajeFormateado.Contains("MODO PRUEBA") Then
            mensajeFormateado &= "<hr style=""margin: 20px 0; border: 1px solid #eee;""/>" &
                           "<p style=""font-size: 11px; color: #666;"">" &
                           "Este correo fue enviado a " & nombre & " en " & empresa & ". " &
                           "Si no deseas recibir más comunicaciones, " &
                           "<a href=""mailto:abraham@avital.mx?subject=Baja de lista"">haz clic aquí</a>.</p>"
        End If

        Return mensajeFormateado
    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim respuesta As DialogResult = MessageBox.Show(
        "¿ESTÁS SEGURO de enviar correos REALES?" & vbCrLf & vbCrLf &
        "Se enviarán correos a contactos REALES, no a Abraham." & vbCrLf &
        "Recomendación: Primero prueba con 2-3 contactos." & vbCrLf & vbCrLf &
        "¿Continuar?",
        "⚠️ CONFIRMAR ENVÍOS REALES",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning)

        If respuesta = DialogResult.Yes Then
            Try
                If ba.State <> 1 Then
                    ba.Open(rutabd)
                End If

                EscribirLog("=== PRIMER ENVÍO REAL - INICIANDO ===")

                ' PRIMERO: Limitar a solo 3 correos para prueba segura
                Dim rs As New ADODB.Recordset
                rs.Open("SELECT TOP 3 * FROM envios_programados WHERE enviado = False ORDER BY fecha_programada", ba,
                    ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                If Not rs.EOF Then
                    ' Ejecutar envíos REALES limitados
                    EnviarAgendaDeHoy_MODO_REAL_LIMITADO()
                Else
                    MessageBox.Show("No hay correos programados para enviar.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                rs.Close()

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                If ba.State = 1 Then ba.Close()
            End Try
        End If
    End Sub

    Private Sub EnviarAgendaDeHoy_MODO_REAL_LIMITADO()
        EscribirLog("=== ENVÍOS REALES LIMITADOS (3 correos) ===")

        Try
            Dim rs As New ADODB.Recordset
            Dim sql As String =
        "SELECT TOP 3 ep.id, ep.contacto_id, ep.secuencia_id, ep.mensaje_id, ep.fecha_programada, " &
        "p.nombrecompleto, p.puesto, p.mail, e.cia, e.segmentoavital, m.asunto, m.cuerpo " &
        "FROM ((envios_programados ep " &
        "INNER JOIN prospectos p ON ep.contacto_id = p.id) " &
        "INNER JOIN listaempresas e ON p.cia = e.cia) " &
        "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
        "WHERE ep.enviado = False " &
        "ORDER BY ep.fecha_programada"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.EOF Then
                EscribirLog("ℹ️ No hay envíos programados.")
                rs.Close()
                Return
            End If

            ' Obtener cuenta de envío
            Dim rsCuenta As New ADODB.Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital WHERE limite > 0", ba,
                  ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                EscribirLog("❌ No hay cuentas disponibles.")
                rsCuenta.Close()
                rs.Close()
                Return
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            ' Configurar SMTP2Go
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            Dim enviados As Integer = 0
            Dim errores As Integer = 0

            rs.MoveFirst()
            While Not rs.EOF
                Dim nombre As String = rs.Fields("nombrecompleto").Value.ToString()
                Dim puesto As String = rs.Fields("puesto").Value.ToString()
                Dim mailReal As String = rs.Fields("mail").Value.ToString()
                Dim empresa As String = rs.Fields("cia").Value.ToString()
                Dim asuntoBase As String = rs.Fields("asunto").Value.ToString()
                Dim cuerpoBase As String = rs.Fields("cuerpo").Value.ToString()
                Dim contactoId As Integer = CInt(rs.Fields("contacto_id").Value)
                Dim mensajeId As Integer = CInt(rs.Fields("mensaje_id").Value)

                Try
                    ' Personalizar usando la función de formateo
                    Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa).Replace("{puesto}", puesto)
                    Dim cuerpo As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

                    Dim mail As New Net.Mail.MailMessage
                    mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
                    mail.To.Add(mailReal)   ' CORREO REAL
                    mail.Subject = asunto
                    mail.Body = cuerpo
                    mail.IsBodyHtml = True

                    smtpClient.Send(mail)
                    enviados += 1

                    ' Marcar como enviado
                    MarcarCorreoEnviado(contactoId, mensajeId)

                    EscribirLog("✅ [" & enviados & "] Enviado REAL a " & nombre & " (" & mailReal & ")")

                Catch ex As Exception
                    errores += 1
                    EscribirLog("❌ Error enviando a " & mailReal & ": " & ex.Message)
                End Try

                rs.MoveNext()
            End While

            rs.Close()

            EscribirLog("=== FIN ENVÍOS LIMITADOS ===")
            EscribirLog("   Enviados: " & enviados)
            EscribirLog("   Errores: " & errores)

            MessageBox.Show("✅ Se enviaron " & enviados & " correos REALES (versión limitada)." & vbCrLf &
                      "Errores: " & errores & vbCrLf & vbCrLf &
                      "Revisa tu bandeja de entrada y spam.",
                      "Primera Prueba Real",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information)

        Catch ex As Exception
            EscribirLog("❌ Error en EnviarAgendaDeHoy_MODO_REAL_LIMITADO: " & ex.Message)
        End Try
    End Sub

    Private Sub EnviarCorreosPendientesConRetraso()
        ' Este método envía correos pendientes con retrasos aleatorios entre ellos

        Try
            ' Obtener correos que deberían haberse enviado YA (fecha_programada <= ahora)
            Dim ahora As DateTime = DateTime.Now
            Dim rs As New ADODB.Recordset
            Dim sql As String =
        "SELECT COUNT(*) as total FROM envios_programados " &
        "WHERE fecha_programada <= #" & ahora.ToString("yyyy-MM-dd HH:mm:ss") & "# " &
        "AND enviado = False"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim pendientes As Integer = 0
            If Not rs.EOF Then
                pendientes = CInt(rs.Fields("total").Value)
            End If
            rs.Close()

            If pendientes = 0 Then
                ' EscribirLog("ℹ️ No hay correos atrasados para enviar")
                Return
            End If

            EscribirLog("📬 Encontré " & pendientes & " correo(s) atrasado(s). Enviando con retraso escalonado...")

            ' Obtener lista de correos pendientes
            sql = "SELECT TOP 10 ep.id, ep.contacto_id, ep.mensaje_id, " &
              "p.nombrecompleto, p.mail, e.cia, m.asunto, m.cuerpo " &
              "FROM ((envios_programados ep " &
              "INNER JOIN prospectos p ON ep.contacto_id = p.id) " &
              "INNER JOIN listaempresas e ON p.cia = e.cia) " &
              "INNER JOIN mensajes m ON ep.mensaje_id = m.id " &
              "WHERE ep.fecha_programada <= #" & ahora.ToString("yyyy-MM-dd HH:mm:ss") & "# " &
              "AND ep.enviado = False " &
              "ORDER BY ep.fecha_programada"

            rs.Open(sql, ba, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.EOF Then
                rs.Close()
                Return
            End If

            ' Obtener cuenta de envío
            Dim rsCuenta As New ADODB.Recordset
            rsCuenta.Open("SELECT TOP 1 correo, nombre FROM mailsavital WHERE limite > 0", ba,
                  ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If rsCuenta.EOF Then
                rsCuenta.Close()
                rs.Close()
                Return
            End If

            Dim correoEnvio As String = rsCuenta.Fields("correo").Value.ToString()
            Dim nombreRemitente As String = rsCuenta.Fields("nombre").Value.ToString()
            rsCuenta.Close()

            ' Configurar SMTP2Go
            Dim smtpClient As New Net.Mail.SmtpClient("mail.smtp2go.com")
            smtpClient.Port = 2525
            smtpClient.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")
            smtpClient.EnableSsl = True

            Dim rand As New Random()
            Dim enviadosEnEsteCiclo As Integer = 0

            rs.MoveFirst()
            While Not rs.EOF AndAlso enviadosEnEsteCiclo < 3 ' Máximo 3 por ciclo
                Dim nombre As String = rs.Fields("nombrecompleto").Value.ToString()
                Dim mailReal As String = rs.Fields("mail").Value.ToString()
                Dim empresa As String = rs.Fields("cia").Value.ToString()
                Dim asuntoBase As String = rs.Fields("asunto").Value.ToString()
                Dim cuerpoBase As String = rs.Fields("cuerpo").Value.ToString()
                Dim contactoId As Integer = CInt(rs.Fields("contacto_id").Value)
                Dim mensajeId As Integer = CInt(rs.Fields("mensaje_id").Value)

                Try
                    ' Personalizar mensaje
                    Dim asunto As String = asuntoBase.Replace("{nombre}", nombre).Replace("{empresa}", empresa)
                    Dim cuerpo As String = FormatearMensajeParaEnvio(cuerpoBase, nombre, empresa)

                    Dim mail As New Net.Mail.MailMessage
                    mail.From = New Net.Mail.MailAddress(correoEnvio, nombreRemitente & " - Avital IT Services")
                    mail.To.Add(mailReal)
                    mail.Subject = asunto
                    mail.Body = cuerpo
                    mail.IsBodyHtml = True

                    ' AGREGAR RETRASO ALEATORIO entre envíos (60-180 segundos)
                    If enviadosEnEsteCiclo > 0 Then
                        Dim retraso As Integer = rand.Next(180, 500)
                        EscribirLog("   ⏳ Esperando " & retraso & " segundos antes del siguiente envío...")
                        Threading.Thread.Sleep(retraso * 1000) ' Convertir a milisegundos
                    End If

                    smtpClient.Send(mail)
                    enviadosEnEsteCiclo += 1

                    ' Marcar como enviado
                    MarcarCorreoEnviado(contactoId, mensajeId)

                    EscribirLog("✅ [" & enviadosEnEsteCiclo & "] Enviado REAL a " & nombre &
                          " (" & mailReal & ") con retraso escalonado")

                Catch ex As Exception
                    EscribirLog("❌ Error enviando a " & mailReal & ": " & ex.Message)
                End Try

                rs.MoveNext()
            End While

            rs.Close()

            If enviadosEnEsteCiclo > 0 Then
                EscribirLog("📤 Enviados " & enviadosEnEsteCiclo & " correo(s) con retraso escalonado")
            End If

        Catch ex As Exception
            EscribirLog("⚠️ Error en EnviarCorreosPendientesConRetraso: " & ex.Message)
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles btnCorreoPrueba.Click
        Try
            Dim correoDestino As String = txtCorreoPrueba.Text.Trim()

            If correoDestino = "" Then
                MessageBox.Show("Escribe un correo destino para la prueba.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Dim smtp As New Net.Mail.SmtpClient()
            smtp.Host = "mail.smtp2go.com"
            smtp.Port = 2525
            smtp.EnableSsl = True
            smtp.Credentials = New Net.NetworkCredential("avitalsmtp2go", "Parnasa1$")

            Dim mail As New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress("ventas1@avital.com.mx", "Avital IT Services")
            mail.To.Add(correoDestino)
            mail.Subject = "PRUEBA SMTP2GO - Avital"
            mail.Body =
                "Este es un correo de prueba enviado desde el sistema Avital." & vbCrLf &
                "Fecha: " & Now.ToString("yyyy-MM-dd HH:mm:ss")

            smtp.Send(mail)

            MessageBox.Show("✅ Correo de prueba enviado correctamente.", "SMTP OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            EscribirLog("✅ SMTP PRUEBA OK -> " & correoDestino)

        Catch ex As Exception
            MessageBox.Show("❌ Error enviando correo de prueba:" & vbCrLf & ex.Message, "ERROR SMTP", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EscribirLog("❌ ERROR SMTP PRUEBA: " & ex.Message)
        End Try
    End Sub
End Class
