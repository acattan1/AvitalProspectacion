Module Module1
    ' VARIABLES GLOBALES DE CONEXIÓN
    Public viejo As Integer
    Public rutabd As String
    Public mitzraim, yerushalaim, paro, egipto As String
    Public q As VariantType
    Public v, ve As String
    Public iq As Integer
    Public ba, ba1, ba2 As New ADODB.Connection

    ' VARIABLES DE TABLAS EXISTENTES
    Public mailsavital, prospectos, correoenviado, mailsinternos, bajas, listaempresas, megalopolis As ADODB.Recordset

    ' VARIABLES GLOBALES DEL SISTEMA
    Public numerotarea, totis As Object
    Public mayor, pq As Integer
    Public fecha1, fecha2, fecha3, fecha4 As Date
    Public hora1, hora2, hora3, hora4 As DateTime
    Public horainicio, horafinal, currenttime As DateTime

    ' NUEVAS VARIABLES PARA EL SISTEMA DE SECUENCIAS
    Public secuencias_maestra, mensajes, envios_programados, roles_horarios As ADODB.Recordset
    ' ✅ Límites y reglas
    Public limite_diario_envios As Integer = 110
    Public buffer_manana As Integer = 25

    Public maxDiaSeguimiento As Integer = 2

    ' Ventana principal de envíos (Mar-Jue)
    Public hora_inicio_envios As String = "06:40"
    Public hora_fin_envios As String = "14:00"

    ' Ventana adicional para NO-CEO (Mensajes 2-5) - Lunes 12:00–15:00
    Public hora_inicio_lunes As String = "12:00"
    Public hora_fin_lunes As String = "15:00"

    ' Gaps humanos entre correos (se reflejan en fecha_programada)
    Public gap_min_seg As Integer = 40
    Public gap_max_seg As Integer = 300


End Module
