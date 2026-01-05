<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Friend WithEvents Timer1 As Timer
    Friend WithEvents txtlog As TextBox
    Friend WithEvents lblestado As Label

    Friend WithEvents btniniciar As Button
    Friend WithEvents btnDetener As Button
    Friend WithEvents btnLimpiarLog As Button
    Friend WithEvents btnCicloManual As Button
    Friend WithEvents btnVerEstado As Button

    Friend WithEvents btnGenerarAgendaHoy As Button
    Friend WithEvents btnProbarEnviosHoy As Button

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button

    Friend WithEvents grpPrincipal As GroupBox
    Friend WithEvents grpPruebas As GroupBox
    Friend WithEvents grpAgenda As GroupBox

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Timer1 = New Timer(components)
        txtlog = New TextBox()
        lblestado = New Label()
        btniniciar = New Button()
        btnDetener = New Button()
        btnLimpiarLog = New Button()
        btnCicloManual = New Button()
        btnVerEstado = New Button()
        btnGenerarAgendaHoy = New Button()
        btnProbarEnviosHoy = New Button()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        Button4 = New Button()
        Button5 = New Button()
        Button6 = New Button()
        grpPrincipal = New GroupBox()
        grpPruebas = New GroupBox()
        grpAgenda = New GroupBox()
        btnCorreoPrueba = New Button()
        txtcorreoprueba = New TextBox()
        grpPrincipal.SuspendLayout()
        grpPruebas.SuspendLayout()
        grpAgenda.SuspendLayout()
        SuspendLayout()
        ' 
        ' Timer1
        ' 
        Timer1.Interval = 120000
        ' 
        ' txtlog
        ' 
        txtlog.Font = New Font("Consolas", 9F)
        txtlog.Location = New Point(12, 272)
        txtlog.Multiline = True
        txtlog.Name = "txtlog"
        txtlog.ScrollBars = ScrollBars.Vertical
        txtlog.Size = New Size(980, 420)
        txtlog.TabIndex = 4
        txtlog.WordWrap = False
        ' 
        ' lblestado
        ' 
        lblestado.AutoSize = True
        lblestado.Location = New Point(12, 9)
        lblestado.Name = "lblestado"
        lblestado.Size = New Size(169, 15)
        lblestado.TabIndex = 0
        lblestado.Text = "✅ Sistema listo - Click INICIAR"
        ' 
        ' btniniciar
        ' 
        btniniciar.Location = New Point(20, 25)
        btniniciar.Name = "btniniciar"
        btniniciar.Size = New Size(160, 30)
        btniniciar.TabIndex = 0
        btniniciar.Text = "INICIAR (Auto)"
        btniniciar.UseVisualStyleBackColor = True
        ' 
        ' btnDetener
        ' 
        btnDetener.Location = New Point(190, 25)
        btnDetener.Name = "btnDetener"
        btnDetener.Size = New Size(140, 30)
        btnDetener.TabIndex = 1
        btnDetener.Text = "DETENER"
        btnDetener.UseVisualStyleBackColor = True
        ' 
        ' btnLimpiarLog
        ' 
        btnLimpiarLog.Location = New Point(340, 25)
        btnLimpiarLog.Name = "btnLimpiarLog"
        btnLimpiarLog.Size = New Size(140, 30)
        btnLimpiarLog.TabIndex = 2
        btnLimpiarLog.Text = "LIMPIAR LOG"
        btnLimpiarLog.UseVisualStyleBackColor = True
        ' 
        ' btnCicloManual
        ' 
        btnCicloManual.Location = New Point(660, 25)
        btnCicloManual.Name = "btnCicloManual"
        btnCicloManual.Size = New Size(180, 30)
        btnCicloManual.TabIndex = 4
        btnCicloManual.Text = "CICLO MANUAL (TEST)"
        btnCicloManual.UseVisualStyleBackColor = True
        ' 
        ' btnVerEstado
        ' 
        btnVerEstado.Location = New Point(490, 25)
        btnVerEstado.Name = "btnVerEstado"
        btnVerEstado.Size = New Size(160, 30)
        btnVerEstado.TabIndex = 3
        btnVerEstado.Text = "VER ESTADO"
        btnVerEstado.UseVisualStyleBackColor = True
        ' 
        ' btnGenerarAgendaHoy
        ' 
        btnGenerarAgendaHoy.Location = New Point(20, 25)
        btnGenerarAgendaHoy.Name = "btnGenerarAgendaHoy"
        btnGenerarAgendaHoy.Size = New Size(220, 30)
        btnGenerarAgendaHoy.TabIndex = 0
        btnGenerarAgendaHoy.Text = "GENERAR AGENDA HOY + BUFFER"
        btnGenerarAgendaHoy.UseVisualStyleBackColor = True
        ' 
        ' btnProbarEnviosHoy
        ' 
        btnProbarEnviosHoy.Location = New Point(250, 25)
        btnProbarEnviosHoy.Name = "btnProbarEnviosHoy"
        btnProbarEnviosHoy.Size = New Size(220, 30)
        btnProbarEnviosHoy.TabIndex = 1
        btnProbarEnviosHoy.Text = "PROBAR ENVIAR 1 (AHORA)"
        btnProbarEnviosHoy.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(20, 30)
        Button1.Name = "Button1"
        Button1.Size = New Size(150, 30)
        Button1.TabIndex = 0
        Button1.Text = "Button1"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(180, 30)
        Button2.Name = "Button2"
        Button2.Size = New Size(150, 30)
        Button2.TabIndex = 1
        Button2.Text = "Button2"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(340, 30)
        Button3.Name = "Button3"
        Button3.Size = New Size(150, 30)
        Button3.TabIndex = 2
        Button3.Text = "Button3"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Button4
        ' 
        Button4.Location = New Point(500, 30)
        Button4.Name = "Button4"
        Button4.Size = New Size(150, 30)
        Button4.TabIndex = 3
        Button4.Text = "Button4"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(660, 30)
        Button5.Name = "Button5"
        Button5.Size = New Size(150, 30)
        Button5.TabIndex = 4
        Button5.Text = "Button5"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' Button6
        ' 
        Button6.Location = New Point(820, 30)
        Button6.Name = "Button6"
        Button6.Size = New Size(150, 30)
        Button6.TabIndex = 5
        Button6.Text = "Button6"
        Button6.UseVisualStyleBackColor = True
        ' 
        ' grpPrincipal
        ' 
        grpPrincipal.Controls.Add(btniniciar)
        grpPrincipal.Controls.Add(btnDetener)
        grpPrincipal.Controls.Add(btnLimpiarLog)
        grpPrincipal.Controls.Add(btnVerEstado)
        grpPrincipal.Controls.Add(btnCicloManual)
        grpPrincipal.Location = New Point(12, 32)
        grpPrincipal.Name = "grpPrincipal"
        grpPrincipal.Size = New Size(980, 70)
        grpPrincipal.TabIndex = 1
        grpPrincipal.TabStop = False
        grpPrincipal.Text = "Operación"
        ' 
        ' grpPruebas
        ' 
        grpPruebas.Controls.Add(Button1)
        grpPruebas.Controls.Add(Button2)
        grpPruebas.Controls.Add(Button3)
        grpPruebas.Controls.Add(Button4)
        grpPruebas.Controls.Add(Button5)
        grpPruebas.Controls.Add(Button6)
        grpPruebas.Location = New Point(12, 184)
        grpPruebas.Name = "grpPruebas"
        grpPruebas.Size = New Size(980, 82)
        grpPruebas.TabIndex = 3
        grpPruebas.TabStop = False
        grpPruebas.Text = "Pruebas internas (Button1..6)"
        ' 
        ' grpAgenda
        ' 
        grpAgenda.Controls.Add(txtcorreoprueba)
        grpAgenda.Controls.Add(btnCorreoPrueba)
        grpAgenda.Controls.Add(btnGenerarAgendaHoy)
        grpAgenda.Controls.Add(btnProbarEnviosHoy)
        grpAgenda.Location = New Point(12, 108)
        grpAgenda.Name = "grpAgenda"
        grpAgenda.Size = New Size(980, 70)
        grpAgenda.TabIndex = 2
        grpAgenda.TabStop = False
        grpAgenda.Text = "Agenda (nuevos) - Mensaje 1"
        ' 
        ' btnCorreoPrueba
        ' 
        btnCorreoPrueba.Location = New Point(476, 22)
        btnCorreoPrueba.Name = "btnCorreoPrueba"
        btnCorreoPrueba.Size = New Size(220, 30)
        btnCorreoPrueba.TabIndex = 2
        btnCorreoPrueba.Text = "PROBAR ENVIAR 1 (AHORA)"
        btnCorreoPrueba.UseVisualStyleBackColor = True
        ' 
        ' txtcorreoprueba
        ' 
        txtcorreoprueba.Location = New Point(746, 23)
        txtcorreoprueba.Name = "txtcorreoprueba"
        txtcorreoprueba.Size = New Size(100, 23)
        txtcorreoprueba.TabIndex = 3
        txtcorreoprueba.Text = "abraham@avital.mx"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1006, 705)
        Controls.Add(txtlog)
        Controls.Add(grpPruebas)
        Controls.Add(grpAgenda)
        Controls.Add(grpPrincipal)
        Controls.Add(lblestado)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Avital Prospectación"
        grpPrincipal.ResumeLayout(False)
        grpPruebas.ResumeLayout(False)
        grpAgenda.ResumeLayout(False)
        grpAgenda.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnCorreoPrueba As Button
    Friend WithEvents txtcorreoprueba As TextBox

End Class
