Imports System.Net.Mail
Imports System.IO

Module Traspasos
    Dim Tasa As Decimal
    Dim Contador As Integer = 0
    Dim AplicaGarantiaLIQ As String



    Sub Main()
        Console.WriteLine("Iniciando ...")
        Dim Fecha As String = Date.Now.ToString("yyyyMMdd")
        Dim FechaD As DateTime = Date.Now.Date

        TraspasoCarteraVencida(FechaD)
        'Fecha = "20150506" 'PARA PRUEBAS
        Console.WriteLine("Procesando Avio Paso 1 ...")
        Calcula_Saldos(Fecha, "H")
        Console.WriteLine("Procesando Avio Paso 2 ...")
        TraspasosAvio(Fecha, "H")
        'Fecha = "20150506" 'PARA PRUEBAS
        Console.WriteLine("Procesando CC Paso 1...")
        Calcula_Saldos(Fecha, "C")
        Console.WriteLine("Procesando CC Paso 2...")
        TraspasosAvio(Fecha, "C")
        Console.WriteLine("Terminado ...")
        EnviaError("ecacerest@lamoderna.com.mx", "Ejecucion de Traspasos " & Fecha & " = " & Contador, "Ejecucion de Traspasos " & Date.Now.ToString)
    End Sub

    Sub TraspasosAvio(Fecha As String, Tipo As String)
        Try
            Dim InteresDias As Decimal = 0

            Dim TaS As New ProduccionDSTableAdapters.TraspasosAvioCCTableAdapter
            Dim ta As New ProduccionDSTableAdapters.SaldosAvioCCTableAdapter
            Dim T As New ProduccionDS.SaldosAvioCCDataTable
            Dim r As ProduccionDS.SaldosAvioCCRow

            If Tipo = "H" Then
                TaS.DeleteFechaAvio(Fecha)
                ta.FillAvio(T, Fecha)
            Else
                TaS.DeleteFechaCC(Fecha)
                ta.FillCC(T, Fecha)
            End If
            For Each r In T.Rows
                Console.WriteLine(r.AnexoCon & " - " & r.CicloPagare)
                Contador += 1
                FijaTasa(r.Anexo, r.Ciclo, CadenaFecha(Fecha))
                InteresDias = CalculaInteres(r)
                TaS.Insert(r.Anexo, r.Ciclo, r.Imp + r.Fega + r.Intereses + InteresDias, r.Imp, r.Intereses, r.Garantia, r.Tipar, r.Fega, Fecha, InteresDias)
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx", "Error en traspasos", ex.Message)
        End Try
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, Asunto) = 0 Then
            Dim Mensage As New MailMessage("TraspasosAvio@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
            Try
                Cliente.Send(Mensage)
            Catch ex As Exception
                'ReportError(ex)
            End Try
        Else
            Console.WriteLine("No se ha encontrado la ruta de acceso de la red")
        End If
    End Sub

    Sub FijaTasa(ByVal a As String, ByVal c As String, ByVal f As Date)
        Dim x As New ProduccionDSTableAdapters.AviosTableAdapter
        Dim y As New ProduccionDS.AviosDataTable
        Dim Diferencial As Double
        x.UpdateAplicaGAR()
        x.FillAnexo(y, a, c)
        Dim v As String = y.Rows(0).Item(0)
        If v = "7" Then
            Tasa = y.Rows(0).Item(5)
        Else
            Dim TIIE As New ProduccionDSTableAdapters.TIIEpromedioTableAdapter
            Diferencial = y.Rows(0).Item(1)

            Tasa = TIIE.SacaTIIE(DateAdd(DateInterval.Month, -1, f).ToString("yyyyMM")) + Diferencial
        End If
        'SOLICITADO POR ELISANDER DEBIDO A PRORROGA, SE SUMA UN PUNTO PORCENTUAL A PARTIR DE ENERO++++++++++++
        If (a = "070320012" Or a = "070860007" Or a = "070790010" Or a = "070780012" Or a = "070600011" Or a = "071330006") And f >= CDate("01/01/2015") Then
            Tasa += 1
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        AplicaGarantiaLIQ = y.Rows(0).Item("AplicaGarantiaLIQ")
    End Sub

    Function CadenaFecha(ByVal f As String) As Date
        Dim ff As New System.DateTime(CInt(Mid(f, 1, 4)), Mid(f, 5, 2), Mid(f, 7, 2))
        Return ff
    End Function

    Function CalculaInteres(ByVal r As ProduccionDS.SaldosAvioCCRow)
        Dim dias As Integer = 0
        Dim Inte As Decimal = 0
        dias = DateDiff(DateInterval.Day, CadenaFecha(r.FechaFinal), CadenaFecha(r.FechaTerminacion))
        If dias < 0 Then dias = 0
        Inte = Math.Round((r.Saldo) * (Tasa / 100 / 360) * dias, 2)
        Return Inte
    End Function

    Sub Calcula_Saldos(Fecha As String, Tipo As String)
        Try
            Dim ta As New ProduccionDSTableAdapters.SaldosAvioCCTableAdapter
            Dim T As New ProduccionDS.SaldosAvioCCDataTable
            Dim r As ProduccionDS.SaldosAvioCCRow

            If Tipo = "H" Then
                ta.FillAvio(T, Fecha)
            Else
                ta.FillCC(T, Fecha)
            End If
            For Each r In T.Rows
                Console.WriteLine(r.AnexoCon & " - " & r.CicloPagare)
                If Directory.Exists("D:\Contratos$\") Then
                    Shell("""D:\Contratos$\Executables\\EstadoCuentaAVCC.exe"" " & r.Anexo & " " & r.Ciclo & " FIN 0", AppWinStyle.NormalFocus, True)
                Else
                    Shell("""F:\Executables\EstadoCuentaAVCC.exe"" " & r.Anexo & " " & r.Ciclo & " FIN 0", AppWinStyle.NormalFocus, True)
                End If
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx", "Error en traspasos", ex.Message)
        End Try

    End Sub

    Sub TraspasoCarteraVencida(Fecha As DateTime)
        Dim TaVenc As New ProduccionDSTableAdapters.CarteraVencidaDETTableAdapter
        Dim TaTrasp As New ProduccionDSTableAdapters.TraspasosVencidosTableAdapter
        Dim DS As New ProduccionDS
        Dim FechaAPP As DateTime = TaVenc.FechaAplicacion
        Dim RR As ProduccionDS.TraspasosVencidosRow
        FechaAPP = FechaAPP.AddDays(-1)
        TaVenc.Fill(DS.CarteraVencidaDET, FechaAPP)
        For Each r As ProduccionDS.CarteraVencidaDETRow In DS.CarteraVencidaDET.Rows
            Select Case r.TipoCredito.Trim
                Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "FULL SERVICE", "ARRENDAMIENTO PURO"
                    If r.Dias >= 30 Then
                        RR = DS.TraspasosVencidos.NewRow
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "CUENTA CORRIENTE"
                    If r.Dias >= 60 Then
                        RR = DS.TraspasosVencidos.NewRow
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "ARRENDAMIENTO FINANCIERO", "CREDITO REFACCIONARIO", "CREDITO SIMPLE"
                    If r.Dias >= 90 Then
                        RR = DS.TraspasosVencidos.NewRow
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
            End Select
        Next



    End Sub

End Module
