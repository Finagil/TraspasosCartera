Imports System.Net.Mail
Imports System.IO

Module Traspasos
    Dim Tasa As Decimal
    Dim Contador As Integer = 0
    Dim AplicaGarantiaLIQ As String
    Dim FechaS As String
    Dim FechaD As DateTime
    Dim DS As New ProduccionDS
    Dim Vencido As Boolean
    Dim TaTrasp As New ProduccionDSTableAdapters.TraspasosVencidosTableAdapter
    Dim Arg() As String

    Sub Main()
        Arg = Environment.GetCommandLineArgs()
        Console.WriteLine("Iniciando ...")
        FechaD = TaTrasp.FechaAplicacion
        FechaD = FechaD.Date
        Console.WriteLine(FechaD.Date.ToShortDateString)
        If Arg.Length > 1 Then
            If Arg(1) = "V" Then
                CorreTraspasos(True)
            End If
        End If

        If Date.Now.Hour > 18 Then ' LOS TRAPASOS SE EJECUTAN POR LA TARDE
            If Date.Now.DayOfWeek = DayOfWeek.Sunday Or Date.Now.DayOfWeek = DayOfWeek.Saturday Then
                ' no se generan traspasos
            ElseIf FechaD.DayOfWeek = DayOfWeek.Monday Then
                ' el lunes se genneran traspasos de sabado y domingo
                FechaS = FechaD.AddDays(-2).Date.ToString("yyyyMMdd") ' sabado
                CorreTraspasos(False)
                FechaS = FechaD.AddDays(-1).Date.ToString("yyyyMMdd") ' Domingo
                CorreTraspasos(False)
                FechaS = FechaD.Date.ToString("yyyyMMdd") ' Lunes
                CorreTraspasos(False)
            Else
                FechaS = FechaD.Date.ToString("yyyyMMdd")
                CorreTraspasos(False)
            End If
        End If
        Console.WriteLine("Terminado ...")
        EnviaError("ecacerest@lamoderna.com.mx", "Ejecucion de Traspasos " & FechaS & " = " & Contador, "Ejecucion de Traspasos " & Date.Now.ToString)
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
            EnviaError("ecacerest@lamoderna.com.mx", "Error en traspasos1", ex.Message)
        End Try
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, Asunto) = 0 Then
            Dim Mensage As New MailMessage("TraspasosAvio@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient(My.Settings.SMTP, My.Settings.SMTP_port)
            Try
                Dim Credenciales As String() = My.Settings.SMTP_creden.Split(",")
                Cliente.Credentials = New System.Net.NetworkCredential(Credenciales(0), Credenciales(1), Credenciales(2))
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
        Dim Cad As String
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
                Console.WriteLine(My.Settings.RutaExecutables)
                Console.WriteLine(r.AnexoCon & " - " & r.CicloPagare)
                If Directory.Exists(My.Settings.RutaExecutables) Then
                    Console.WriteLine("""" & My.Settings.RutaExecutables & "EstadoCuentaAVCC.exe"" " & r.Anexo & " " & r.Ciclo & " FIN 0")
                    Shell("""" & My.Settings.RutaExecutables & "EstadoCuentaAVCC.exe"" " & r.Anexo & " " & r.Ciclo & " FIN 0", AppWinStyle.NormalFocus, True)
                Else
                    Console.WriteLine("""F:\Executables\EstadoCuentaAVCC.exe"" ")
                    Shell("""F:\Executables\EstadoCuentaAVCC.exe"" " & r.Anexo & " " & r.Ciclo & " FIN 0", AppWinStyle.NormalFocus, True)
                End If
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx", "Error en traspasos2", ex.Message & cad)
        End Try

    End Sub

    Sub TraspasoCarteraVencida(Fecha As DateTime)
        Dim TaVenc As New ProduccionDSTableAdapters.CarteraVencidaDETTableAdapter
        'Dim FechaAPP As DateTime = TaVenc.FechaAplicacion
        Dim Diass As Integer
        Dim RR As ProduccionDS.TraspasosVencidosRow

        'FechaAPP = FechaAPP.AddDays(-1)
        TaVenc.Fill(DS.CarteraVencidaDET, Fecha)
        For Each r As ProduccionDS.CarteraVencidaDETRow In DS.CarteraVencidaDET.Rows
            Console.WriteLine("Cartera Vencida " & r.AnexoCon)
            Vencido = False
            Select Case r.TipoCredito.Trim
                Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "FULL SERVICE", "ARRENDAMIENTO PURO"
                    If r.Dias >= 30 Then
                        RR = DS.TraspasosVencidos.NewRow
                        Select Case r.TipoCredito.Trim
                            Case "ANTICIPO AVÍO", "CREDITO DE AVÍO"
                                TraspasaAVCC(RR, r, 30 - r.Dias)
                            Case "FULL SERVICE", "ARRENDAMIENTO PURO"
                                TraspasaTRA(RR, r, 30 - r.Dias)
                        End Select
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "CUENTA CORRIENTE"
                    If r.Dias >= 60 Then
                        RR = DS.TraspasosVencidos.NewRow
                        TraspasaAVCC(RR, r, 60 - r.Dias)
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "ARRENDAMIENTO FINANCIERO", "CREDITO REFACCIONARIO", "CREDITO SIMPLE", "CREDITO LIQUIDEZ INMEDIATA"
                    If TaVenc.EsPagoUnicoInteresMensual(r.Anexo) = 1 Then
                        Diass = TaVenc.DiasCapital(Fecha, r.Anexo)
                        If Diass < 30 Then
                            Diass = 90
                        Else
                            r.Dias = Diass
                            Diass = 30
                        End If
                    Else
                        Diass = 90
                    End If

                    If r.Dias >= Diass Then
                        RR = DS.TraspasosVencidos.NewRow
                        TraspasaTRA(RR, r, Diass - r.Dias)
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
            End Select
            If Vencido = True Then
                DS.TraspasosVencidos.GetChanges()
                TaTrasp.Update(DS.TraspasosVencidos)
                Select Case r.TipoCredito.Trim
                    Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "CUENTA CORRIENTE"
                        TaVenc.MarcaVencidaAV("VENCIDA", r.Anexo, r.Ciclo)
                    Case Else
                        Dim BLOQ As Integer = DesBloqueaContrato(r.Anexo) 'DESBLOQUEO MESA DE CONTROL+++++++++++++
                        TaVenc.MarcaVencidaTRA("VENCIDA", r.Anexo)
                        BloqueaContrato(r.Anexo, BLOQ) '*******************BLOQUEO MESA DE CONTROL++++++++++++++++
                End Select
            End If
        Next
        ' POR REESTRUCTURAS
        DS.TraspasosVencidos.Clear()
        TaVenc.FillByREEST(DS.CarteraVencidaDET, Fecha)
        For Each r As ProduccionDS.CarteraVencidaDETRow In DS.CarteraVencidaDET.Rows
            Console.WriteLine("Cartera Vencida " & r.AnexoCon)
            Vencido = False
            Select Case r.TipoCredito.Trim
                Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "FULL SERVICE", "ARRENDAMIENTO PURO"
                    RR = DS.TraspasosVencidos.NewRow
                    Select Case r.TipoCredito.Trim
                        Case "ANTICIPO AVÍO", "CREDITO DE AVÍO"
                            TraspasaAVCC(RR, r, 0)
                        Case "FULL SERVICE", "ARRENDAMIENTO PURO"
                            TraspasaTRA(RR, r, 0)
                    End Select
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                Case "CUENTA CORRIENTE"
                    RR = DS.TraspasosVencidos.NewRow
                    TraspasaAVCC(RR, r, 0)
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                Case "ARRENDAMIENTO FINANCIERO", "CREDITO REFACCIONARIO", "CREDITO SIMPLE", "CREDITO LIQUIDEZ INMEDIATA"
                    RR = DS.TraspasosVencidos.NewRow
                    TraspasaTRA(RR, r, 0)
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
            End Select
            If Vencido = True Then
                DS.TraspasosVencidos.GetChanges()
                TaTrasp.Update(DS.TraspasosVencidos)
                Select Case r.TipoCredito.Trim
                    Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "CUENTA CORRIENTE"
                        TaVenc.MarcaVencidaAV("VENCIDA", r.Anexo, r.Ciclo)
                    Case Else
                        Dim BLOQ As Integer = DesBloqueaContrato(r.Anexo) 'DESBLOQUEO MESA DE CONTROL+++++++++++++
                        TaVenc.MarcaVencidaTRA("VENCIDA", r.Anexo)
                        BloqueaContrato(r.Anexo, BLOQ) '*******************BLOQUEO MESA DE CONTROL++++++++++++++++
                End Select
            Else
                DS.TraspasosVencidos.GetChanges()
                TaTrasp.Update(DS.TraspasosVencidos)
                Select Case r.TipoCredito.Trim
                    Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "CUENTA CORRIENTE"
                        TaVenc.MarcaVencidaAV("", r.Anexo, r.Ciclo)
                    Case Else
                        Dim BLOQ As Integer = DesBloqueaContrato(r.Anexo) 'DESBLOQUEO MESA DE CONTROL+++++++++++++
                        TaVenc.MarcaVencidaTRA("", r.Anexo)
                        BloqueaContrato(r.Anexo, BLOQ) '*******************BLOQUEO MESA DE CONTROL++++++++++++++++
                End Select
            End If
        Next

    End Sub

    Sub TraspasaTRA(ByRef RR As ProduccionDS.TraspasosVencidosRow, ByRef r As ProduccionDS.CarteraVencidaDETRow, dias As Integer)
        TaTrasp.DeleteAnexo(r.Anexo, r.Regreso)
        Dim EdoV As New ProduccionDSTableAdapters.EdoctavTableAdapter
        Dim EdoS As New ProduccionDSTableAdapters.EdoctasTableAdapter
        Dim EdoO As New ProduccionDSTableAdapters.EdoctaoTableAdapter
        EdoV.Fill(DS.Edoctav, r.Anexo)
        EdoS.Fill(DS.Edoctas, r.Anexo)
        EdoO.Fill(DS.Edoctao, r.Anexo)

        RR.Anexo = r.Anexo
        RR.Ciclo = ""
        RR.Tipo = r.Tipo
        RR.Fecha = FechaD.AddDays(dias)
        RR.Tipar = r.Tipar
        RR.Segmento_Negocio = r.Segmento_Negocio
        RR.Regreso = r.Regreso

        If DS.Edoctav.Rows.Count > 0 Then
            RR.SaldoInsoluto = DS.Edoctav.Rows(0).Item("Capital")
            RR.CargaFinanciera = DS.Edoctav.Rows(0).Item("Interes")
        Else
            RR.SaldoInsoluto = 0
            RR.CargaFinanciera = 0
        End If
        If DS.Edoctas.Rows.Count > 0 Then
            RR.SaldoInsolutoSEG = DS.Edoctas.Rows(0).Item("Capital")
            RR.CargaFinancieraSEG = DS.Edoctas.Rows(0).Item("Interes")
        Else
            RR.SaldoInsolutoSEG = 0
            RR.CargaFinancieraSEG = 0
        End If
        If DS.Edoctao.Rows.Count > 0 Then
            RR.SaldoInsolutoOTR = DS.Edoctao.Rows(0).Item("Capital")
            RR.CargaFinancieraOTR = DS.Edoctao.Rows(0).Item("Interes")
        Else
            RR.SaldoInsolutoOTR = 0
            RR.CargaFinancieraOTR = 0
        End If

        RR.CapitalVencido = 0
        RR.InteresVencido = 0
        RR.CapitalVencidoOt = 0
        RR.InteresVencidoOt = 0
        RR.IvaCapital = 0
        Dim Fact As New ProduccionDSTableAdapters.FacturasTableAdapter
        Fact.Fill(DS.Facturas, r.Anexo)
        Dim Pagado, Capital, Interes, InteresOt, CapitalOt, IvaCpital, IntereSEG As Decimal

        For Each w As ProduccionDS.FacturasRow In DS.Facturas.Rows
            If r.Tipar = "P" Then
                IntereSEG = w.IntSe
                Capital = (w.RenPr - w.IntPr) + w.RenSe + w.ImporteFEGA + w.SeguroVida + w.IntPr
                Interes = 0
            Else
                IntereSEG = 0
                Capital = (w.RenPr - w.IntPr) + w.RenSe + w.ImporteFEGA + w.SeguroVida
                Interes = w.IntPr + w.IntSe
            End If

            CapitalOt = w.CapitalOt
            InteresOt = w.InteresOt
            IvaCpital = w.IvaCapital + w.IvaPr + w.IvaSe + w.IvaOt
            Pagado = w.ImporteFac - w.SaldoFac

            If w.SaldoFac <> w.ImporteFac Then
                AplicaSaldos(AplicaPagoAnterior(AplicaPrelacion(w, r.Tipar), Pagado), Capital, CapitalOt, Interes, InteresOt, IvaCpital)
            End If


            RR.IvaCapital += IvaCpital
            RR.CapitalVencido += Capital
            RR.InteresVencido += Interes
            RR.CapitalVencidoOt += CapitalOt
            RR.InteresVencidoOt += InteresOt
            If IntereSEG > 0 Then RR.CargaFinancieraSEG += IntereSEG
        Next
        If RR.Regreso = False Then
            Vencido = True
        Else
            Vencido = False
        End If


    End Sub

    Sub TraspasaAVCC(ByRef RR As ProduccionDS.TraspasosVencidosRow, ByRef r As ProduccionDS.CarteraVencidaDETRow, dias As Integer)
        RR.Anexo = r.Anexo
        RR.Ciclo = r.Ciclo
        RR.Tipo = r.Tipo
        RR.Fecha = FechaD.AddDays(dias)
        RR.Tipar = r.Tipar
        RR.Segmento_Negocio = r.Segmento_Negocio
        RR.SaldoInsolutoOTR = 0
        RR.CargaFinancieraOTR = 0
        RR.SaldoInsolutoSEG = 0
        RR.CargaFinancieraSEG = 0
        RR.SaldoInsoluto = 0
        RR.CargaFinanciera = 0
        RR.CapitalVencido = 0
        RR.InteresVencido = 0
        RR.CapitalVencidoOt = 0
        RR.InteresVencidoOt = 0
        RR.IvaCapital = 0
        RR.Regreso = r.Regreso

        Dim TaAV As New ProduccionDSTableAdapters.SaldosAvioTableAdapter
        TaAV.Fill(DS.SaldosAvio, r.Anexo, r.Ciclo)
        Dim Capital, Interes As Decimal

        For Each w As ProduccionDS.SaldosAvioRow In DS.SaldosAvio.Rows
            Capital = w.Imp + w.Fega + w.Garantia
            Interes = w.Intereses
            RR.CapitalVencido += Capital
            RR.InteresVencido += Interes
        Next
        Vencido = True
    End Sub

    Sub CorreTraspasos(Vencida As Boolean)
        If Vencida = True Then
            Console.WriteLine("Cartera Vencida ...")
            TraspasoCarteraVencida(FechaD)
        Else
            Console.WriteLine("Procesando Avio Paso 1 ...")
            Calcula_Saldos(FechaS, "H")
            Console.WriteLine("Procesando Avio Paso 2 ...")
            TraspasosAvio(FechaS, "H")
            Console.WriteLine("Procesando CC Paso 1...")
            Calcula_Saldos(FechaS, "C")
            Console.WriteLine("Procesando CC Paso 2...")
            TraspasosAvio(FechaS, "C")
        End If
    End Sub

End Module
