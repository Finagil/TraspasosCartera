Imports System.Net.Mail
Imports System.IO

Module Traspasos
    Dim Tasa As Decimal
    Dim Contador As Integer = 0
    Dim AplicaGarantiaLIQ As String
    Dim FechaS As String = Date.Now.ToString("yyyyMMdd")
    Dim FechaD As DateTime = Date.Now.Date
    Dim DS As New ProduccionDS
    Dim Vencido As Boolean = False
    Dim TaTrasp As New ProduccionDSTableAdapters.TraspasosVencidosTableAdapter

    Sub Main()
        Console.WriteLine("Iniciando ...")
        Console.WriteLine("Cartera Vencida ...")
        TraspasoCarteraVencida(FechaD)
        'Fecha = "20150506" 'PARA PRUEBAS
        Console.WriteLine("Procesando Avio Paso 1 ...")
        Calcula_Saldos(FechaS, "H")
        Console.WriteLine("Procesando Avio Paso 2 ...")
        TraspasosAvio(FechaS, "H")
        'Fecha = "20150506" 'PARA PRUEBAS
        Console.WriteLine("Procesando CC Paso 1...")
        Calcula_Saldos(FechaS, "C")
        Console.WriteLine("Procesando CC Paso 2...")
        TraspasosAvio(FechaS, "C")
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
        Dim FechaAPP As DateTime = TaVenc.FechaAplicacion
        Dim RR As ProduccionDS.TraspasosVencidosRow

        FechaAPP = FechaAPP.AddDays(-1)
        TaVenc.Fill(DS.CarteraVencidaDET, FechaAPP)
        For Each r As ProduccionDS.CarteraVencidaDETRow In DS.CarteraVencidaDET.Rows
            Console.WriteLine("Cartera Vencida " & r.AnexoCon)
            Vencido = False
            Select Case r.TipoCredito.Trim
                Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "FULL SERVICE", "ARRENDAMIENTO PURO"
                    If r.Dias >= 30 Then
                        RR = DS.TraspasosVencidos.NewRow
                        Select Case r.TipoCredito.Trim
                            Case "ANTICIPO AVÍO", "CREDITO DE AVÍO"
                                TraspasaAVCC(RR, r)
                            Case "FULL SERVICE", "ARRENDAMIENTO PURO"
                                TraspasaTRA(RR, r)
                        End Select
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "CUENTA CORRIENTE"
                    If r.Dias >= 60 Then
                        RR = DS.TraspasosVencidos.NewRow
                        TraspasaAVCC(RR, r)
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
                Case "ARRENDAMIENTO FINANCIERO", "CREDITO REFACCIONARIO", "CREDITO SIMPLE"
                    If r.Dias >= 90 Then
                        RR = DS.TraspasosVencidos.NewRow
                        TraspasaTRA(RR, r)
                        DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                    End If
            End Select
            If Vencido = True Then
                DS.TraspasosVencidos.GetChanges()
                TaTrasp.Update(DS.TraspasosVencidos)
                Select Case r.TipoCredito.Trim
                    Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "CUENTA CORRIENTE"
                        TaVenc.MarcaVencidaAV(r.Anexo)
                    Case Else
                        TaVenc.MarcaVencidaTRA(r.Anexo)
                End Select
            End If
        Next
        ' POR REESTRUCTURAS
        DS.TraspasosVencidos.Clear()
        TaVenc.FillByREEST(DS.CarteraVencidaDET, FechaAPP)
        For Each r As ProduccionDS.CarteraVencidaDETRow In DS.CarteraVencidaDET.Rows
            Console.WriteLine("Cartera Vencida " & r.AnexoCon)
            Vencido = False
            Select Case r.TipoCredito.Trim
                Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "FULL SERVICE", "ARRENDAMIENTO PURO"
                    RR = DS.TraspasosVencidos.NewRow
                    Select Case r.TipoCredito.Trim
                        Case "ANTICIPO AVÍO", "CREDITO DE AVÍO"
                            TraspasaAVCC(RR, r)
                        Case "FULL SERVICE", "ARRENDAMIENTO PURO"
                            TraspasaTRA(RR, r)
                    End Select
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                Case "CUENTA CORRIENTE"
                    RR = DS.TraspasosVencidos.NewRow
                    TraspasaAVCC(RR, r)
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
                Case "ARRENDAMIENTO FINANCIERO", "CREDITO REFACCIONARIO", "CREDITO SIMPLE"
                    RR = DS.TraspasosVencidos.NewRow
                    TraspasaTRA(RR, r)
                    DS.TraspasosVencidos.AddTraspasosVencidosRow(RR)
            End Select
            If Vencido = True Then
                DS.TraspasosVencidos.GetChanges()
                TaTrasp.Update(DS.TraspasosVencidos)
                Select Case r.TipoCredito.Trim
                    Case "ANTICIPO AVÍO", "CREDITO DE AVÍO", "CUENTA CORRIENTE"
                        TaVenc.MarcaVencidaAV(r.Anexo)
                    Case Else
                        TaVenc.MarcaVencidaTRA(r.Anexo)
                End Select
            End If
        Next

    End Sub

    Sub TraspasaTRA(ByRef RR As ProduccionDS.TraspasosVencidosRow, ByRef r As ProduccionDS.CarteraVencidaDETRow)
        TaTrasp.DeleteAnexo(r.Anexo, False)
        Dim EdoV As New ProduccionDSTableAdapters.EdoctavTableAdapter
        Dim EdoS As New ProduccionDSTableAdapters.EdoctasTableAdapter
        Dim EdoO As New ProduccionDSTableAdapters.EdoctaoTableAdapter
        EdoV.Fill(DS.Edoctav, r.Anexo)
        EdoS.Fill(DS.Edoctas, r.Anexo)
        EdoO.Fill(DS.Edoctao, r.Anexo)

        RR.Anexo = r.Anexo
        RR.Ciclo = ""
        RR.Tipo = r.Tipo
        RR.Fecha = FechaD
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
        Vencido = True
    End Sub

    Sub TraspasaAVCC(ByRef RR As ProduccionDS.TraspasosVencidosRow, ByRef r As ProduccionDS.CarteraVencidaDETRow)
        RR.Anexo = r.Anexo
        RR.Ciclo = ""
        RR.Tipo = r.Tipo
        RR.Fecha = FechaD
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
        TaAV.Fill(DS.SaldosAvio, r.Anexo)
        Dim Capital, Interes As Decimal

        For Each w As ProduccionDS.SaldosAvioRow In DS.SaldosAvio.Rows
            Capital = w.Imp + w.Fega + w.Garantia
            Interes = w.Intereses
            RR.CapitalVencido += Capital
            RR.InteresVencido += Interes
        Next
        Vencido = True
    End Sub

End Module
