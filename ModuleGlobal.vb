Module ModuleGlobal
    Private Structure Conceptos
        Public Concepto As String
        Public Importe As Decimal
        Public Porcentaje As Decimal
    End Structure

    Function AplicaPrelacion(ByRef drFactura As ProduccionDS.FacturasRow, cTipar As String) As ArrayList
        Dim aConceptos As New ArrayList()
        Dim aConcepto As New Conceptos()
        Dim NBaseFega, nIvaFEGA, nCapitalEquipo, nInteres, nIvaInteres, nTasaIVA, nInteresOtros As Decimal

        nTasaIVA = Math.Round(drFactura("TasaIVA") / 100, 2)
        NBaseFega = Math.Round(drFactura.ImporteFEGA / (1 + nTasaIVA), 2)
        nIvaFEGA = Math.Round(drFactura.ImporteFEGA - NBaseFega, 2)
        nInteresOtros = drFactura("InteresOt") + drFactura("VarOt")
        nIvaInteres = drFactura("IvaPr") + drFactura("IvaSe")
        nInteres = drFactura("IntPr") + drFactura("VarPr") + drFactura("IntSe") + drFactura("VarSe")
        nCapitalEquipo = drFactura("RenPr") - drFactura("IntPr")

        If drFactura.ImporteFEGA > 0 Then
            aConcepto.Concepto = "FEGA"
            aConcepto.Importe = NBaseFega
            aConcepto.Porcentaje = NBaseFega / (NBaseFega + nIvaFEGA)
            aConceptos.Add(aConcepto)

            aConcepto.Concepto = "IVA FEGA"
            aConcepto.Importe = nIvaFEGA
            aConcepto.Porcentaje = 1
            aConceptos.Add(aConcepto)
        End If

        If drFactura.SeguroVida > 0 Then
            aConcepto.Concepto = "SEGURO DE VIDA"
            aConcepto.Importe = drFactura.SeguroVida
            aConcepto.Porcentaje = 1
            aConceptos.Add(aConcepto)
        End If

        If drFactura.InteresOt > 0 Then
            aConcepto.Concepto = "INTERES OTROS ADEUDOS"
            aConcepto.Importe = nInteresOtros
            aConcepto.Porcentaje = nInteresOtros / (nInteresOtros + drFactura.IvaOt)
            aConceptos.Add(aConcepto)

            If drFactura.IvaOt > 0 Then
                aConcepto.Concepto = "IVA INTERES OTROS ADEUDOS"
                aConcepto.Importe = drFactura.IvaOt
                aConcepto.Porcentaje = 1
                aConceptos.Add(aConcepto)
            End If
        End If

        If drFactura.CapitalOt > 0 Then
            aConcepto.Concepto = "CAPITAL OTROS ADEUDOS"
            aConcepto.Importe = drFactura.CapitalOt
            aConcepto.Porcentaje = 1
            aConceptos.Add(aConcepto)
        End If

        If cTipar = "P" Then

            If drFactura("IntSe") + drFactura("VarSe") > 0 Then
                aConcepto.Concepto = "INTERES SEGURO"
                aConcepto.Importe = drFactura("IntSe") + drFactura("VarSe")
                aConcepto.Porcentaje = (drFactura("IntSe") + drFactura("VarSe")) / (drFactura("IntSe") + drFactura("VarSe") + drFactura("IvaSe"))
                aConceptos.Add(aConcepto)

                If drFactura("IvaSe") > 0 Then
                    aConcepto.Concepto = "IVA INTERES SEGURO"
                    aConcepto.Importe = drFactura("IvaSe")
                    aConcepto.Porcentaje = 1
                    aConceptos.Add(aConcepto)
                End If
            End If

            If drFactura("Rense") > 0 Then
                aConcepto.Concepto = "CAPITAL SEGURO"
                aConcepto.Importe = drFactura("Rense")
                aConcepto.Porcentaje = 1
                aConceptos.Add(aConcepto)
            End If

            If drFactura("RenPr") + drFactura("VarPr") > 0 Then
                aConcepto.Concepto = "PAGO DE RENTA"
                aConcepto.Importe = drFactura("RenPr") + drFactura("VarPr")
                aConcepto.Porcentaje = (drFactura("RenPr") + drFactura("VarPr")) / (drFactura("RenPr") + drFactura("VarPr") + drFactura("IvaCapital") + drFactura("IvaPr"))
                aConceptos.Add(aConcepto)

                aConcepto.Concepto = "IVA DEL PAGO DE RENTA"
                aConcepto.Importe = drFactura("IvaCapital") + drFactura("IvaPr")
                aConcepto.Porcentaje = 1
                aConceptos.Add(aConcepto)
            End If

        ElseIf cTipar = "B" Then


            If nCapitalEquipo > 0 Then
                aConcepto.Concepto = "MENSUALIDAD"
                aConcepto.Importe = nCapitalEquipo
                aConcepto.Porcentaje = nCapitalEquipo / (nCapitalEquipo + drFactura.IvaCapital)
                aConceptos.Add(aConcepto)
                If drFactura.IvaCapital > 0 Then
                    aConcepto.Concepto = "IVA MENSUALIDAD"
                    aConcepto.Importe = drFactura.IvaCapital
                    aConcepto.Porcentaje = 1
                    aConceptos.Add(aConcepto)
                End If
            End If
        Else
            If nInteres > 0 Then
                aConcepto.Concepto = "INTERESES"
                aConcepto.Importe = nInteres
                aConcepto.Porcentaje = nInteres / (nInteres + 0)
                aConceptos.Add(aConcepto)

                If nIvaInteres > 0 Then                                    ' Puede darse el caso en que haya Intereses pero no haya IVA de los intereses, por ejemplo
                    aConcepto.Concepto = "IVA INTERESES"                   ' en un Crédito Refaccionario o Crédito Simple a Persona Moral o Persona Física con Actividad Empresarial
                    aConcepto.Importe = nIvaInteres
                    aConcepto.Porcentaje = 1
                    aConceptos.Add(aConcepto)
                End If
            End If

            If drFactura.RenSe > 0 Then
                aConcepto.Concepto = "CAPITAL SEGURO"
                aConcepto.Importe = drFactura.RenSe
                aConcepto.Porcentaje = 1
                aConceptos.Add(aConcepto)
            End If

            If nCapitalEquipo > 0 Then
                aConcepto.Concepto = "CAPITAL EQUIPO"
                aConcepto.Importe = nCapitalEquipo
                aConcepto.Porcentaje = nCapitalEquipo / (nCapitalEquipo + drFactura.IvaCapital)
                aConceptos.Add(aConcepto)

                If drFactura.IvaCapital > 0 Then                                    ' Puede darse el caso en que haya Capital Equipo pero no haya IVA del Capital
                    aConcepto.Concepto = "IVA CAPITAL"                     ' ya que éste solamente existe para Arrendamiento Financiero
                    aConcepto.Importe = drFactura.IvaCapital
                    aConcepto.Porcentaje = 1
                    aConceptos.Add(aConcepto)

                    If drFactura.Bonifica > 0 Then                                  ' Solamente puede haber bonificación cuando existe IVA del Capital
                        aConcepto.Concepto = "BONIFICACION"
                        aConcepto.Importe = -drFactura.Bonifica
                        aConcepto.Porcentaje = 1
                        aConceptos.Add(aConcepto)
                    End If
                End If
            End If

        End If
        Return (aConceptos)
    End Function

    Function AplicaPagoAnterior(ByVal aConceptos As ArrayList, nPAGADO As Decimal) As ArrayList '#ECT new
        ' Lo primero que tenemos que determinar es cuánto había pagado el cliente y cuáles conceptos se cubrieron con dicho pago
        Dim concep As String
        Dim ImporteT As Decimal

        AplicaPagoAnterior = New ArrayList()
        For Each aConceptoX As Conceptos In aConceptos
            concep = aConceptoX.Concepto
            If nPAGADO > 0 Then
                If nPAGADO >= aConceptoX.Importe / aConceptoX.Porcentaje Then
                    'Pago completo del importe
                    ImporteT = Math.Round(aConceptoX.Importe, 2)
                    nPAGADO = nPAGADO - ImporteT
                    aConceptoX.Importe = 0
                Else
                    ' Pago parcial del importe
                    ImporteT = Math.Round(nPAGADO * aConceptoX.Porcentaje, 2)
                    aConceptoX.Importe = aConceptoX.Importe - ImporteT
                    nPAGADO = nPAGADO - ImporteT
                End If
            End If
            AplicaPagoAnterior.Add(aConceptoX)
        Next
        Return (AplicaPagoAnterior)
    End Function '#ECT new

    Sub AplicaSaldos(ByRef aConceptos As ArrayList, ByRef Capital As Decimal, ByRef CapitalOt As Decimal, ByRef Interes As Decimal, ByRef InteresOt As Decimal, ByRef IVA As Decimal)
        Capital = 0
        CapitalOt = 0
        Interes = 0
        InteresOt = 0
        IVA = 0
        For Each aConcepto In aConceptos
            Select Case aConcepto.Concepto
                Case "FEGA", "SEGURO DE VIDA", "CAPITAL SEGURO", "PAGO DE RENTA", "MENSUALIDAD", "CAPITAL EQUIPO"
                    Capital += aConcepto.importe
                Case "IVA FEGA", "IVA INTERES OTROS ADEUDOS", "IVA INTERES SEGURO", "IVA DEL PAGO DE RENTA", "IVA INTERESES", "IVA CAPITAL", "IVA MENSUALIDAD"
                    IVA += aConcepto.importe
                Case "CAPITAL OTROS ADEUDOS"
                    CapitalOt += aConcepto.importe
                Case "INTERES SEGURO", "INTERESES"
                    Interes += aConcepto.importe
                Case "INTERES OTROS ADEUDOS"
                    InteresOt += aConcepto.importe
                Case "BONIFICACION"
            End Select
        Next
    End Sub

End Module
