Option Explicit

'=========================================================
' MACRO PRINCIPAL V12
' INFORME_QA_P2P + BASE_DATOS COMO VALORES CON FILTROS
' V12 respecto a V11:
'   - Cabecera: nueva "Duración de la Prueba" (Fecha Inicio -> Fecha Fin)
'   - RESUMEN EJECUTIVO: tarjetas KPI + tabla de estados con barra de
'     datos + gráfico de columnas, con anchos balanceados para simetría
'   - HALLAZGOS: se elimina la recolección de fecha/hora que ya no se
'     mostraba (código muerto) y se uniforman las alturas de los banners
'   - DETALLE P2P: se oculta la columna Hora, se filtra por el rango
'     Fecha Inicio de Prueba / Fecha Fin de Prueba (además del filtro
'     por estado), banner ajustado al ancho real de la tabla, bandas
'     alternadas y autofiltro
'   - BASE_DATOS se mantiene sin cambios (copia íntegra de la plantilla
'     visible, como valores)
'=========================================================
Public Sub GenerarInformeMixtoQAP2P_Afinia_V12()

    On Error GoTo errorHandler

    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook

    Dim wsOrigen As Worksheet
    Dim wsInforme As Worksheet
    Dim wsBaseDatos As Worksheet

    Dim ultimaFila As Long
    Dim ultimaFilaInforme As Long

    Dim fila As Long
    Dim filaInf As Long
    Dim filaDetalle As Long
    Dim filaEstado As Long
    Dim filaInicioDetalleDatos As Long
    Dim r As Long

    Dim totalSenales As Long
    Dim totalOK As Long
    Dim totalNO As Long
    Dim totalFAIL As Long
    Dim totalADD As Long
    Dim totalDEL As Long
    Dim totalPENDMarcado As Long
    Dim totalNA As Long
    Dim totalSinPrueba As Long
    Dim totalPendientes As Long
    Dim totalProbadas As Long

    Dim avance As Double
    Dim estadoGeneral As String
    Dim estado As String
    Dim obs As String

    Dim dictHallazgos As Object
    Dim claveObs As String
    Dim textoObs As String
    Dim textoSenal As String
    Dim datos As Variant
    Dim key As Variant
    Dim indiceHallazgo As Long

    Dim fechaInicioPrueba As String
    Dim fechaFinPrueba As String
    Dim personalSitio As String
    Dim estadoPrueba As String
    Dim psmDescCons As String
    Dim duracionPrueba As String

    Dim dFechaInicio As Variant
    Dim dFechaFin As Variant
    Dim rangoFechasValido As Boolean

    Dim rutaGuardar As String
    Dim nombreArchivo As String
    Dim carpetaGuardar As String

    Dim filaInfo As Long
    Dim filaResultadoGeneral As Long
    Dim filaResumenTitulo As Long
    Dim filaKPILabel As Long
    Dim filaKPIValor As Long
    Dim filaEstadoHeader As Long
    Dim rngCantidadEstados As Range
    Dim chtP2P As ChartObject
    Dim filaChartBottom As Long

    Dim colsDetalle As Variant
    Dim idxCol As Long
    Dim destCol As Long
    Dim colPruebaDest As Long
    Dim colFechaDest As Long
    Dim filaIndexDatos As Long
    Dim incluirPorFecha As Boolean
    Dim fechaFilaVar As Variant

    '=====================================================
    ' CONFIGURACIÓN GENERAL
    '=====================================================
    Const HOJA_ORIGEN As String = "Plantilla"

    Const FILA_ENCABEZADO_TABLA As Long = 11
    Const FILA_INICIO_DATOS As Long = 12

    Const COL_SUBESTACION As Long = 1   ' A
    Const COL_NIVEL As Long = 2         ' B
    Const COL_BAHIA As Long = 3         ' C
    Const COL_SENAL As Long = 4         ' D
    Const COL_C1 As Long = 5            ' E
    Const COL_DESCRIPCION As Long = 6   ' F
    Const COL_IEC As Long = 7           ' G
    Const COL_COMANDO As Long = 8       ' H
    Const COL_PRUEBA As Long = 9        ' I
    Const COL_FECHA As Long = 10        ' J
    Const COL_HORA As Long = 11         ' K
    Const COL_OBS As Long = 12          ' L

    Const GUARDAR_ARCHIVO As Boolean = False

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wbOrigen = ThisWorkbook
    Set wsOrigen = wbOrigen.Worksheets(HOJA_ORIGEN)

    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, COL_SUBESTACION).End(xlUp).Row

    If ultimaFila < FILA_INICIO_DATOS Then
        GoTo salidaSegura
    End If

    '=====================================================
    ' LEER DATOS DE CABECERA DESDE PLANTILLA
    '=====================================================
    fechaInicioPrueba = QAP2P12_BuscarValorPorEtiqueta(wsOrigen, "Fecha Inicio de Prueba")
    fechaFinPrueba = QAP2P12_BuscarValorPorEtiqueta(wsOrigen, "Fecha Fin de Prueba")
    personalSitio = QAP2P12_BuscarValorPorEtiqueta(wsOrigen, "Personal en sitio")
    estadoPrueba = QAP2P12_BuscarValorPorEtiqueta(wsOrigen, "Estado")
    psmDescCons = QAP2P12_BuscarValorPorEtiqueta(wsOrigen, "PSM/DESC/CONS")

    dFechaInicio = QAP2P12_ParsearFecha(fechaInicioPrueba)
    dFechaFin = QAP2P12_ParsearFecha(fechaFinPrueba)
    rangoFechasValido = (Not IsEmpty(dFechaInicio)) And (Not IsEmpty(dFechaFin))

    duracionPrueba = QAP2P12_CalcularDuracion(dFechaInicio, dFechaFin)

    '=====================================================
    ' CREAR LIBRO NUEVO
    '=====================================================
    Set wbNuevo = Workbooks.Add

    Do While wbNuevo.Worksheets.Count > 1
        wbNuevo.Worksheets(wbNuevo.Worksheets.Count).Delete
    Loop

    Set wsInforme = wbNuevo.Worksheets(1)
    wsInforme.Name = "INFORME_QA_P2P"

    Set wsBaseDatos = wbNuevo.Worksheets.Add(After:=wbNuevo.Worksheets(wbNuevo.Worksheets.Count))
    wsBaseDatos.Name = "BASE_DATOS"

    '=====================================================
    ' CREAR BASE_DATOS COMO COPIA DE PLANTILLA FILTRADA (SIN CAMBIOS)
    '=====================================================
    Call QAP2P12_CopiarPlantillaVisibleComoValores(wsOrigen, wsBaseDatos)

    '=====================================================
    ' CONTAR KPIS SOBRE FILAS VISIBLES
    '=====================================================
    For fila = FILA_INICIO_DATOS To ultimaFila

        If wsOrigen.Rows(fila).Hidden = False Then

            If Trim(CStr(wsOrigen.Cells(fila, COL_SUBESTACION).Value)) <> "" Then

                totalSenales = totalSenales + 1
                estado = QAP2P12_LimpiarEstado(wsOrigen.Cells(fila, COL_PRUEBA).Value)

                Select Case estado

                    Case ""
                        totalSinPrueba = totalSinPrueba + 1

                    Case "ok"
                        totalOK = totalOK + 1

                    Case "no", "nok"
                        totalNO = totalNO + 1

                    Case "fail"
                        totalFAIL = totalFAIL + 1

                    Case "add"
                        totalADD = totalADD + 1

                    Case "del"
                        totalDEL = totalDEL + 1

                    Case "pend"
                        totalPENDMarcado = totalPENDMarcado + 1

                    Case "n/a", "na"
                        totalNA = totalNA + 1

                End Select

            End If

        End If

    Next fila

    totalPendientes = totalSinPrueba + totalPENDMarcado
    totalProbadas = totalOK + totalNO + totalFAIL + totalADD + totalDEL + totalNA

    If totalSenales > 0 Then
        avance = totalProbadas / totalSenales
    Else
        avance = 0
    End If

    '=====================================================
    ' RESULTADO GENERAL
    '=====================================================
    If totalFAIL > 0 Then

        estadoGeneral = "RECHAZADO"

    ElseIf totalNO > 0 Or totalADD > 0 Or totalDEL > 0 Then

        estadoGeneral = "APROBADO CON OBSERVACIONES"

    ElseIf totalPendientes > 0 Then

        estadoGeneral = "EN PROCESO / PENDIENTE"

    Else

        estadoGeneral = "APROBADO"

    End If

    '=====================================================
    ' AGRUPAR HALLAZGOS POR OBS
    '=====================================================
    Set dictHallazgos = CreateObject("Scripting.Dictionary")

    For fila = FILA_INICIO_DATOS To ultimaFila

        If wsOrigen.Rows(fila).Hidden = False Then

            estado = QAP2P12_LimpiarEstado(wsOrigen.Cells(fila, COL_PRUEBA).Value)
            obs = Trim(CStr(wsOrigen.Cells(fila, COL_OBS).Value))

            If obs <> "" _
            Or estado = "no" _
            Or estado = "nok" _
            Or estado = "fail" _
            Or estado = "add" _
            Or estado = "del" _
            Or estado = "pend" Then

                If obs <> "" Then
                    textoObs = obs
                    claveObs = QAP2P12_NormalizarTexto(obs)
                Else
                    textoObs = "Sin observación registrada - Estado " & QAP2P12_EstadoMostrar(estado)
                    claveObs = QAP2P12_NormalizarTexto(textoObs)
                End If

                If Not dictHallazgos.Exists(claveObs) Then

                    ' 0 = OBS original
                    ' 1 = Cantidad
                    ' 2 = Estados
                    ' 3 = Subestaciones
                    ' 4 = Niveles
                    ' 5 = Bahías
                    ' 6 = Señales
                    ' 7 = Descripciones
                    dictHallazgos.Add claveObs, Array(textoObs, 0, "", "", "", "", "", "")

                End If

                datos = dictHallazgos(claveObs)

                datos(1) = CLng(datos(1)) + 1

                datos(2) = QAP2P12_AgregarUnicoTexto(CStr(datos(2)), QAP2P12_EstadoMostrar(estado), ", ")

                datos(3) = QAP2P12_AgregarUnicoTexto( _
                            CStr(datos(3)), _
                            Trim(CStr(wsOrigen.Cells(fila, COL_SUBESTACION).Value)), _
                            ", ")

                datos(4) = QAP2P12_AgregarUnicoTexto( _
                            CStr(datos(4)), _
                            Trim(CStr(wsOrigen.Cells(fila, COL_NIVEL).Value)), _
                            ", ")

                datos(5) = QAP2P12_AgregarUnicoTexto( _
                            CStr(datos(5)), _
                            Trim(CStr(wsOrigen.Cells(fila, COL_BAHIA).Value)), _
                            ", ")

                textoSenal = Trim(CStr(wsOrigen.Cells(fila, COL_SENAL).Value))

                If textoSenal = "" Then
                    textoSenal = "Sin nombre"
                End If

                textoSenal = textoSenal & " [" & QAP2P12_EstadoMostrar(estado) & "]"

                datos(6) = QAP2P12_AgregarLineaConLimite(CStr(datos(6)), textoSenal, 6)

                If Trim(CStr(wsOrigen.Cells(fila, COL_DESCRIPCION).Value)) <> "" Then
                    datos(7) = QAP2P12_AgregarLineaConLimite( _
                                CStr(datos(7)), _
                                Trim(CStr(wsOrigen.Cells(fila, COL_DESCRIPCION).Value)), _
                                6)
                End If

                dictHallazgos(claveObs) = datos

            End If

        End If

    Next fila

    '=====================================================
    ' INFORME - ENCABEZADO
    '=====================================================
    With wsInforme

        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10

        .Range("A1:L1").Merge
        .Range("A1").Value = "INFORME DE PRUEBAS PUNTO A PUNTO - P2P SCADA"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 15
        .Range("A1").Font.Color = RGB(255, 255, 255)
        .Range("A1").Interior.Color = RGB(0, 119, 200)
        .Range("A1").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 26

        ' Logo robusto o recuadro
        If QAP2P12_CopiarLogoDesdePlantilla(wsOrigen, wsInforme, "LogoAfinia", "K3") = False Then

            .Range("K3:L6").Merge
            .Range("K3").Value = "AFINIA"
            .Range("K3").HorizontalAlignment = xlCenter
            .Range("K3").VerticalAlignment = xlCenter
            .Range("K3").Font.Bold = True
            .Range("K3").Font.Size = 16
            .Range("K3").Font.Color = RGB(0, 119, 200)
            .Range("K3").Interior.Color = RGB(221, 235, 247)
            .Range("K3").Borders.LineStyle = xlContinuous

        End If

        ' Datos generales (bloque dinámico: incluye Duración de la Prueba)
        filaInfo = 3

        .Cells(filaInfo, 1).Value = "Fecha Inicio de Prueba:"
        .Cells(filaInfo, 2).Value = fechaInicioPrueba
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "Fecha Fin de Prueba:"
        .Cells(filaInfo, 2).Value = fechaFinPrueba
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "Duración de la Prueba:"
        .Cells(filaInfo, 2).Value = duracionPrueba
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "Personal en sitio:"
        .Cells(filaInfo, 2).Value = personalSitio
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "Estado:"
        .Cells(filaInfo, 2).Value = estadoPrueba
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "PSM/DESC/CONS:"
        .Cells(filaInfo, 2).Value = psmDescCons
        filaInfo = filaInfo + 1

        .Cells(filaInfo, 1).Value = "Ingeniero SCADA:"
        .Cells(filaInfo, 2).Value = QAP2P12_ObtenerIngenieroSCADA()
        filaInfo = filaInfo + 1

        filaResultadoGeneral = filaInfo
        .Cells(filaInfo, 1).Value = "Resultado General:"
        .Cells(filaInfo, 2).Value = estadoGeneral

        .Range(.Cells(3, 1), .Cells(filaResultadoGeneral, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(filaResultadoGeneral, 2)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(filaResultadoGeneral, 2)).VerticalAlignment = xlCenter
        .Cells(filaResultadoGeneral, 2).Font.Bold = True

        Select Case estadoGeneral

            Case "APROBADO"
                .Cells(filaResultadoGeneral, 2).Interior.Color = RGB(198, 239, 206)
                .Cells(filaResultadoGeneral, 2).Font.Color = RGB(0, 97, 0)

            Case "APROBADO CON OBSERVACIONES"
                .Cells(filaResultadoGeneral, 2).Interior.Color = RGB(255, 235, 156)
                .Cells(filaResultadoGeneral, 2).Font.Color = RGB(156, 101, 0)

            Case "RECHAZADO"
                .Cells(filaResultadoGeneral, 2).Interior.Color = RGB(192, 0, 0)
                .Cells(filaResultadoGeneral, 2).Font.Color = RGB(255, 255, 255)

            Case Else
                .Cells(filaResultadoGeneral, 2).Interior.Color = RGB(217, 217, 217)
                .Cells(filaResultadoGeneral, 2).Font.Color = RGB(0, 0, 0)

        End Select

    End With

    '=====================================================
    ' RESUMEN EJECUTIVO - TARJETAS KPI + TABLA DE ESTADOS + GRÁFICO
    '=====================================================
    filaResumenTitulo = filaResultadoGeneral + 2

    With wsInforme

        .Range("A" & filaResumenTitulo & ":L" & filaResumenTitulo).Merge
        .Range("A" & filaResumenTitulo).Value = "RESUMEN EJECUTIVO"
        .Range("A" & filaResumenTitulo).Font.Bold = True
        .Range("A" & filaResumenTitulo).Font.Color = RGB(255, 255, 255)
        .Range("A" & filaResumenTitulo).Interior.Color = RGB(0, 119, 200)
        .Range("A" & filaResumenTitulo).HorizontalAlignment = xlCenter
        .Range("A" & filaResumenTitulo).VerticalAlignment = xlCenter
        .Rows(filaResumenTitulo).RowHeight = 22

    End With

    ' --- Tarjetas KPI (anchos balanceados según ancho real de columnas A:L) ---
    filaKPILabel = filaResumenTitulo + 2
    filaKPIValor = filaKPILabel + 1

    Call QAP2P12_EscribirKPICard(wsInforme, filaKPILabel, filaKPIValor, 1, 5, _
        "Total Señales", totalSenales, "0")

    Call QAP2P12_EscribirKPICard(wsInforme, filaKPILabel, filaKPIValor, 6, 7, _
        "Señales Probadas", totalProbadas, "0", RGB(198, 239, 206), RGB(0, 97, 0))

    Call QAP2P12_EscribirKPICard(wsInforme, filaKPILabel, filaKPIValor, 8, 9, _
        "Pendientes", totalPendientes, "0", RGB(255, 235, 156), RGB(156, 101, 0))

    Call QAP2P12_EscribirKPICard(wsInforme, filaKPILabel, filaKPIValor, 10, 12, _
        "% Avance", avance, "0.0%")

    ' --- Tabla de estados con barra de datos + gráfico de distribución ---
    filaEstadoHeader = filaKPIValor + 2

    With wsInforme

        .Cells(filaEstadoHeader, 4).Value = "Estado"
        .Cells(filaEstadoHeader, 5).Value = "Cantidad"
        .Cells(filaEstadoHeader, 6).Value = "Significado"

        With .Range(.Cells(filaEstadoHeader, 4), .Cells(filaEstadoHeader, 6))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
        End With

    End With

    filaEstado = filaEstadoHeader + 1

    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "OK", totalOK, "Señal probada correctamente", "ok")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "NOK", totalNO, "Señal probada no conforme", "no")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "FAIL", totalFAIL, "Falla detectada durante la prueba", "fail")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "ADD", totalADD, "Señal adicionada durante la validación", "add")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "DEL", totalDEL, "Señal eliminada o no encontrada", "del")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "PEND", totalPendientes, "Pendiente por probar / sin prueba", "pend")
    Call QAP2P12_AgregarEstadoResumen(wsInforme, filaEstado, "N/A", totalNA, "No aplica al alcance de la prueba", "n/a")

    filaChartBottom = 0

    If filaEstado > filaEstadoHeader + 1 Then

        ' Barra de datos en la columna Cantidad
        On Error Resume Next
        Set rngCantidadEstados = wsInforme.Range( _
            wsInforme.Cells(filaEstadoHeader + 1, 5), wsInforme.Cells(filaEstado - 1, 5))
        rngCantidadEstados.FormatConditions.Delete
        With rngCantidadEstados.FormatConditions.AddDatabar
            .BarColor.Color = RGB(0, 119, 200)
            .ShowValue = True
        End With
        On Error GoTo 0

        ' Gráfico de columnas con la distribución de estados
        Set chtP2P = QAP2P12_CrearGraficoEstados(wsInforme, filaEstadoHeader, filaEstado - 1)

        If Not chtP2P Is Nothing Then

            Dim chartBottomPts As Double
            chartBottomPts = chtP2P.Top + chtP2P.Height

            filaChartBottom = filaEstadoHeader

            Do While wsInforme.Rows(filaChartBottom).Top < chartBottomPts _
                    And filaChartBottom < filaEstadoHeader + 300
                filaChartBottom = filaChartBottom + 1
            Loop

        End If

    End If

    '=====================================================
    ' HALLAZGOS / OBSERVACIONES AGRUPADAS DE LA PRUEBA
    '=====================================================
    filaInf = Application.WorksheetFunction.Max(filaEstado + 2, filaChartBottom + 2)

    With wsInforme

        .Range("A" & filaInf & ":I" & filaInf).Merge
        .Range("A" & filaInf).Value = "HALLAZGOS / OBSERVACIONES AGRUPADAS DE LA PRUEBA"
        .Range("A" & filaInf).Font.Bold = True
        .Range("A" & filaInf).Font.Color = RGB(255, 255, 255)
        .Range("A" & filaInf).Interior.Color = RGB(0, 119, 200)
        .Range("A" & filaInf).HorizontalAlignment = xlCenter
        .Range("A" & filaInf).VerticalAlignment = xlCenter
        .Range("A" & filaInf).Borders.LineStyle = xlContinuous
        .Rows(filaInf).RowHeight = 22

        filaInf = filaInf + 2

        .Cells(filaInf, 1).Value = "#"
        .Cells(filaInf, 2).Value = "Estado"
        .Cells(filaInf, 3).Value = "Cant."
        .Cells(filaInf, 4).Value = "Subestación"
        .Cells(filaInf, 5).Value = "Nivel"
        .Cells(filaInf, 6).Value = "Bahía"
        .Cells(filaInf, 7).Value = "Observación"
        .Cells(filaInf, 8).Value = "Señales afectadas"
        .Cells(filaInf, 9).Value = "Descripción"

        With .Range(.Cells(filaInf, 1), .Cells(filaInf, 9))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 119, 200)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        .Rows(filaInf).RowHeight = 24

        filaInf = filaInf + 1

    End With

    If dictHallazgos.Count = 0 Then

        With wsInforme

            .Range("A" & filaInf & ":I" & filaInf).Merge
            .Range("A" & filaInf).Value = "No se registraron hallazgos u observaciones relevantes."
            .Range("A" & filaInf).Interior.Color = RGB(198, 239, 206)
            .Range("A" & filaInf).Font.Color = RGB(0, 97, 0)
            .Range("A" & filaInf).Font.Bold = True
            .Range("A" & filaInf).Borders.LineStyle = xlContinuous
            .Range("A" & filaInf).HorizontalAlignment = xlCenter

        End With

        filaInf = filaInf + 1

    Else

        indiceHallazgo = 1

        For Each key In dictHallazgos.Keys

            datos = dictHallazgos(key)

            With wsInforme

                .Cells(filaInf, 1).Value = indiceHallazgo
                .Cells(filaInf, 2).Value = CStr(datos(2))
                .Cells(filaInf, 3).Value = CLng(datos(1))
                .Cells(filaInf, 4).Value = CStr(datos(3))
                .Cells(filaInf, 5).Value = CStr(datos(4))
                .Cells(filaInf, 6).Value = CStr(datos(5))
                .Cells(filaInf, 7).Value = CStr(datos(0))
                .Cells(filaInf, 8).Value = CStr(datos(6))
                .Cells(filaInf, 9).Value = CStr(datos(7))

                With .Range(.Cells(filaInf, 1), .Cells(filaInf, 9))
                    .Borders.LineStyle = xlContinuous
                    .WrapText = True
                    .VerticalAlignment = xlTop
                End With

                ' Alineaciones
                .Cells(filaInf, 1).HorizontalAlignment = xlCenter
                .Cells(filaInf, 2).HorizontalAlignment = xlCenter
                .Cells(filaInf, 3).HorizontalAlignment = xlCenter
                .Cells(filaInf, 4).HorizontalAlignment = xlCenter
                .Cells(filaInf, 5).HorizontalAlignment = xlCenter
                .Cells(filaInf, 6).HorizontalAlignment = xlCenter
                .Cells(filaInf, 7).HorizontalAlignment = xlLeft
                .Cells(filaInf, 8).HorizontalAlignment = xlLeft
                .Cells(filaInf, 9).HorizontalAlignment = xlLeft

                ' Colores
                .Cells(filaInf, 2).Interior.Color = QAP2P12_ColorPrioritarioHallazgo(CStr(datos(2)))
                .Cells(filaInf, 2).Font.Color = RGB(255, 255, 255)
                .Cells(filaInf, 2).Font.Bold = True

                .Cells(filaInf, 7).Interior.Color = RGB(242, 242, 242)
                .Cells(filaInf, 8).Interior.Color = RGB(250, 250, 250)
                .Cells(filaInf, 9).Interior.Color = RGB(250, 250, 250)

                .Rows(filaInf).AutoFit

                If .Rows(filaInf).RowHeight < 36 Then
                    .Rows(filaInf).RowHeight = 36
                End If

                If .Rows(filaInf).RowHeight > 90 Then
                    .Rows(filaInf).RowHeight = 90
                End If

            End With

            filaInf = filaInf + 1
            indiceHallazgo = indiceHallazgo + 1

        Next key

    End If

    '=====================================================
    ' DETALLE P2P - SOLO SEÑALES PROBADAS
    ' - Sin columna Hora (oculta)
    ' - Filtrado por estado (con etiqueta, distinto de PEND)
    ' - Filtrado por rango Fecha Inicio de Prueba / Fecha Fin de Prueba
    ' - Banner ajustado al ancho real de la tabla (11 columnas -> A:K)
    '=====================================================
    colsDetalle = Array(COL_SUBESTACION, COL_NIVEL, COL_BAHIA, COL_SENAL, COL_C1, _
                         COL_DESCRIPCION, COL_IEC, COL_COMANDO, COL_PRUEBA, COL_FECHA, COL_OBS)

    filaDetalle = filaInf + 2

    With wsInforme

        .Range("A" & filaDetalle & ":K" & filaDetalle).Merge
        .Range("A" & filaDetalle).Value = "DETALLE P2P - SOLO SEÑALES PROBADAS"
        .Range("A" & filaDetalle).Font.Bold = True
        .Range("A" & filaDetalle).Font.Color = RGB(255, 255, 255)
        .Range("A" & filaDetalle).Interior.Color = RGB(0, 119, 200)
        .Range("A" & filaDetalle).HorizontalAlignment = xlCenter
        .Range("A" & filaDetalle).VerticalAlignment = xlCenter
        .Rows(filaDetalle).RowHeight = 22

        filaDetalle = filaDetalle + 2

        For idxCol = LBound(colsDetalle) To UBound(colsDetalle)

            destCol = idxCol - LBound(colsDetalle) + 1

            If colsDetalle(idxCol) = COL_PRUEBA Then colPruebaDest = destCol
            If colsDetalle(idxCol) = COL_FECHA Then colFechaDest = destCol

            .Cells(filaDetalle, destCol).Value = wsOrigen.Cells(FILA_ENCABEZADO_TABLA, colsDetalle(idxCol)).Value
            .Cells(filaDetalle, destCol).Interior.Color = RGB(0, 119, 200)
            .Cells(filaDetalle, destCol).Font.Color = RGB(255, 255, 255)
            .Cells(filaDetalle, destCol).Font.Bold = True
            .Cells(filaDetalle, destCol).HorizontalAlignment = xlCenter
            .Cells(filaDetalle, destCol).Borders.LineStyle = xlContinuous

        Next idxCol

        On Error Resume Next
        .Range(.Cells(filaDetalle, 1), .Cells(filaDetalle, UBound(colsDetalle) - LBound(colsDetalle) + 1)).AutoFilter
        On Error GoTo 0

        filaDetalle = filaDetalle + 1
        filaInicioDetalleDatos = filaDetalle

    End With

    filaIndexDatos = 0

    For fila = FILA_INICIO_DATOS To ultimaFila

        If wsOrigen.Rows(fila).Hidden = False Then

            estado = QAP2P12_LimpiarEstado(wsOrigen.Cells(fila, COL_PRUEBA).Value)

            incluirPorFecha = True

            If rangoFechasValido Then

                fechaFilaVar = wsOrigen.Cells(fila, COL_FECHA).Value

                If IsDate(fechaFilaVar) Then

                    If DateValue(fechaFilaVar) < DateValue(CDate(dFechaInicio)) _
                    Or DateValue(fechaFilaVar) > DateValue(CDate(dFechaFin)) Then
                        incluirPorFecha = False
                    End If

                Else

                    incluirPorFecha = False

                End If

            End If

            If estado <> "" And estado <> "pend" And incluirPorFecha Then

                With wsInforme

                    For idxCol = LBound(colsDetalle) To UBound(colsDetalle)

                        destCol = idxCol - LBound(colsDetalle) + 1

                        If colsDetalle(idxCol) = COL_FECHA Then

                            If Trim(CStr(wsOrigen.Cells(fila, COL_FECHA).Value)) <> "" Then
                                .Cells(filaDetalle, destCol).Value = wsOrigen.Cells(fila, COL_FECHA).Value
                                .Cells(filaDetalle, destCol).NumberFormat = "dd/mm/yyyy"
                            End If

                        Else

                            .Cells(filaDetalle, destCol).Value = wsOrigen.Cells(fila, colsDetalle(idxCol)).Value2

                        End If

                        .Cells(filaDetalle, destCol).Borders.LineStyle = xlContinuous
                        .Cells(filaDetalle, destCol).WrapText = True

                        If filaIndexDatos Mod 2 = 1 And colsDetalle(idxCol) <> COL_PRUEBA Then
                            .Cells(filaDetalle, destCol).Interior.Color = RGB(245, 249, 252)
                        End If

                    Next idxCol

                    Call QAP2P12_ColorearRangoEstado(.Cells(filaDetalle, colPruebaDest), estado)

                End With

                filaDetalle = filaDetalle + 1
                filaIndexDatos = filaIndexDatos + 1

            End If

        End If

    Next fila

    ultimaFilaInforme = filaDetalle - 1

    If ultimaFilaInforme >= filaInicioDetalleDatos Then
        Call QAP2P12_FormatearFechaHora(wsInforme, filaInicioDetalleDatos, ultimaFilaInforme, colFechaDest, 0)
    End If

    '=====================================================
    ' FORMATO FINAL INFORME
    '=====================================================
    With wsInforme

        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 16
        .Columns("G").ColumnWidth = 34
        .Columns("H").ColumnWidth = 30
        .Columns("I").ColumnWidth = 36
        .Columns("J").ColumnWidth = 12
        .Columns("K").ColumnWidth = 18
        .Columns("L").ColumnWidth = 18

        .Columns("A:L").VerticalAlignment = xlTop
        .Columns("A:L").WrapText = True

        .Rows.AutoFit

        For r = 1 To .UsedRange.Rows.Count

            If .Rows(r).RowHeight > 70 Then
                .Rows(r).RowHeight = 70
            End If

        Next r

        ' Las filas de tarjetas KPI y encabezados de sección deben conservar
        ' su altura fija (no dejar que AutoFit las aplaste ni las estire)
        .Rows(filaKPILabel).RowHeight = 16
        .Rows(filaKPIValor).RowHeight = 30

        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.CenterHorizontally = True
        .PageSetup.LeftMargin = Application.InchesToPoints(0.25)
        .PageSetup.RightMargin = Application.InchesToPoints(0.25)
        .PageSetup.TopMargin = Application.InchesToPoints(0.4)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.4)
        .PageSetup.PrintArea = .Range("A1:L" & ultimaFilaInforme).Address

    End With

    '=====================================================
    ' GUARDAR SI ESTÁ ACTIVADO
    '=====================================================
    If GUARDAR_ARCHIVO = True Then

        nombreArchivo = "P2P_QA_SCADA_MIXTO_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"

        If wbOrigen.Path <> "" Then
            carpetaGuardar = wbOrigen.Path
        Else
            carpetaGuardar = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        End If

        rutaGuardar = carpetaGuardar & Application.PathSeparator & nombreArchivo

        wbNuevo.SaveAs Filename:=rutaGuardar, FileFormat:=xlOpenXMLWorkbook

    End If

salidaSegura:

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Exit Sub

errorHandler:

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "ERROR al generar el informe QA P2P" & vbCrLf & _
           Err.Number & vbCrLf & _
           Err.Description, vbCritical

End Sub


'=========================================================
' BASE_DATOS: COPIA LA PLANTILLA FILTRADA COMO VALORES
'=========================================================
Public Sub QAP2P12_CopiarPlantillaVisibleComoValores(ByVal hojaOrigen As Worksheet, _
                                                      ByVal hojaNueva As Worksheet)

    On Error GoTo errorHandler

    Dim rngVisible As Range
    Dim i As Long
    Dim ultimaFilaNueva As Long

    On Error Resume Next
    Set rngVisible = hojaOrigen.UsedRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo errorHandler

    If rngVisible Is Nothing Then
        Err.Raise vbObjectError + 100, , "No hay datos visibles para copiar."
    End If

    rngVisible.Copy

    hojaNueva.Activate
    hojaNueva.Range("A1").Select
    hojaNueva.Paste

    Application.CutCopyMode = False

    hojaNueva.UsedRange.Value = hojaNueva.UsedRange.Value

    For i = 1 To hojaOrigen.UsedRange.Columns.Count
        hojaNueva.Columns(i).ColumnWidth = hojaOrigen.Columns(i).ColumnWidth
    Next i

    ultimaFilaNueva = hojaNueva.Cells(hojaNueva.Rows.Count, 1).End(xlUp).Row

    With hojaNueva

        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 9

        On Error Resume Next
        .Range("A6:L6").AutoFilter
        On Error GoTo 0

        .Columns("J").NumberFormat = "dd/mm/yyyy"
        .Columns("K").NumberFormat = "hh:mm:ss"

        If ultimaFilaNueva >= 12 Then
            Call QAP2P12_FormatearFechaHora(hojaNueva, 12, ultimaFilaNueva, 10, 11)
        End If

        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False

    End With

    Exit Sub

errorHandler:

    MsgBox "ERROR al crear BASE_DATOS" & vbCrLf & _
           Err.Number & vbCrLf & _
           Err.Description, vbCritical

End Sub


'=========================================================
' FORMATEAR FECHA Y HORA COMO VALORES REALES DE EXCEL
' Si colHora = 0, solo formatea la columna de fecha (usado en
' DETALLE P2P, donde la columna Hora ya no se muestra)
'=========================================================
Public Sub QAP2P12_FormatearFechaHora(ByVal ws As Worksheet, _
                                      ByVal filaInicio As Long, _
                                      ByVal filaFin As Long, _
                                      ByVal colFecha As Long, _
                                      ByVal colHora As Long)

    Dim f As Long
    Dim valorFecha As Variant
    Dim valorHora As Variant
    Dim textoHora As String

    If filaFin < filaInicio Then Exit Sub

    For f = filaInicio To filaFin

        If colFecha > 0 Then

            valorFecha = ws.Cells(f, colFecha).Value

            If Trim(CStr(valorFecha)) <> "" Then
                If IsDate(valorFecha) Then
                    ws.Cells(f, colFecha).Value = CDate(valorFecha)
                    ws.Cells(f, colFecha).NumberFormat = "dd/mm/yyyy"
                End If
            End If

        End If

        If colHora > 0 Then

            valorHora = ws.Cells(f, colHora).Value
            textoHora = Trim(CStr(valorHora))

            If textoHora <> "" Then

                If IsNumeric(valorHora) Then
                    ws.Cells(f, colHora).Value = CDbl(valorHora)
                    ws.Cells(f, colHora).NumberFormat = "hh:mm:ss"

                ElseIf IsDate(valorHora) Then
                    ws.Cells(f, colHora).Value = TimeValue(valorHora)
                    ws.Cells(f, colHora).NumberFormat = "hh:mm:ss"
                End If

            End If

        End If

    Next f

End Sub


'=========================================================
' AGREGAR ESTADO AL RESUMEN EJECUTIVO SOLO SI > 0
'=========================================================
Public Sub QAP2P12_AgregarEstadoResumen(ByVal ws As Worksheet, _
                                        ByRef filaEstado As Long, _
                                        ByVal etiquetaEstado As String, _
                                        ByVal cantidad As Long, _
                                        ByVal significado As String, _
                                        ByVal estadoColor As String)

    If cantidad <= 0 Then Exit Sub

    With ws
        .Cells(filaEstado, 4).Value = etiquetaEstado
        .Cells(filaEstado, 5).Value = cantidad
        .Cells(filaEstado, 6).Value = significado

        .Range(.Cells(filaEstado, 4), .Cells(filaEstado, 6)).Borders.LineStyle = xlContinuous
        Call QAP2P12_ColorearRangoEstado(.Range(.Cells(filaEstado, 4), .Cells(filaEstado, 6)), estadoColor)
    End With

    filaEstado = filaEstado + 1

End Sub


'=========================================================
' ESCRIBIR TARJETA KPI (etiqueta + valor grande, con borde)
' colIni/colFin definen el rango de fusión: se eligen de forma
' que las 4 tarjetas del Resumen Ejecutivo queden con anchos
' balanceados según el ancho real de columnas A:L
'=========================================================
Public Sub QAP2P12_EscribirKPICard(ByVal ws As Worksheet, _
                                    ByVal filaLabel As Long, ByVal filaValor As Long, _
                                    ByVal colIni As Long, ByVal colFin As Long, _
                                    ByVal etiqueta As String, ByVal valor As Variant, _
                                    Optional ByVal formatoValor As String = "", _
                                    Optional ByVal colorFondo As Long = -1, _
                                    Optional ByVal colorTexto As Long = -1)

    Dim cFondo As Long
    Dim cTexto As Long

    If colorFondo = -1 Then cFondo = RGB(221, 235, 247) Else cFondo = colorFondo
    If colorTexto = -1 Then cTexto = RGB(0, 51, 102) Else cTexto = colorTexto

    With ws

        .Range(.Cells(filaLabel, colIni), .Cells(filaLabel, colFin)).Merge
        .Cells(filaLabel, colIni).Value = UCase(etiqueta)
        .Cells(filaLabel, colIni).Font.Bold = True
        .Cells(filaLabel, colIni).Font.Size = 9
        .Cells(filaLabel, colIni).Font.Color = RGB(255, 255, 255)
        .Cells(filaLabel, colIni).Interior.Color = RGB(0, 119, 200)
        .Cells(filaLabel, colIni).HorizontalAlignment = xlCenter
        .Cells(filaLabel, colIni).VerticalAlignment = xlCenter

        .Range(.Cells(filaValor, colIni), .Cells(filaValor, colFin)).Merge
        .Cells(filaValor, colIni).Value = valor

        If formatoValor <> "" Then
            .Cells(filaValor, colIni).NumberFormat = formatoValor
        End If

        .Cells(filaValor, colIni).Font.Bold = True
        .Cells(filaValor, colIni).Font.Size = 18
        .Cells(filaValor, colIni).Font.Color = cTexto
        .Cells(filaValor, colIni).Interior.Color = cFondo
        .Cells(filaValor, colIni).HorizontalAlignment = xlCenter
        .Cells(filaValor, colIni).VerticalAlignment = xlCenter

        With .Range(.Cells(filaLabel, colIni), .Cells(filaValor, colFin)).Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With

    End With

    ws.Rows(filaLabel).RowHeight = 16
    ws.Rows(filaValor).RowHeight = 30

End Sub


'=========================================================
' GRÁFICO DE COLUMNAS CON LA DISTRIBUCIÓN DE ESTADOS
' Se ubica a la derecha de la tabla de estados (columnas H:L).
' Devuelve Nothing si no fue posible crearlo (no interrumpe
' la generación del informe).
'=========================================================
Public Function QAP2P12_CrearGraficoEstados(ByVal ws As Worksheet, _
                                             ByVal filaEncabezado As Long, _
                                             ByVal filaFinDatos As Long) As ChartObject

    On Error GoTo salidaError

    Dim chtObj As ChartObject
    Dim rngDatos As Range
    Dim celdaAncla As Range
    Dim altoDisponible As Double
    Dim r As Long

    If filaFinDatos < filaEncabezado + 1 Then
        Set QAP2P12_CrearGraficoEstados = Nothing
        Exit Function
    End If

    Set rngDatos = ws.Range(ws.Cells(filaEncabezado + 1, 4), ws.Cells(filaFinDatos, 5))
    Set celdaAncla = ws.Cells(filaEncabezado, 8)

    altoDisponible = ws.Range(ws.Rows(filaEncabezado), ws.Rows(filaFinDatos)).Height

    If altoDisponible < 140 Then altoDisponible = 140

    Set chtObj = ws.ChartObjects.Add( _
                    Left:=celdaAncla.Left, _
                    Top:=celdaAncla.Top, _
                    Width:=ws.Range(ws.Cells(filaEncabezado, 8), ws.Cells(filaEncabezado, 12)).Width, _
                    Height:=altoDisponible)

    chtObj.Placement = xlMoveAndSize

    With chtObj.Chart

        .SetSourceData Source:=rngDatos
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Distribución de Estados"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Bold = True
        .HasLegend = False

        On Error Resume Next

        For r = 1 To .SeriesCollection(1).Points.Count
            .SeriesCollection(1).Points(r).Format.Fill.ForeColor.RGB = _
                QAP2P12_ColorFondoEstado(QAP2P12_LimpiarEstado(ws.Cells(filaEncabezado + r, 4).Value))
        Next r

        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.Font.Size = 9
        .SeriesCollection(1).DataLabels.Font.Bold = True

        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Size = 9

        On Error GoTo salidaError

    End With

    Set QAP2P12_CrearGraficoEstados = chtObj
    Exit Function

salidaError:

    Set QAP2P12_CrearGraficoEstados = Nothing

End Function


'=========================================================
' BUSCAR VALOR POR ETIQUETA
'=========================================================
Public Function QAP2P12_BuscarValorPorEtiqueta(ByVal ws As Worksheet, _
                                               ByVal etiqueta As String) As String

    Dim celda As Range
    Dim textoCelda As String
    Dim etiquetaNormal As String

    etiquetaNormal = QAP2P12_NormalizarTexto(etiqueta)

    For Each celda In ws.UsedRange.Cells

        textoCelda = QAP2P12_NormalizarTexto(CStr(celda.Value))

        If textoCelda <> "" Then
            If InStr(1, textoCelda, etiquetaNormal, vbTextCompare) > 0 Then
                QAP2P12_BuscarValorPorEtiqueta = Trim(CStr(celda.Offset(0, 1).Text))
                Exit Function
            End If
        End If

    Next celda

    QAP2P12_BuscarValorPorEtiqueta = ""

End Function


'=========================================================
' PARSEAR FECHA DE FORMA SEGURA (Empty si no es una fecha válida)
'=========================================================
Public Function QAP2P12_ParsearFecha(ByVal texto As String) As Variant

    texto = Trim(CStr(texto))

    If texto = "" Then
        QAP2P12_ParsearFecha = Empty
        Exit Function
    End If

    If IsDate(texto) Then
        QAP2P12_ParsearFecha = CDate(texto)
    Else
        QAP2P12_ParsearFecha = Empty
    End If

End Function


'=========================================================
' CALCULAR DURACIÓN DE LA PRUEBA (Fecha Inicio -> Fecha Fin)
'=========================================================
Public Function QAP2P12_CalcularDuracion(ByVal valInicio As Variant, ByVal valFin As Variant) As String

    Dim diasTotales As Double
    Dim dias As Long
    Dim horas As Double

    If IsEmpty(valInicio) Or IsEmpty(valFin) Then
        QAP2P12_CalcularDuracion = "N/D"
        Exit Function
    End If

    diasTotales = CDbl(valFin) - CDbl(valInicio)

    If diasTotales < 0 Then
        QAP2P12_CalcularDuracion = "N/D (rango inválido)"
        Exit Function
    End If

    If diasTotales = 0 Then
        QAP2P12_CalcularDuracion = "Mismo día"
        Exit Function
    End If

    If diasTotales < 1 Then
        horas = diasTotales * 24
        QAP2P12_CalcularDuracion = Format(horas, "0.0") & " horas"
        Exit Function
    End If

    dias = Int(diasTotales)
    horas = (diasTotales - dias) * 24

    If horas >= 0.5 Then
        QAP2P12_CalcularDuracion = dias & " día(s) " & Format(horas, "0") & " h"
    Else
        QAP2P12_CalcularDuracion = dias & " día(s)"
    End If

End Function


'=========================================================
' COPIAR LOGO DESDE PLANTILLA - ROBUSTO
'=========================================================
Public Function QAP2P12_CopiarLogoDesdePlantilla(ByVal wsOrigen As Worksheet, _
                                                 ByVal wsDestino As Worksheet, _
                                                 ByVal nombreShape As String, _
                                                 ByVal celdaDestino As String) As Boolean

    On Error GoTo salidaError

    Dim shpOrigen As Shape
    Dim shpDestino As Shape
    Dim shp As Shape
    Dim imagenEncontrada As Boolean

    imagenEncontrada = False

    On Error Resume Next
    Set shpOrigen = wsOrigen.Shapes(nombreShape)
    On Error GoTo salidaError

    If Not shpOrigen Is Nothing Then
        imagenEncontrada = True
    End If

    If imagenEncontrada = False Then
        For Each shp In wsOrigen.Shapes
            If shp.Type = msoPicture _
            Or shp.Type = msoLinkedPicture _
            Or shp.Type = msoGroup Then
                If shp.TopLeftCell.Row <= 12 Then
                    Set shpOrigen = shp
                    imagenEncontrada = True
                    Exit For
                End If
            End If
        Next shp
    End If

    If imagenEncontrada = False Then
        For Each shp In wsOrigen.Shapes
            If shp.Type = msoPicture _
            Or shp.Type = msoLinkedPicture _
            Or shp.Type = msoGroup Then
                Set shpOrigen = shp
                imagenEncontrada = True
                Exit For
            End If
        Next shp
    End If

    If imagenEncontrada = False Then
        QAP2P12_CopiarLogoDesdePlantilla = False
        Exit Function
    End If

    shpOrigen.Copy

    wsDestino.Activate
    wsDestino.Paste

    Set shpDestino = wsDestino.Shapes(wsDestino.Shapes.Count)

    With shpDestino
        .Top = wsDestino.Range(celdaDestino).Top
        .Left = wsDestino.Range(celdaDestino).Left
        .LockAspectRatio = msoTrue
        .Height = 60

        If .Width > 160 Then
            .Width = 160
        End If

        .Placement = xlMoveAndSize
    End With

    QAP2P12_CopiarLogoDesdePlantilla = True
    Exit Function

salidaError:

    QAP2P12_CopiarLogoDesdePlantilla = False

End Function


'=========================================================
' INGENIERO SCADA SEGÚN USUARIO DE RED
'=========================================================
Public Function QAP2P12_ObtenerIngenieroSCADA() As String

    Dim usuario As String

    usuario = QAP2P12_LimpiarEstado(Environ("USERNAME"))

    Select Case usuario

        Case "jcanchan"
            QAP2P12_ObtenerIngenieroSCADA = "Julio Cesar Canchano Parody (" & usuario & ")"

        Case "jacostapa"
            QAP2P12_ObtenerIngenieroSCADA = "Jonnathan Acosta Pajaro (" & usuario & ")"

        Case "oscarc"
            QAP2P12_ObtenerIngenieroSCADA = "Oscar Javier Carmona Parra (" & usuario & ")"

        Case "gsanchez"
            QAP2P12_ObtenerIngenieroSCADA = "Guido Jose Sanchez Sanchez (" & usuario & ")"

        Case Else
            QAP2P12_ObtenerIngenieroSCADA = usuario

    End Select

End Function


'=========================================================
' LIMPIAR ESTADO
'=========================================================
Public Function QAP2P12_LimpiarEstado(ByVal valor As Variant) As String

    Dim texto As String

    texto = CStr(valor)
    texto = Replace(texto, Chr(160), " ")
    texto = Trim(texto)
    texto = LCase(texto)

    QAP2P12_LimpiarEstado = texto

End Function


'=========================================================
' NORMALIZAR TEXTO
'=========================================================
Public Function QAP2P12_NormalizarTexto(ByVal texto As String) As String

    texto = Trim(CStr(texto))

    texto = Replace(texto, ":", "")
    texto = Replace(texto, Chr(160), " ")
    texto = Replace(texto, vbCrLf, " ")
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    texto = Replace(texto, vbTab, " ")

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    texto = UCase(texto)
    texto = Trim(texto)

    QAP2P12_NormalizarTexto = texto

End Function


'=========================================================
' ESTADO A TEXTO PRESENTABLE
'=========================================================
Public Function QAP2P12_EstadoMostrar(ByVal estado As String) As String

    estado = QAP2P12_LimpiarEstado(estado)

    Select Case estado

        Case "ok"
            QAP2P12_EstadoMostrar = "OK"

        Case "no", "nok"
            QAP2P12_EstadoMostrar = "NOK"

        Case "fail"
            QAP2P12_EstadoMostrar = "FAIL"

        Case "add"
            QAP2P12_EstadoMostrar = "ADD"

        Case "del"
            QAP2P12_EstadoMostrar = "DEL"

        Case "pend"
            QAP2P12_EstadoMostrar = "PEND"

        Case "n/a", "na"
            QAP2P12_EstadoMostrar = "N/A"

        Case Else
            QAP2P12_EstadoMostrar = UCase(estado)

    End Select

End Function


'=========================================================
' AGREGAR TEXTO ÚNICO
'=========================================================
Public Function QAP2P12_AgregarUnicoTexto(ByVal textoActual As String, _
                                          ByVal textoNuevo As String, _
                                          ByVal separador As String) As String

    textoNuevo = Trim(CStr(textoNuevo))

    If textoNuevo = "" Then
        QAP2P12_AgregarUnicoTexto = textoActual
        Exit Function
    End If

    If textoActual = "" Then

        QAP2P12_AgregarUnicoTexto = textoNuevo

    ElseIf InStr(1, separador & textoActual & separador, _
                 separador & textoNuevo & separador, _
                 vbTextCompare) = 0 Then

        QAP2P12_AgregarUnicoTexto = textoActual & separador & textoNuevo

    Else

        QAP2P12_AgregarUnicoTexto = textoActual

    End If

End Function


'=========================================================
' AGREGAR LÍNEAS CON LÍMITE
'=========================================================
Public Function QAP2P12_AgregarLineaConLimite(ByVal textoActual As String, _
                                              ByVal textoNuevo As String, _
                                              ByVal limiteLineas As Long) As String

    Dim cantidadLineas As Long

    textoNuevo = Trim(CStr(textoNuevo))

    If textoNuevo = "" Then
        QAP2P12_AgregarLineaConLimite = textoActual
        Exit Function
    End If

    cantidadLineas = QAP2P12_ContarLineas(textoActual)

    If cantidadLineas < limiteLineas Then

        If textoActual = "" Then
            QAP2P12_AgregarLineaConLimite = textoNuevo
        Else
            QAP2P12_AgregarLineaConLimite = textoActual & vbLf & textoNuevo
        End If

    ElseIf InStr(1, textoActual, "... ver detalle completo en el informe", vbTextCompare) = 0 Then

        QAP2P12_AgregarLineaConLimite = textoActual & vbLf & "... ver detalle completo en el informe"

    Else

        QAP2P12_AgregarLineaConLimite = textoActual

    End If

End Function


'=========================================================
' CONTAR LÍNEAS
'=========================================================
Public Function QAP2P12_ContarLineas(ByVal texto As String) As Long

    If Trim(CStr(texto)) = "" Then
        QAP2P12_ContarLineas = 0
    Else
        QAP2P12_ContarLineas = UBound(Split(texto, vbLf)) + 1
    End If

End Function


'=========================================================
' COLOR PRIORITARIO HALLAZGO
'=========================================================
Public Function QAP2P12_ColorPrioritarioHallazgo(ByVal estados As String) As Long

    estados = UCase(CStr(estados))

    If InStr(estados, "FAIL") > 0 Then

        QAP2P12_ColorPrioritarioHallazgo = RGB(192, 0, 0)

    ElseIf InStr(estados, "NOK") > 0 Then

        QAP2P12_ColorPrioritarioHallazgo = RGB(255, 199, 206)

    ElseIf InStr(estados, "DEL") > 0 Then

        QAP2P12_ColorPrioritarioHallazgo = RGB(246, 141, 46)

    ElseIf InStr(estados, "ADD") > 0 Then

        QAP2P12_ColorPrioritarioHallazgo = RGB(78, 195, 224)

    ElseIf InStr(estados, "PEND") > 0 Then

        QAP2P12_ColorPrioritarioHallazgo = RGB(255, 235, 156)

    Else

        QAP2P12_ColorPrioritarioHallazgo = RGB(0, 119, 200)

    End If

End Function


'=========================================================
' COLOR DE FONDO / TEXTO POR ESTADO (reutilizado en tabla de
' estados, gráfico y tabla de Detalle P2P)
'=========================================================
Public Function QAP2P12_ColorFondoEstado(ByVal estado As String) As Long

    estado = QAP2P12_LimpiarEstado(estado)

    Select Case estado

        Case "ok"
            QAP2P12_ColorFondoEstado = RGB(198, 239, 206)

        Case "no", "nok"
            QAP2P12_ColorFondoEstado = RGB(255, 199, 206)

        Case "fail"
            QAP2P12_ColorFondoEstado = RGB(192, 0, 0)

        Case "add"
            QAP2P12_ColorFondoEstado = RGB(189, 215, 238)

        Case "del"
            QAP2P12_ColorFondoEstado = RGB(246, 141, 46)

        Case "pend"
            QAP2P12_ColorFondoEstado = RGB(255, 235, 156)

        Case "n/a", "na"
            QAP2P12_ColorFondoEstado = RGB(217, 217, 217)

        Case Else
            QAP2P12_ColorFondoEstado = RGB(255, 255, 255)

    End Select

End Function


Public Function QAP2P12_ColorTextoEstado(ByVal estado As String) As Long

    estado = QAP2P12_LimpiarEstado(estado)

    Select Case estado

        Case "ok"
            QAP2P12_ColorTextoEstado = RGB(0, 97, 0)

        Case "no", "nok"
            QAP2P12_ColorTextoEstado = RGB(156, 0, 6)

        Case "fail"
            QAP2P12_ColorTextoEstado = RGB(255, 255, 255)

        Case "add"
            QAP2P12_ColorTextoEstado = RGB(0, 51, 102)

        Case "del"
            QAP2P12_ColorTextoEstado = RGB(255, 255, 255)

        Case "pend"
            QAP2P12_ColorTextoEstado = RGB(156, 101, 0)

        Case "n/a", "na"
            QAP2P12_ColorTextoEstado = RGB(0, 0, 0)

        Case Else
            QAP2P12_ColorTextoEstado = RGB(0, 0, 0)

    End Select

End Function


'=========================================================
' COLOREAR RANGO POR ESTADO
'=========================================================
Public Sub QAP2P12_ColorearRangoEstado(ByVal rng As Range, ByVal estado As String)

    With rng
        .Interior.Color = QAP2P12_ColorFondoEstado(estado)
        .Font.Color = QAP2P12_ColorTextoEstado(estado)
        .Font.Bold = True
    End With

End Sub
