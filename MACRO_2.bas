Attribute VB_Name = "MACRO_2"
Option Explicit

Public N°, ACTIVIDADES, Horas, Dias, Semanas As Long

Sub Cuadro()
Attribute Cuadro.VB_Description = "Macro grabada el 29/11/2018 por Administrador"
Attribute Cuadro.VB_ProcData.VB_Invoke_Func = " \n14"
Dim UltimaColumna, UltimaFila, UltimaColumnaSol, UltimaFilaInd As Long
Dim Cabecera, CeldaCabecera As Range
Dim Solicitudes, CeldaSolicitudes As Range
Dim ActividadRan, CeldaActividad As Range
Dim ITEM, Actividad, SOLICITADOS, RESPONDIDOS, PENDIENTES, CANT_TOTAL, ANTERIOR As Long
Dim AtendidoMes, MesSolicitado, MesRepuesta, FechaSolicitud, Solicitante, Titulo, Documento, _
    TOPICO, DESTINO, Status, FechaEnvioCorreo, FechaRecibido, FechaRepuesta, FechaRepuestaUsuario, TextoSolicitud, _
    TextoRepuesta, Recordatorio, TiempoMora, Observaciones As Long
Dim RefMes As String
Dim ValorSuma As Range
Dim MatrizActividad() As String
Dim A, Z, i As Long
Dim RanItem, RanRespondidos, RanPendientes As Range
Dim TotalSolicitado As Long
Dim TotalProcesado As Long

Application.ScreenUpdating = False
Worksheets("Solicitudes").Activate
If Worksheets("Solicitudes").FilterMode = True Then
    Worksheets("Solicitudes").ShowAllData
End If
Worksheets("INDICADORES").Activate

'BORRAR CUADRO
UltimaFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, 3).End(xlUp).Row
UltimaColumna = Worksheets("INDICADORES").Cells(5, Cells.Columns.Count).End(xlToLeft).Column

If UltimaFila > 5 Then
    Worksheets("INDICADORES").Range(Cells(6, 2), Cells(UltimaFila, 8)).Clear
End If

'VARIABLES DE CABECERA CUADRO INDICADORES

Worksheets("INDICADORES").Activate

RefMes = Cells(2, 5) & "-" & Cells(3, 5)

UltimaColumna = Worksheets("INDICADORES").Cells(5, Cells.Columns.Count).End(xlToLeft).Column
Set Cabecera = Worksheets("INDICADORES").Range(Cells(5, 1), Cells(5, UltimaColumna))
For Each CeldaCabecera In Cabecera
    If CeldaCabecera = "ITEM" Then
        ITEM = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "ACTIVIDAD" Then
        Actividad = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "SOLICITADOS MESES ANTERIORES (Pendientes)" Then
        ANTERIOR = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "SOLICITADOS EN EL MES" Then
        SOLICITADOS = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "RESPONDIDOS (R.)" Then
        RESPONDIDOS = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "PENDIENTES (P.)" Then
        PENDIENTES = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "CANT TOTAL (R+P)" Then
        CANT_TOTAL = CeldaCabecera.Column
    End If
        If CeldaCabecera = "N°" Then
        N° = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "ACTIVIDADES" Then
        ACTIVIDADES = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "Horas" Then
        Horas = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "Dias" Then
        Dias = CeldaCabecera.Column
        GoTo SIG1
    End If
    If CeldaCabecera = "Semanas" Then
        Semanas = CeldaCabecera.Column
        GoTo SIG1
    End If
SIG1:
Next CeldaCabecera

'VARIABLES DE CABECERA CUADRO INDICADORES

'VARIABLES DE CABECERA CUADRO SOLICITUDES

Worksheets("Solicitudes").Activate
UltimaColumnaSol = Worksheets("Solicitudes").Cells(1, Cells.Columns.Count).End(xlToLeft).Column
Set Cabecera = Worksheets("Solicitudes").Range(Cells(1, 1), Cells(1, UltimaColumnaSol))
For Each CeldaCabecera In Cabecera
    If CeldaCabecera = "ATENDIDO MES" Then
        AtendidoMes = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "MES SOLICITADO" Then
        MesSolicitado = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "MES REPUESTA" Then
        MesRepuesta = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "FECHA SOLICITUD" Then
        FechaSolicitud = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "SOLICITANTE" Then
        Solicitante = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "TITULO DEL CORREO" Then
        Titulo = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "DOCUMENTO" Then
        Documento = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "TOPICO" Then
        TOPICO = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "Ur" Then
        Ur = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "STATUS" Then
        Status = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "FECHA DE ENVIO DE CORREO" Then
        FechaEnvioCorreo = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "FECHA RECIBIDO CORREO" Then
        FechaRecibido = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "FECHA REPUESTA CORREO" Then
        FechaRepuesta = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "FECHA REPUESTA USUARIO" Then
        FechaRepuestaUsuario = CeldaCabecera.Column
        GoTo SIG2
    End If
        If CeldaCabecera = "SOLICITUD" Then
        TextoSolicitud = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "REPUESTA" Then
        TextoRepuesta = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "RECORDATORIO" Then
        Recordatorio = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "TIEMPO DE MORA" Then
        TiempoMora = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "OBSERVACIONES" Then
        Observaciones = CeldaCabecera.Column
        GoTo SIG2
    End If
    

SIG2:
Next CeldaCabecera

'VARIABLES DE CABECERA CUADRO SOLICITUDES

'MATRIZ ACTIVIDAD
UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 2).End(xlUp).Row

Set ActividadRan = Worksheets("Solicitudes").Range(Cells(1, AtendidoMes), Cells(UltimaFila, AtendidoMes))
For Each CeldaActividad In ActividadRan
    If IsEmpty(Z) Then GoTo SIGUIENTE1
    
    For i = 0 To Z
        If CeldaActividad.Offset(0, 7).Value = MatrizActividad(0, i) Then GoTo SIGUIENTE2
    Next i
    
SIGUIENTE1:
    If CeldaActividad = RefMes Or CeldaActividad.Offset(0, 10) = "PENDIENTE" Then
        If IsEmpty(A) Then A = 0
        ReDim Preserve MatrizActividad(0, A)
        MatrizActividad(0, A) = CeldaActividad.Offset(0, 7).Value
        Z = A
        A = A + 1
    End If
SIGUIENTE2:

Next CeldaActividad

If IsEmpty(Z) = True Then GoTo SIGUIENTE3 'NO HAY ACTIVIDADES DEL MES

For i = 0 To Z
    Worksheets("INDICADORES").Cells(6 + i, Actividad) = MatrizActividad(0, i)
    Worksheets("INDICADORES").Cells(6 + i, Actividad).Offset(0, -1) = i + 1
Next i

'MATRIZ ACTIVIDAD
Worksheets("INDICADORES").Activate
UltimaFilaInd = Worksheets("INDICADORES").Cells(Cells.Rows.Count, Actividad).End(xlUp).Row

'SOLICITADOS MESES ANTERIORES (Pendientes)
For i = 6 To UltimaFilaInd
    For Each CeldaActividad In ActividadRan
        If ((CeldaActividad = RefMes And CeldaActividad.Offset(0, 1) <> RefMes) Or _
            (CeldaActividad.Offset(0, Status - AtendidoMes) = "PENDIENTE" And CeldaActividad <> RefMes And CeldaActividad.Offset(0, 1) <> RefMes)) And CeldaActividad.Offset(0, TOPICO - AtendidoMes) = Worksheets("INDICADORES").Cells(i, Actividad) Then
            Worksheets("INDICADORES").Cells(i, ANTERIOR) = Cells(i, ANTERIOR).Value + 1
            Worksheets("INDICADORES").Cells(i, ANTERIOR).NumberFormat = "0"
            Worksheets("INDICADORES").Cells(i, ANTERIOR).HorizontalAlignment = xlCenter
        End If
    Next CeldaActividad
Next i


'SOLICITADOS
For i = 6 To UltimaFilaInd
    For Each CeldaActividad In ActividadRan
        If CeldaActividad.Offset(0, 1) = RefMes And CeldaActividad.Offset(0, TOPICO - AtendidoMes) = Worksheets("INDICADORES").Cells(i, Actividad) Then
            Worksheets("INDICADORES").Cells(i, SOLICITADOS) = Cells(i, SOLICITADOS).Value + 1
            Worksheets("INDICADORES").Cells(i, SOLICITADOS).NumberFormat = "0"
            Worksheets("INDICADORES").Cells(i, SOLICITADOS).HorizontalAlignment = xlCenter
        End If
    Next CeldaActividad
Next i

'RESPONDIDOS (R.)
For i = 6 To UltimaFilaInd
    For Each CeldaActividad In ActividadRan
        If CeldaActividad.Offset(0, 2) = RefMes And CeldaActividad.Offset(0, TOPICO - AtendidoMes) = Worksheets("INDICADORES").Cells(i, Actividad) And CeldaActividad.Offset(0, Status - AtendidoMes) = "LISTO" Then
            Worksheets("INDICADORES").Cells(i, RESPONDIDOS) = Cells(i, RESPONDIDOS).Value + 1
            Worksheets("INDICADORES").Cells(i, RESPONDIDOS).NumberFormat = "0"
            Worksheets("INDICADORES").Cells(i, RESPONDIDOS).HorizontalAlignment = xlCenter
        End If
    Next CeldaActividad
Next i

'PENDIENTES
For i = 6 To UltimaFilaInd
    For Each CeldaActividad In ActividadRan
        If CeldaActividad.Offset(0, TOPICO - AtendidoMes) = Worksheets("INDICADORES").Cells(i, Actividad) And CeldaActividad.Offset(0, Status - AtendidoMes) = "PENDIENTE" Then
            Worksheets("INDICADORES").Cells(i, PENDIENTES) = Cells(i, PENDIENTES).Value + 1
            Worksheets("INDICADORES").Cells(i, PENDIENTES).NumberFormat = "0"
            Worksheets("INDICADORES").Cells(i, PENDIENTES).HorizontalAlignment = xlCenter
        End If
    Next CeldaActividad
Next i

'CANT TOTAL
For i = 6 To UltimaFilaInd
    Worksheets("INDICADORES").Cells(i, CANT_TOTAL) = Worksheets("INDICADORES").Cells(i, RESPONDIDOS) + Worksheets("INDICADORES").Cells(i, PENDIENTES)
Next i

'TOTALES
Worksheets("INDICADORES").Cells(UltimaFilaInd + 1, Actividad) = "TOTAL"
Worksheets("INDICADORES").Cells(UltimaFilaInd + 1, Actividad).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .ColorIndex = 44
        .Pattern = xlSolid
    End With

For i = ANTERIOR To CANT_TOTAL
    Set ValorSuma = Worksheets("INDICADORES").Range(Cells(6, i), Cells(UltimaFilaInd, i))
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 1, i) = Application.WorksheetFunction.Sum(ValorSuma)
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 1, i).Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .ColorIndex = 44
        .Pattern = xlSolid
    End With
Next i

'VALIDACIÓN

TotalSolicitado = Worksheets("INDICADORES").Cells(UltimaFilaInd + 2, ANTERIOR) + Worksheets("INDICADORES").Cells(UltimaFilaInd + 2, SOLICITADOS)
TotalProcesado = Worksheets("INDICADORES").Cells(UltimaFilaInd + 2, CANT_TOTAL)

If TotalSolicitado <> TotalProcesado Then
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 4, ACTIVIDADES).FormulaR1C1 = "VERIFICAR TOTALES"
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 4, ACTIVIDADES).Select
    Selection.Interior.ColorIndex = 3
    Selection.Font.ColorIndex = 2
Else
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 4, ACTIVIDADES).FormulaR1C1 = ""
    Worksheets("INDICADORES").Cells(UltimaFilaInd + 4, ACTIVIDADES).Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.ColorIndex = 0
End If

'BORDES
Worksheets("INDICADORES").Range(Cells(5, ITEM), Cells(UltimaFilaInd + 1, CANT_TOTAL)).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Worksheets("INDICADORES").Range("C6").Select


'GRAFICA 1
Set RanItem = Worksheets("INDICADORES").Range(Cells(6, ITEM), Cells(UltimaFilaInd, ITEM))
Set RanRespondidos = Worksheets("INDICADORES").Range(Cells(6, RESPONDIDOS), Cells(UltimaFilaInd, RESPONDIDOS))
Set RanPendientes = Worksheets("INDICADORES").Range(Cells(6, PENDIENTES), Cells(UltimaFilaInd, PENDIENTES))
ActiveSheet.ChartObjects(1).Activate
ActiveChart.SeriesCollection(1).XValues = RanItem
ActiveChart.SeriesCollection(1).Values = RanRespondidos
ActiveChart.SeriesCollection(2).Values = RanPendientes

Worksheets("INDICADORES").Range("C4").Select

Call CuadroDepartamento
SIGUIENTE3:
Call HorasActividad

End Sub

Sub CuadroDepartamento()
Dim MesRang, MesCel As Range
Dim ResRang, ResCelda As Range
Dim UltimaFila2, i, Row, Column, Tiempo, Cant As Long
Application.ScreenUpdating = False

Worksheets("INDICADORES").Activate
Worksheets("INDICADORES").Range(Cells(6, 26), Cells(8, 37)).ClearContents

Set MesRang = Worksheets("INDICADORES").Range(Cells(5, 26), Cells(5, 37))

Worksheets("Solicitudes").Activate
UltimaFila2 = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 4).End(xlUp).Row
Set ResRang = Worksheets("Solicitudes").Range(Cells(2, 4), Cells(UltimaFila2, 4))
Worksheets("INDICADORES").Activate

For Each MesCel In MesRang
    MesRepuesta = MesCel.Offset(-1, 0).Value
    Row = MesCel.Row
    Column = MesCel.Column
    For i = 6 To 8
        Tiempo = 0
        Cant = 0
        For Each ResCelda In ResRang
            If ResCelda = MesRepuesta And ResCelda.Offset(0, 6) = Worksheets("INDICADORES").Cells(i, 25) Then
                Tiempo = Tiempo + ResCelda.Offset(0, 17)
                Cant = Cant + 1
            End If
        Next ResCelda
        If Cant = 0 Then GoTo SIGUIENTE
        Worksheets("INDICADORES").Cells(i, Column) = Format(Tiempo / Cant, "0")
SIGUIENTE:
    Next i
Next MesCel

End Sub


Sub HorasActividad()
Dim UlFila, FilaSup, FilaInf, UlColumna As Long
Dim RefRang, RefCel As Range
Dim FeriadosRang, FeriadosCelda As Range
Dim ActvidadRang, ActividadCelda As Range
Dim GraficoItem, GraficoHoras As Range
Dim MES As String
Dim AÑO, TemMes, i, A, B As Long
Dim MatrizActividades() As String
Dim RangoMes, CeldaMes As Range
Dim RangoDiasMes, CeldaDiasMes As Range
Dim RangoTrabajo, Celdatrabajo As Range
Dim HoraMes As Long
Dim HoraTrabajo, HoraPermiso, HoraReposo As Single

'BORRAR CUADRO
Worksheets("INDICADORES").Activate
UlFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, ACTIVIDADES).End(xlUp).Row

If UlFila > 5 Then
    Worksheets("INDICADORES").Range(Cells(6, N°), Cells(UlFila, Semanas)).Clear
End If

AÑO = Worksheets("INDICADORES").Cells(2, 5)
TemMes = Worksheets("INDICADORES").Cells(3, 5)

If TemMes = 1 Then MES = "ENERO"
If TemMes = 2 Then MES = "FEBRERO"
If TemMes = 3 Then MES = "MARZO"
If TemMes = 4 Then MES = "ABRIL"
If TemMes = 5 Then MES = "MAYO"
If TemMes = 6 Then MES = "JUNIO"
If TemMes = 7 Then MES = "JULIO"
If TemMes = 8 Then MES = "AGOSTO"
If TemMes = 9 Then MES = "SEPTIEMBRE"
If TemMes = 10 Then MES = "OCTUBRE"
If TemMes = 11 Then MES = "NOVIEMBRE"
If TemMes = 12 Then MES = "DICIEMBRE"

Worksheets("Solicitudes").Activate

UlFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 29).End(xlUp).Row
Set FeriadosRang = Worksheets("Solicitudes").Range(Cells(1, 29), Cells(UlFila, 28))

Worksheets("BITACORA").Activate

UlFila = Worksheets("BITACORA").Cells(Cells.Rows.Count, 1).End(xlUp).Row
Set RefRang = Worksheets("BITACORA").Range(Cells(1, 1), Cells(UlFila, 1))

'CARGA LA MATRIZ DE ACTIVIDADES

For Each RefCel In RefRang
    If RefCel = AÑO & "-" & MES Then
        FilaSup = RefCel.Row + 2
        FilaInf = FilaSup + 39
        UlColumna = Worksheets("BITACORA").Cells(FilaSup - 1, Cells.Columns.Count).End(xlToLeft).Column
        Set ActvidadRang = Worksheets("BITACORA").Range(Cells(FilaSup, 3), Cells(FilaInf, UlColumna))
        For Each ActividadCelda In ActvidadRang
            If IsEmpty(ActividadCelda.Value) = True Then GoTo SIG1
                For Each FeriadosCelda In FeriadosRang
                    If ActividadCelda.Value = FeriadosCelda.Value Then GoTo SIG1
                Next FeriadosCelda
            If ActividadCelda = "TOMAR TRANSPORTE PARA COMEDOR" Then GoTo SIG1
            If ActividadCelda = "COMEDOR" Then GoTo SIG1
            If ActividadCelda = "TOMAR TRANSPORTE DE REGRESO" Then GoTo SIG1
            If IsEmpty(i) Then
                i = 0
                ReDim Preserve MatrizActividades(4, i)
                MatrizActividades(0, i) = 1 'N°
                MatrizActividades(1, i) = ActividadCelda.Value
                MatrizActividades(2, i) = 30 'HORA
                MatrizActividades(3, i) = 30 'DIAS
                MatrizActividades(4, i) = 30 'SEMENAS
                GoTo SIG2
            End If
            For B = 0 To A
                If ActividadCelda.Value = MatrizActividades(1, B) Then
                    MatrizActividades(2, B) = MatrizActividades(2, B) + 30 'HORA
                    MatrizActividades(3, B) = MatrizActividades(3, B) + 30 'DIAS
                    MatrizActividades(4, B) = MatrizActividades(4, B) + 30 'SEMENAS
                    GoTo SIG1
                End If
            Next B
            ReDim Preserve MatrizActividades(4, i)
            MatrizActividades(0, i) = i + 1 'N°
            MatrizActividades(1, i) = ActividadCelda.Value 'ACTIVIDAD
            MatrizActividades(2, i) = 30 'HORA
            MatrizActividades(3, i) = 30 'DIAS
            MatrizActividades(4, i) = 30 'SEMENAS
SIG2:
            A = i
            i = i + 1
SIG1:
        Next ActividadCelda
        
    End If
Next RefCel

Worksheets("INDICADORES").Activate

'DESCARGA DE MATRIZ DE ACTIVIDADES
B = 0
For B = 0 To A
    Worksheets("INDICADORES").Cells(6 + B, N°) = MatrizActividades(0, B)
    Worksheets("INDICADORES").Cells(6 + B, ACTIVIDADES) = MatrizActividades(1, B)
    Worksheets("INDICADORES").Cells(6 + B, Horas) = Format(MatrizActividades(2, B) / 60, "#,##0.0") * 1
    Worksheets("INDICADORES").Cells(6 + B, Dias) = Format(MatrizActividades(3, B) / 480, "#,##0.0") * 1
    Worksheets("INDICADORES").Cells(6 + B, Semanas) = Format(MatrizActividades(4, B) / 2400, "#,##0.0") * 1
Next B

UlFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, ACTIVIDADES).End(xlUp).Row

'BORDES
Worksheets("INDICADORES").Range(Cells(5, N°), Cells(UlFila, Semanas)).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

'NUMEROS CENTRADOS
    Worksheets("INDICADORES").Range(Cells(6, Horas), Cells(UlFila, Semanas)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'ORDEN
Worksheets("INDICADORES").Range(Cells(6, ACTIVIDADES), Cells(UlFila, Semanas)).Select
    Selection.Sort Key1:=Worksheets("INDICADORES").Range(Cells(1, Horas), Cells(UlFila, Horas)), Order1:=xlDescending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    
'GRAFICA 1
Set GraficoItem = Worksheets("INDICADORES").Range(Cells(6, N°), Cells(UlFila, N°))
Set GraficoHoras = Worksheets("INDICADORES").Range(Cells(6, Horas), Cells(UlFila, Horas))
ActiveSheet.ChartObjects("Gráfico 13").Activate
ActiveChart.ChartArea.Select
ActiveChart.SeriesCollection(1).XValues = GraficoItem
ActiveChart.SeriesCollection(1).Values = GraficoHoras

'CUADRO HH-TRABAJO HABILES
Worksheets("INDICADORES").Activate
Worksheets("INDICADORES").Range("C6").Select
Set RangoMes = Worksheets("INDICADORES").Range(Cells(4, 49), Cells(4, 60))
For Each CeldaMes In RangoMes
    For Each RefCel In RefRang
        If CeldaMes = RefCel Then
            HoraMes = 0
            HoraTrabajo = 0
            HoraPermiso = 0
            HoraReposo = 0
            UlColumna = Worksheets("BITACORA").Cells(RefCel.Row, Cells.Columns.Count).End(xlToLeft).Column
            Worksheets("BITACORA").Activate
            'IDENTIFICAR DIAS HABILES
            Set RangoDiasMes = Worksheets("BITACORA").Range(Cells(RefCel.Row, 3), Cells(RefCel.Row, UlColumna))
            For Each CeldaDiasMes In RangoDiasMes
                For Each FeriadosCelda In FeriadosRang
                    If CeldaDiasMes = "" Or CeldaDiasMes = FeriadosCelda Or CeldaDiasMes.Offset(1, 0) = "SABADO" Or _
                        CeldaDiasMes.Offset(1, 0) = "DOMINGO" Then
                        GoTo Sigui44
                    End If
                Next FeriadosCelda
                HoraMes = HoraMes + 8
Sigui44:
                'IDENTIFICAR HORAS TRABAJADAS
                If CeldaDiasMes > 0 Then
                    Set RangoTrabajo = Worksheets("BITACORA").Range(Cells(RefCel.Row + 2, CeldaDiasMes.Column), Cells(RefCel.Row + 39, CeldaDiasMes.Column))
                    For Each Celdatrabajo In RangoTrabajo
                        Celdatrabajo.Select
                        If CeldaDiasMes > 0 And Celdatrabajo.Interior.ColorIndex = xlNone And Celdatrabajo <> "" And Celdatrabajo <> "Permiso" And _
                            Celdatrabajo <> "Reposo Medico" Then
                            HoraTrabajo = HoraTrabajo + 0.5
                        ElseIf CeldaDiasMes > 0 And Celdatrabajo.Interior.ColorIndex = xlNone And Celdatrabajo <> "" And Celdatrabajo.Value = "Permiso" Then
                            HoraPermiso = HoraPermiso + 0.5
                        ElseIf CeldaDiasMes > 0 And Celdatrabajo.Interior.ColorIndex = xlNone And Celdatrabajo <> "" And Celdatrabajo.Value = "Reposo Medico" Then
                            HoraReposo = HoraReposo + 0.5
                        End If
                    Next Celdatrabajo
                End If
            Next CeldaDiasMes
            GoTo SIGUI111
        End If
    Next RefCel
SIGUI111:
    CeldaMes.Offset(2, 0) = HoraMes
    CeldaMes.Offset(3, 0) = Format(HoraTrabajo, "#,##0.0") * 1
    CeldaMes.Offset(4, 0) = Format(HoraPermiso, "#,##0.0") * 1
    CeldaMes.Offset(5, 0) = Format(HoraReposo, "#,##0.0") * 1
    CeldaMes.Offset(6, 0) = Format((CeldaMes.Offset(3, 0) / CeldaMes.Offset(2, 0)), "00%")
Next CeldaMes

Call ComentarioActiviades

Worksheets("INDICADORES").Activate
Worksheets("INDICADORES").Range("C6").Select

End Sub

Sub ComentarioActiviades()

Dim UlFila, FilaSup, FilaInf, UlColumna As Long
Dim RefRang, RefCel As Range
Dim FeriadosRang, FeriadosCelda As Range
Dim ActvidadRang, ActividadCelda As Range
Dim GraficoItem, GraficoHoras As Range
Dim MES As String
Dim AÑO, TemMes, i, A, B, D As Long
Dim MatrizActividades() As String
Dim RangoMes, CeldaMes As Range
Dim RangoDiasMes, CeldaDiasMes As Range
Dim RangoTrabajo, Celdatrabajo As Range
Dim HoraMes As Long
Dim HoraTrabajo As Single
Dim RangoLista, CeldaLista As Range

AÑO = Worksheets("INDICADORES").Cells(2, 5)
TemMes = Worksheets("INDICADORES").Cells(3, 5)

If TemMes = 1 Then MES = "ENERO"
If TemMes = 2 Then MES = "FEBRERO"
If TemMes = 3 Then MES = "MARZO"
If TemMes = 4 Then MES = "ABRIL"
If TemMes = 5 Then MES = "MAYO"
If TemMes = 6 Then MES = "JUNIO"
If TemMes = 7 Then MES = "JULIO"
If TemMes = 8 Then MES = "AGOSTO"
If TemMes = 9 Then MES = "SEPTIEMBRE"
If TemMes = 10 Then MES = "OCTUBRE"
If TemMes = 11 Then MES = "NOVIEMBRE"
If TemMes = 12 Then MES = "DICIEMBRE"

Worksheets("Solicitudes").Activate

UlFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 29).End(xlUp).Row
Set FeriadosRang = Worksheets("Solicitudes").Range(Cells(1, 29), Cells(UlFila, 28))

Worksheets("BITACORA").Activate

UlFila = Worksheets("BITACORA").Cells(Cells.Rows.Count, 1).End(xlUp).Row
Set RefRang = Worksheets("BITACORA").Range(Cells(1, 1), Cells(UlFila, 1))

'CARGA LA MATRIZ DE ACTIVIDADES

For Each RefCel In RefRang
    If RefCel = AÑO & "-" & MES Then
        FilaSup = RefCel.Row + 2
        FilaInf = FilaSup + 39
        UlColumna = Worksheets("BITACORA").Cells(FilaSup - 1, Cells.Columns.Count).End(xlToLeft).Column
        For D = 3 To UlColumna
        Set ActvidadRang = Worksheets("BITACORA").Range(Cells(FilaSup, D), Cells(FilaInf, D))
        For Each ActividadCelda In ActvidadRang
            If ActividadCelda.Comment Is Nothing Then
                GoTo SIG1
            End If
            If IsEmpty(ActividadCelda.Value) = True Then GoTo SIG1
                For Each FeriadosCelda In FeriadosRang
                    If ActividadCelda.Value = FeriadosCelda.Value Then GoTo SIG1
                Next FeriadosCelda
            If ActividadCelda = "TOMAR TRANSPORTE PARA COMEDOR" Then GoTo SIG1
            If ActividadCelda = "COMEDOR" Then GoTo SIG1
            If ActividadCelda = "TOMAR TRANSPORTE DE REGRESO" Then GoTo SIG1
            If IsEmpty(i) Then
                i = 0
                ReDim Preserve MatrizActividades(3, i)
                MatrizActividades(0, i) = ActividadCelda.Value 'Actividad
                MatrizActividades(1, i) = ActividadCelda.Comment.Text 'Comentario
                MatrizActividades(2, i) = Worksheets("BITACORA").Cells(RefCel.Row, ActividadCelda.Column) 'Fecha
                MatrizActividades(3, i) = 0.5 'Min
                GoTo SIG2
            End If
            For B = 0 To A
                If ActividadCelda.Value = MatrizActividades(0, B) And ActividadCelda.Comment.Text = MatrizActividades(1, B) Then
                    MatrizActividades(3, B) = MatrizActividades(3, B) + 0.5 'Min
                    GoTo SIG1
                End If
            Next B
            
            ReDim Preserve MatrizActividades(3, i)
            MatrizActividades(0, i) = ActividadCelda.Value 'Actividad
            MatrizActividades(1, i) = ActividadCelda.Comment.Text 'Comentario
            MatrizActividades(2, i) = Worksheets("BITACORA").Cells(RefCel.Row, ActividadCelda.Column) 'Fecha
            MatrizActividades(3, i) = 0.5 'Min
SIG2:
            A = i
            i = i + 1
SIG1:
        Next ActividadCelda
        Next D
        
    End If
Next RefCel

Worksheets("INDICADORES").Activate

'DESCARGA DE MATRIZ DE ACTIVIDADES
UlFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, ACTIVIDADES).End(xlUp).Row
Set RangoLista = Worksheets("INDICADORES").Range(Cells(6, ACTIVIDADES), Cells(UlFila, ACTIVIDADES))
For Each CeldaLista In RangoLista
    For B = 0 To A
        If CeldaLista = MatrizActividades(0, B) And CeldaLista.Comment Is Nothing Then
            CeldaLista.AddComment
            CeldaLista.Comment.Text Text:=MatrizActividades(1, B) & Chr(10) & "Fecha: " & MatrizActividades(2, B) & " HH: " & MatrizActividades(3, B) & Chr(10)
        ElseIf CeldaLista = MatrizActividades(0, B) Then
            CeldaLista.Comment.Text Text:=CeldaLista.Comment.Text & Chr(10) & _
            MatrizActividades(1, B) & Chr(10) & "Fecha: " & MatrizActividades(2, B) & " HH: " & MatrizActividades(3, B) & Chr(10)
        End If
    Next B
    
    If CeldaLista.Comment Is Nothing Then GoTo Siguiente12
        CeldaLista.Comment.Shape.ScaleHeight 3, msoFalse, msoScaleFromTopLeft
        CeldaLista.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
    
Siguiente12:
Next CeldaLista

End Sub
