Attribute VB_Name = "MACRO_1"
Dim Ejecutando As Boolean
Public CajaTexto, Comentario As String
Public Hora As Date
Public ColumnaDia, FilaDia, TempColumnaDia, TempFilaDia As Long
Public RefHora2 As String
Public Ejecutando2 As Boolean
Public Tiempo2 As Date
Public CANCELADO As Boolean
Public LibroActivo As String
Public HojaActiva As String
Public Hoja As Boolean
Public NombreLibro As String
Public Libro As Workbook


Sub ProgramarMacro()

    Tiempo0 = "06:55:00"
    Tiempo1 = "07:25:00"
    Tiempo2 = "07:55:00"
    Tiempo3 = "08:25:00"
    Tiempo4 = "08:55:00"
    Tiempo5 = "09:25:00"
    Tiempo6 = "09:55:00"
    Tiempo7 = "10:25:00"
    Tiempo8 = "10:55:00"
    Tiempo9 = "11:05:00"
    Tiempo99 = "11:35:00"
  
    Tiempo10 = "13:55:00"
    Tiempo11 = "14:25:00"
    Tiempo12 = "14:55:00"
    Tiempo13 = "15:25:00"
    Tiempo14 = "15:55:00"
    Tiempo15 = "15:58:00" '16:20
    
    Tiempo16 = "16:55:00"
    Tiempo17 = "17:25:00"
    Tiempo18 = "17:55:00"
    Tiempo19 = "18:25:00"
    Tiempo20 = "18:55:00"
    Tiempo21 = "19:25:00"
    Tiempo22 = "19:55:00"
    Tiempo23 = "20:25:00"
    Tiempo24 = "20:55:00"
    
    Application.OnTime Tiempo0, "DiaHora", , Ejecutando
    Application.OnTime Tiempo1, "DiaHora", , Ejecutando
    Application.OnTime Tiempo2, "DiaHora", , Ejecutando
    Application.OnTime Tiempo3, "DiaHora", , Ejecutando
    Application.OnTime Tiempo4, "DiaHora", , Ejecutando
    Application.OnTime Tiempo5, "DiaHora", , Ejecutando
    Application.OnTime Tiempo6, "DiaHora", , Ejecutando
    Application.OnTime Tiempo7, "DiaHora", , Ejecutando
    Application.OnTime Tiempo8, "DiaHora", , Ejecutando
    Application.OnTime Tiempo9, "DiaHora", , Ejecutando
    Application.OnTime Tiempo99, "DiaHora", , Ejecutando
    Application.OnTime Tiempo10, "DiaHora", , Ejecutando
    Application.OnTime Tiempo11, "DiaHora", , Ejecutando
    Application.OnTime Tiempo12, "DiaHora", , Ejecutando
    Application.OnTime Tiempo13, "DiaHora", , Ejecutando
    Application.OnTime Tiempo14, "DiaHora", , Ejecutando
    Application.OnTime Tiempo15, "DiaHora", , Ejecutando
    Application.OnTime Tiempo16, "DiaHora", , Ejecutando
    Application.OnTime Tiempo17, "DiaHora", , Ejecutando
    Application.OnTime Tiempo18, "DiaHora", , Ejecutando
    Application.OnTime Tiempo19, "DiaHora", , Ejecutando
    Application.OnTime Tiempo20, "DiaHora", , Ejecutando
    Application.OnTime Tiempo21, "DiaHora", , Ejecutando
    Application.OnTime Tiempo22, "DiaHora", , Ejecutando
    Application.OnTime Tiempo23, "DiaHora", , Ejecutando
    Application.OnTime Tiempo24, "DiaHora", , Ejecutando
        
End Sub

Sub DetenerReloj()
    Ejecutando = False
    Application.OnTime Tiempo, "Prueba", , False
End Sub

Sub IniciarReloj()
    Ejecutando = True
    Call ProgramarMacro
End Sub

Sub DiaHora()
Attribute DiaHora.VB_ProcData.VB_Invoke_Func = " \n14"
Dim UltimaFila, FilaPrincipal As Long
Dim Rango, CELDA, RangoColumna, CeldaColumna, RangoFila, CeldaFila As Range
Dim AñoMes As String
Dim HoraReferenciaMenor As Date
Dim HoraReferenciaMayor As Date
Dim RefHora As String

Ejecutando2 = False
CANCELADO = False
If Hoja = False Then
    LibroActivo = ActiveWorkbook.Name
    HojaActiva = ActiveSheet.Name
    Hoja = True
End If

If UserForm3.Visible = True Then GoTo Cancelar

NombreLibro = ThisWorkbook.Name
Windows(NombreLibro).Activate
ThisWorkbook.Activate

UltimaFila = Worksheets("BITACORA").Range("A" & Rows.Count).End(xlUp).Row
Set Rango = Worksheets("BITACORA").Range("A1:A" & UltimaFila)
AñoMes = Format(Date, "yyyy") & "-" & UCase(Format(Date, "MMMM"))

For Each CELDA In Rango
    If CELDA = AñoMes Then
        FilaPrincipal = CELDA.Row
        UltimaColumna = Worksheets("BITACORA").Cells(FilaPrincipal, Cells.Columns.Count).End(xlToLeft).Column
        Worksheets("BITACORA").Activate
        Set RangoColumna = Worksheets("BITACORA").Range(Cells(FilaPrincipal, 2), Cells(FilaPrincipal, UltimaColumna))
        For Each CeldaColumna In RangoColumna
            If CeldaColumna = Date Then
                ColumnaDia = CeldaColumna.Column
                Hora = Format(Hour(Now), "##") & ":" & Format(Minute(Now), "##")
                
                HoraReferenciaMenor = "06:50"
                HoraReferenciaMayor = "07:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 0
                    RefHora2 = 1
                End If
                HoraReferenciaMenor = "07:20"
                HoraReferenciaMayor = "07:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 1
                    RefHora2 = 1
                End If
                HoraReferenciaMenor = "07:50"
                HoraReferenciaMayor = "08:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 2
                    RefHora2 = 2
                End If
                HoraReferenciaMenor = "08:20"
                HoraReferenciaMayor = "08:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 3
                    RefHora2 = 3
                End If
                HoraReferenciaMenor = "08:50"
                HoraReferenciaMayor = "09:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 4
                    RefHora2 = 4
                End If
                HoraReferenciaMenor = "09:20"
                HoraReferenciaMayor = "09:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 5
                    RefHora2 = 5
                End If
                HoraReferenciaMenor = "09:50"
                HoraReferenciaMayor = "10:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 6
                    RefHora2 = 6
                End If
                HoraReferenciaMenor = "10:20"
                HoraReferenciaMayor = "10:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 7
                    RefHora2 = 7
                End If
                HoraReferenciaMenor = "10:50"
                HoraReferenciaMayor = "11:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 8
                    RefHora2 = 8
                End If
                HoraReferenciaMenor = "11:00"
                HoraReferenciaMayor = "11:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 9
                    RefHora2 = 9
                End If
                HoraReferenciaMenor = "11:30"
                HoraReferenciaMayor = "12:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 10
                    RefHora2 = 10
                End If

                HoraReferenciaMenor = "12:20"
                HoraReferenciaMayor = "12:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 11
                    RefHora2 = 11
                End If
                HoraReferenciaMenor = "12:50"
                HoraReferenciaMayor = "13:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 12
                    RefHora2 = 12
                End If
                HoraReferenciaMenor = "13:20"
                HoraReferenciaMayor = "13:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 13
                    RefHora2 = 13
                End If
                HoraReferenciaMenor = "13:50"   '13:30
                HoraReferenciaMayor = "14:00"   '14:00
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 14
                    RefHora2 = 14
                End If
                HoraReferenciaMenor = "14:20"   '14:00
                HoraReferenciaMayor = "14:30"   '14:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 15
                    RefHora2 = 15
                End If
                HoraReferenciaMenor = "14:50"   '14:30
                HoraReferenciaMayor = "15:00"   '15:0
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 16
                    RefHora2 = 16
                End If
                HoraReferenciaMenor = "15:20"   '15:00
                HoraReferenciaMayor = "15:30"   '15:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 17
                    RefHora2 = 17
                End If
                HoraReferenciaMenor = "15:50"   '15:30
                HoraReferenciaMayor = "15:56"   '16:00
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 18
                    RefHora2 = 18
                End If
                HoraReferenciaMenor = "15:57"   '16:00
                HoraReferenciaMayor = "16:30"   '16:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 19
                    RefHora2 = 19
                End If
                HoraReferenciaMenor = "16:50"
                HoraReferenciaMayor = "17:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 20
                    RefHora2 = 20
                End If
                HoraReferenciaMenor = "17:20"
                HoraReferenciaMayor = "17:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 21
                    RefHora2 = 21
                End If
                HoraReferenciaMenor = "17:50"
                HoraReferenciaMayor = "18:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 22
                    RefHora2 = 22
                End If
                HoraReferenciaMenor = "18:20"
                HoraReferenciaMayor = "18:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 23
                    RefHora2 = 23
                End If
                HoraReferenciaMenor = "18:50"
                HoraReferenciaMayor = "19:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 24
                    RefHora2 = 24
                End If
                HoraReferenciaMenor = "19:20"
                HoraReferenciaMayor = "19:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 25
                    RefHora2 = 25
                End If
                HoraReferenciaMenor = "19:50"
                HoraReferenciaMayor = "20:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 26
                    RefHora2 = 26
                End If
                HoraReferenciaMenor = "20:20"
                HoraReferenciaMayor = "20:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 27
                    RefHora2 = 27
                End If
                HoraReferenciaMenor = "20:50"
                HoraReferenciaMayor = "21:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 28
                    RefHora2 = 28
                End If
                HoraReferenciaMenor = "21:20"
                HoraReferenciaMayor = "21:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 29
                    RefHora2 = 29
                End If
                HoraReferenciaMenor = "21:50"
                HoraReferenciaMayor = "22:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 30
                    RefHora2 = 30
                End If
                HoraReferenciaMenor = "22:20"
                HoraReferenciaMayor = "22:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 31
                    RefHora2 = 31
                End If
                HoraReferenciaMenor = "22:50"
                HoraReferenciaMayor = "23:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 32
                    RefHora2 = 32
                End If
                HoraReferenciaMenor = "23:20"
                HoraReferenciaMayor = "23:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 33
                    RefHora2 = 33
                End If
                HoraReferenciaMenor = "23:50"
                HoraReferenciaMayor = "23:59"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 34
                    RefHora2 = 34
                End If
                HoraReferenciaMenor = "00:20"
                HoraReferenciaMayor = "00:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 35
                    RefHora2 = 35
                End If
                HoraReferenciaMenor = "00:50"
                HoraReferenciaMayor = "01:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 36
                    RefHora2 = 36
                End If
                Set RangoFila = Worksheets("BITACORA").Range("A1:A" & UltimaFila)
                If RefHora = "" Then GoTo Cancelar2
                For Each CeldaFila In Rango
                    If CeldaFila = RefHora Then
                        FilaDia = CeldaFila.Row
                        If CANCELADO = True Then Exit Sub
                        Call Verificacion
                        If CANCELADO = True Then Exit Sub
                        CajaTexto = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia)
                        If Not Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                            Comentario = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text
                        End If
                        
                        
                        Ejecutando2 = True
                        Call MACRO_1.TiempoEspera
                        UserForm3.Show
                        Call Verificacion
                        If CANCELADO = True Then Exit Sub
                        Workbooks("GESTIÓN.xls").Activate
                        Worksheets("BITACORA").Activate
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                                Selection.Interior.ColorIndex = xlNone
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) = CajaTexto
                        If IsEmpty(Comentario) = False And Comentario <> "" Then
                            If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).AddComment
                            End If
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Visible = False
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text Text:=Comentario
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
                        End If
                        GoTo Cancelar
                    End If
                Next CeldaFila
            End If
        Next CeldaColumna
    End If
Next CELDA

If RefHora2 = 6 Or RefHora2 = 14 Or RefHora2 = 17 Then
    Call ActividadPendiente
End If
Cancelar:
For Each Libro In Workbooks
    If Libro.Name = LibroActivo Then
        Workbooks(LibroActivo).Worksheets(HojaActiva).Activate
        GoTo Cancelar2
    End If
Next Libro

Cancelar2:
Hoja = False
ThisWorkbook.Save
End Sub


Sub Time()

UserForm3.Show
MsgBox CajaTexto

End Sub

Sub Verificacion()
Dim i As Long
Dim C As Long

ThisWorkbook.Activate
If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Interior.ColorIndex <> xlNone And _
    Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia).Interior.ColorIndex <> xlNone Then
    GoTo Cancelar
ElseIf IsEmpty(Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia)) = False Then
    GoTo Cancelar
End If

For i = 2 To 10
    If RefHora2 = i Then
        For A = i - 1 To 1 Step -1
            ThisWorkbook.Activate
            If IsEmpty(Worksheets("BITACORA").Cells(FilaDia - A, ColumnaDia)) = True Then
                FilaDia = FilaDia - A
                'Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                CajaTexto = ""
                Ejecutando2 = True
                Call MACRO_1.TiempoEspera
                UserForm3.Show
                If CANCELADO = True Then Exit Sub
                Workbooks("GESTIÓN.xls").Activate
                Worksheets("BITACORA").Activate
                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    Selection.Interior.ColorIndex = xlNone
                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) = CajaTexto
                If IsEmpty(Comentario) = False And Comentario <> "" Then
                    If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).AddComment
                    End If
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Visible = False
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text Text:=Comentario
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
                End If
                Comentario = ""
                CajaTexto = ""
                Call DiaHora
                If A = 1 Then CANCELADO = True
                GoTo Cancelar:
            End If
        Next A
    End If
Next i

If Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia).Interior.ColorIndex = xlNone And _
    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Interior.ColorIndex <> xlNone Then
    C = 36
Else
    C = 22
End If

For i = 14 To C ' 14
    If RefHora2 = i Then
        For A = i To 15 Step -1
            ThisWorkbook.Activate
            If IsEmpty(Worksheets("BITACORA").Cells(FilaDia - A + 14, ColumnaDia)) = True Then
                FilaDia = FilaDia - A + 14
                'Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                CajaTexto = ""
                Comentario = ""
                Ejecutando2 = True
                Call MACRO_1.TiempoEspera
                UserForm3.Show
                If CANCELADO = True Then Exit Sub
                Workbooks("GESTIÓN.xls").Activate
                Worksheets("BITACORA").Activate
                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    Selection.Interior.ColorIndex = xlNone
                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) = CajaTexto
                If IsEmpty(Comentario) = False And Comentario <> "" Then
                    If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).AddComment
                    End If
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Visible = False
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text Text:=Comentario
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
                    Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
                End If
                Comentario = ""
                CajaTexto = ""
                Call DiaHora
                If A = 15 Then CANCELADO = True
                GoTo Cancelar:
            End If
        Next A
    End If
Next i

Cancelar:

End Sub

Sub TiempoEspera()
Tiempo2 = Now + TimeValue("00:05:00")
Application.OnTime Tiempo2, "Cancelar", , Ejecutando2
End Sub

Sub Cancelar()
CajaTexto = ""
Comentario = ""
UserForm3.Hide
Ejecutando2 = False
CANCELADO = True
End Sub

Sub ActividadPendiente()
Dim UltimaFila As Long
Dim Rango As Range
Dim CELDA As Range

UltimaFila = Worksheets("SOL INDICADORES").Range("G" & Rows.Count).End(xlUp).Row
Set Rango = Worksheets("SOL INDICADORES").Range("G2:G" & UltimaFila)
For Each CELDA In Rango
    If Dia = CELDA Then
        UserForm4.Label1 = "PENDIENTE PARA HOY:" & vbCrLf & CELDA.Offset(0, -5)
        UserForm4.Show
    End If
    If CELDA.Offset(0, 3) = "PENDIENTE" And Dia > CELDA Then
        UserForm4.Label1 = "PENDIENTE PARA HOY:" & vbCrLf & CELDA.Offset(0, -5)
        UserForm4.Show
    End If
Next CELDA

End Sub

Sub ListaComentarios()

Dim UlFila, FilaSup, FilaInf, UlColumna As Long
Dim RefRang, RefCel As Range
Dim FeriadosRang, FeriadosCelda As Range
Dim ActvidadRang, ActividadCelda As Range
Dim GraficoItem, GraficoHoras As Range
Dim MES As String
Dim AÑO, TemMes, i, A, B, D As Long
Dim MatrizActividades() As String
Dim TempMatriz() As String
Dim RangoMes, CeldaMes As Range
Dim RangoDiasMes, CeldaDiasMes As Range
Dim RangoTrabajo, Celdatrabajo As Range
Dim HoraMes As Long
Dim HoraTrabajo As Single
Dim RangoLista, CeldaLista As Range

Erase TempMatriz
Erase MatrizActividades
UserForm3.ComboBox2.Clear

AÑO = Format(Date, "yyyy")
MES = UCase(Format(Date, "MMMM"))

ThisWorkbook.Worksheets("Solicitudes").Activate

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
            If ActividadCelda.Value <> UserForm3.ComboBox1.Value Then GoTo SIG1
            If IsEmpty(i) Then
                i = 0
                ReDim Preserve MatrizActividades(0, i)
                MatrizActividades(0, i) = ActividadCelda.Comment.Text 'Comentario
                GoTo SIG2
            End If
            For B = 0 To A
                If ActividadCelda.Comment.Text = MatrizActividades(0, B) Then
                    GoTo SIG1
                End If
            Next B
            
            ReDim Preserve MatrizActividades(0, i)
            MatrizActividades(0, i) = ActividadCelda.Comment.Text 'Comentario
SIG2:
            A = i
            i = i + 1
SIG1:
        Next ActividadCelda
        Next D
        
    End If
Next RefCel

'ULTIMA FILA
If IsEmpty(A) = False Then
    ReDim Preserve TempMatriz(A, 0)
    For B = 0 To A
        TempMatriz(B, 0) = MatrizActividades(0, B)
    Next B
    UserForm3.ComboBox2.ListRows = A + 1
    UserForm3.ComboBox2.List = TempMatriz
End If

End Sub

Sub EnCelda()
Attribute EnCelda.VB_Description = "Actividad en Celda"
Attribute EnCelda.VB_ProcData.VB_Invoke_Func = "s\n14"

FilaDia = ActiveCell.Row
ColumnaDia = ActiveCell.Column
CajaTexto = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia)
UserForm3.Show
Worksheets("BITACORA").Activate
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Interior.ColorIndex = xlNone
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) = CajaTexto
If IsEmpty(Comentario) = False And Comentario <> "" Then
    If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).AddComment
    End If
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Visible = False
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text Text:=Comentario
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
End If

End Sub



Sub DiaHoraMomento()
Attribute DiaHoraMomento.VB_ProcData.VB_Invoke_Func = "d\n14"
Dim UltimaFila, FilaPrincipal As Long
Dim Rango, CELDA, RangoColumna, CeldaColumna, RangoFila, CeldaFila As Range
Dim AñoMes As String
Dim HoraReferenciaMenor As Date
Dim HoraReferenciaMayor As Date
Dim RefHora As String

Ejecutando2 = False
CANCELADO = False
If Hoja = False Then
    LibroActivo = ActiveWorkbook.Name
    HojaActiva = ActiveSheet.Name
    Hoja = True
End If

If UserForm3.Visible = True Then GoTo Cancelar

NombreLibro = ThisWorkbook.Name
Windows(NombreLibro).Activate
ThisWorkbook.Activate

UltimaFila = Worksheets("BITACORA").Range("A" & Rows.Count).End(xlUp).Row
Set Rango = Worksheets("BITACORA").Range("A1:A" & UltimaFila)
AñoMes = Format(Date, "yyyy") & "-" & UCase(Format(Date, "MMMM"))

For Each CELDA In Rango
    If CELDA = AñoMes Then
        FilaPrincipal = CELDA.Row
        UltimaColumna = Worksheets("BITACORA").Cells(FilaPrincipal, Cells.Columns.Count).End(xlToLeft).Column
        Worksheets("BITACORA").Activate
        Set RangoColumna = Worksheets("BITACORA").Range(Cells(FilaPrincipal, 2), Cells(FilaPrincipal, UltimaColumna))
        For Each CeldaColumna In RangoColumna
            If CeldaColumna = Date Then
                ColumnaDia = CeldaColumna.Column
                Hora = Format(Hour(Now), "##") & ":" & Format(Minute(Now), "##")
                
                HoraReferenciaMenor = "06:30"
                HoraReferenciaMayor = "07:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 0
                    RefHora2 = 1
                End If
                HoraReferenciaMenor = "07:00"
                HoraReferenciaMayor = "07:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 1
                    RefHora2 = 1
                End If
                HoraReferenciaMenor = "07:30"
                HoraReferenciaMayor = "08:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 2
                    RefHora2 = 2
                End If
                HoraReferenciaMenor = "08:00"
                HoraReferenciaMayor = "08:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 3
                    RefHora2 = 3
                End If
                HoraReferenciaMenor = "08:30"
                HoraReferenciaMayor = "09:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 4
                    RefHora2 = 4
                End If
                HoraReferenciaMenor = "09:00"
                HoraReferenciaMayor = "09:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 5
                    RefHora2 = 5
                End If
                HoraReferenciaMenor = "09:30"
                HoraReferenciaMayor = "10:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 6
                    RefHora2 = 6
                End If
                HoraReferenciaMenor = "10:00"
                HoraReferenciaMayor = "10:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 7
                    RefHora2 = 7
                End If
                HoraReferenciaMenor = "10:30"
                HoraReferenciaMayor = "11:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 8
                    RefHora2 = 8
                End If
                HoraReferenciaMenor = "11:00"
                HoraReferenciaMayor = "11:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 9
                    RefHora2 = 9
                End If
                HoraReferenciaMenor = "11:30"
                HoraReferenciaMayor = "12:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 10
                    RefHora2 = 10
                End If

                HoraReferenciaMenor = "12:00"
                HoraReferenciaMayor = "12:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 11
                    RefHora2 = 11
                End If
                HoraReferenciaMenor = "12:30"
                HoraReferenciaMayor = "13:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 12
                    RefHora2 = 12
                End If
                HoraReferenciaMenor = "13:00"
                HoraReferenciaMayor = "13:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 13
                    RefHora2 = 13
                End If
                HoraReferenciaMenor = "13:30"   '13:30
                HoraReferenciaMayor = "14:00"   '14:00
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 14
                    RefHora2 = 14
                End If
                HoraReferenciaMenor = "14:00"   '14:00
                HoraReferenciaMayor = "14:30"   '14:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 15
                    RefHora2 = 15
                End If
                HoraReferenciaMenor = "14:30"   '14:30
                HoraReferenciaMayor = "15:00"   '15:0
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 16
                    RefHora2 = 16
                End If
                HoraReferenciaMenor = "15:00"   '15:00
                HoraReferenciaMayor = "15:30"   '15:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 17
                    RefHora2 = 17
                End If
                HoraReferenciaMenor = "15:30"   '15:30
                HoraReferenciaMayor = "16:00"   '16:00
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 18
                    RefHora2 = 18
                End If
                HoraReferenciaMenor = "16:00"   '16:00
                HoraReferenciaMayor = "16:30"   '16:30
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 19
                    RefHora2 = 19
                End If
                HoraReferenciaMenor = "16:30"
                HoraReferenciaMayor = "17:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 20
                    RefHora2 = 20
                End If
                HoraReferenciaMenor = "17:00"
                HoraReferenciaMayor = "17:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 21
                    RefHora2 = 21
                End If
                HoraReferenciaMenor = "17:30"
                HoraReferenciaMayor = "18:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 22
                    RefHora2 = 22
                End If
                HoraReferenciaMenor = "18:00"
                HoraReferenciaMayor = "18:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 23
                    RefHora2 = 23
                End If
                HoraReferenciaMenor = "18:30"
                HoraReferenciaMayor = "19:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 24
                    RefHora2 = 24
                End If
                HoraReferenciaMenor = "19:00"
                HoraReferenciaMayor = "19:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 25
                    RefHora2 = 25
                End If
                HoraReferenciaMenor = "19:30"
                HoraReferenciaMayor = "20:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 26
                    RefHora2 = 26
                End If
                HoraReferenciaMenor = "20:00"
                HoraReferenciaMayor = "20:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 27
                    RefHora2 = 27
                End If
                HoraReferenciaMenor = "20:30"
                HoraReferenciaMayor = "21:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 28
                    RefHora2 = 28
                End If
                HoraReferenciaMenor = "21:00"
                HoraReferenciaMayor = "21:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 29
                    RefHora2 = 29
                End If
                HoraReferenciaMenor = "21:30"
                HoraReferenciaMayor = "22:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 30
                    RefHora2 = 30
                End If
                HoraReferenciaMenor = "22:00"
                HoraReferenciaMayor = "22:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 31
                    RefHora2 = 31
                End If
                HoraReferenciaMenor = "22:30"
                HoraReferenciaMayor = "23:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 32
                    RefHora2 = 32
                End If
                HoraReferenciaMenor = "23:00"
                HoraReferenciaMayor = "23:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 33
                    RefHora2 = 33
                End If
                HoraReferenciaMenor = "23:30"
                HoraReferenciaMayor = "23:59"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 34
                    RefHora2 = 34
                End If
                HoraReferenciaMenor = "00:00"
                HoraReferenciaMayor = "00:30"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 35
                    RefHora2 = 35
                End If
                HoraReferenciaMenor = "00:30"
                HoraReferenciaMayor = "01:00"
                If Hora > HoraReferenciaMenor And Hora < HoraReferenciaMayor Then
                    RefHora = AñoMes & "-" & 36
                    RefHora2 = 36
                End If
                Set RangoFila = Worksheets("BITACORA").Range("A1:A" & UltimaFila)
                If RefHora = "" Then GoTo Cancelar2
                For Each CeldaFila In Rango
                    If CeldaFila = RefHora Then
                        FilaDia = CeldaFila.Row
                        If CANCELADO = True Then Exit Sub
                        Call Verificacion
                        If CANCELADO = True Then Exit Sub
                        CajaTexto = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia)
                        If Not Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                            Comentario = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text
                        End If
                        
                        
                        Ejecutando2 = True
                        Call MACRO_1.TiempoEspera
                        UserForm3.Show
                        Call Verificacion
                        If CANCELADO = True Then Exit Sub
                        Workbooks("GESTIÓN.xls").Activate
                        Worksheets("BITACORA").Activate
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Select
                                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                                Selection.Interior.ColorIndex = xlNone
                        Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) = CajaTexto
                        If IsEmpty(Comentario) = False And Comentario <> "" Then
                            If Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment Is Nothing Then
                                Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).AddComment
                            End If
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Visible = False
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Text Text:=Comentario
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
                            Worksheets("BITACORA").Cells(FilaDia, ColumnaDia).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
                        End If
                        GoTo Cancelar
                    End If
                Next CeldaFila
            End If
        Next CeldaColumna
    End If
Next CELDA

If RefHora2 = 6 Or RefHora2 = 14 Or RefHora2 = 17 Then
    Call ActividadPendiente
End If
Cancelar:
For Each Libro In Workbooks
    If Libro.Name = LibroActivo Then
        Workbooks(LibroActivo).Worksheets(HojaActiva).Activate
        GoTo Cancelar2
    End If
Next Libro

Cancelar2:
Hoja = False
ThisWorkbook.Save
End Sub

