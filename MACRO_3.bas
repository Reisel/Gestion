Attribute VB_Name = "MACRO_3"
Option Explicit

Public NCaso, AtendidoMes, MesSolicitado, MesRepuesta, FechaSolicitud, Solicitante, Titulo, Documento, _
    TOPICO, DESTINO, Ur, Status, FechaEnvioCorreo, FechaRecibido, FechaRepuesta, TextoSolicitud, _
    TextoRepuesta, Recordatorio, TiempoMora, Observaciones, TOPICOS, DESTINOS, FechaRepuestaUsuario As Long

Sub BotonCopiar()
Dim Answer As Integer
Answer = MsgBox("¿Desea copiar en las carpetas publicas?", vbYesNo)
If Answer = vbNo Then GoTo CANCELAR11

ThisWorkbook.Worksheets("Solicitudes").Activate
If ThisWorkbook.Worksheets("Solicitudes").FilterMode = True Then
    Worksheets("Solicitudes").ShowAllData
End If

'Call CopiaPublica
'Call InsertarHojas
Call CopiaPublica2
Call InsertarHojas

CANCELAR11:
End Sub


Sub CopiaPublica()
Attribute CopiaPublica.VB_Description = "Macro grabada el 22/02/2019 por Administrador"
Attribute CopiaPublica.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Libro As Workbook
On Error Resume Next


'ABRIR GESTIÓN INF *** INICIO***
For Each Libro In Workbooks
    If Libro.Name = "GESTIÓN INF.xls" Then
        GoTo AAA
    End If
Next Libro
        
    Workbooks.Open Filename:= _
    "H:\GESTION\02 SISTEMA INFORMACIÓN\GESTIÓN INF.xls", WriteResPassword:="BARIVEN"

BBB:
For Each Libro In Workbooks
    If Libro.Name = "GESTIÓN INF.xls" Then
        GoTo AAA
    End If
Next Libro
        MsgBox ("NO SE ENCUENTRA ADOCUMENTO ""..GESTIÓN INF.xls.."" EN GESTIÓN (PUBLICO)" & vbNewLine & vbNewLine & _
        "PARA LA EJECUCIÓN DE LA MACRO SE REQUIERE TENER ABIERTO EL ARCHIVO ""..GESTIÓN INF..""")
        GoTo Cancelar
AAA:

    ThisWorkbook.Activate
    
'ABRIR LIBRO DE COMPRADORES *** FINAL***
Cancelar:
End Sub

Sub CopiaPublica2()
Dim Libro As Workbook
On Error Resume Next

'ABRIR GESTIÓN INF *** INICIO***
For Each Libro In Workbooks
    If Libro.Name = "GESTIÓN INF.xls" Then
        GoTo AAA
    End If
Next Libro
     
    Workbooks.Open Filename:= _
    "H:\INFORME GESTION\05 SISTEMA INFORMACIÓN\GESTIÓN INF.xls", WriteResPassword:="BARIVEN"

BBB:
For Each Libro In Workbooks
    If Libro.Name = "GESTIÓN INF.xls" Then
        GoTo AAA
    End If
Next Libro
        MsgBox ("NO SE ENCUENTRA EL DOCUMENTO ""..GESTIÓN INF.xls.."" EN INFORME GESTION (PUBLICO)" & vbNewLine & vbNewLine & _
        "PARA LA EJECUCIÓN DE LA MACRO SE REQUIERE TENER ABIERTO EL ARCHIVO ""..GESTIÓN INF..""")
        GoTo Cancelar
AAA:

    ThisWorkbook.Activate
    
'ABRIR LIBRO DE COMPRADORES *** FINAL***
Cancelar:
End Sub

Sub InsertarHojas()
On Error Resume Next

'***INSERTAR HOJA DE INDICADORES****
Workbooks("GESTIÓN INF").Worksheets("INDICADORES").Delete
Workbooks("GESTIÓN").Activate
Workbooks("GESTIÓN").Worksheets("INDICADORES").Select
Workbooks("GESTIÓN").Worksheets("INDICADORES").Copy Before:=Workbooks("GESTIÓN INF.xls").Sheets(1)
Workbooks("GESTIÓN INF.xls").Worksheets("INDICADORES").Activate
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
Workbooks("GESTIÓN INF.xls").Worksheets("INDICADORES").Cells(1, 1).Select

'***INSERTAR HOJA DE SEGUIMIENTO****

Workbooks("GESTIÓN INF").Worksheets("Seguimiento").Delete
Workbooks("GESTIÓN").Activate
Workbooks("GESTIÓN").Worksheets("Seguimiento").Select
Workbooks("GESTIÓN").Worksheets("Seguimiento").Copy Before:=Workbooks("GESTIÓN INF.xls").Sheets(1)
Workbooks("GESTIÓN INF.xls").Worksheets("Seguimiento").Activate
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
Workbooks("GESTIÓN INF.xls").Worksheets("Seguimiento").Cells(1, 1).Select

'***INSERTAR HOJA DE SOLICITUDES****

Workbooks("GESTIÓN INF").Worksheets("Solicitudes").Delete
Workbooks("GESTIÓN").Activate
Workbooks("GESTIÓN").Worksheets("Solicitudes").Select
Workbooks("GESTIÓN").Worksheets("Solicitudes").Copy Before:=Workbooks("GESTIÓN INF.xls").Sheets(1)
Workbooks("GESTIÓN INF.xls").Worksheets("Solicitudes").Activate
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
Workbooks("GESTIÓN INF.xls").Worksheets("Solicitudes").Cells(1, 1).Select

Workbooks("GESTIÓN INF").Save
Workbooks("GESTIÓN INF").Close

ThisWorkbook.Activate

End Sub

Sub Formulario()
Attribute Formulario.VB_Description = "Llamar Formulario"
Attribute Formulario.VB_ProcData.VB_Invoke_Func = "f\n14"

Workbooks("GESTIÓN").Activate
Call VariablesCabecera
UserForm5.Show

End Sub

Sub VariablesCabecera()
Dim UltimaColumnaSol As Long
Dim Cabecera, CeldaCabecera As Range

Worksheets("Solicitudes").Activate
UltimaColumnaSol = Worksheets("Solicitudes").Cells(1, Cells.Columns.Count).End(xlToLeft).Column
Set Cabecera = Worksheets("Solicitudes").Range(Cells(1, 1), Cells(1, UltimaColumnaSol))
For Each CeldaCabecera In Cabecera
    If CeldaCabecera = "N° CASO" Then
       NCaso = CeldaCabecera.Column
       GoTo SIG2
    End If
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
    If CeldaCabecera = "DESTINO" Then
        DESTINO = CeldaCabecera.Column
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
    If CeldaCabecera = "TOPICOS" Then
        TOPICOS = CeldaCabecera.Column
        GoTo SIG2
    End If
    If CeldaCabecera = "DESTINOS" Then
        DESTINOS = CeldaCabecera.Column
        GoTo SIG2
    End If
SIG2:
Next CeldaCabecera

End Sub

Sub EliminarFormulas()
Dim StatusCel, StatusRang As Range
Dim UltimaFila, Fila As Long

Call VariablesCabecera

Worksheets("Solicitudes").Activate
UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, Status).End(xlUp).Row
Set StatusRang = Worksheets("Solicitudes").Range(Cells(2, Status), Cells(UltimaFila, Status))

For Each StatusCel In StatusRang
    If StatusCel = "LISTO" And StatusCel.Offset(0, 3) > Date - 30 Then
        Fila = StatusCel.Row
        Worksheets("Solicitudes").Rows(Fila).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
Next StatusCel

Worksheets("Solicitudes").Cells(UltimaFila + 1, FechaSolicitud).Select

End Sub

