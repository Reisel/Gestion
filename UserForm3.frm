VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Descripción de la Actividad"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ComentarioAnt As String


Private Sub ComboBox1_Change()
If ComboBox1.Value = "SELECCIONAR" Then
    CommandButton2.Enabled = False
Else
    CommandButton2.Enabled = True
    
    TextBox1.Visible = False
    ComboBox2.Visible = True
    CommandButton6.Visible = True

    Call MACRO_1.ListaComentarios
    
    If ComboBox1.Value = Label1.Caption Then
        CommandButton1.Enabled = True
    Else
        CommandButton1.Enabled = False
        If ComboBox2.ListCount = 0 Then
            ComboBox2.Value = ""
        Else
            ComboBox2.Value = ComboBox2.List(ComboBox2.ListCount - 1)
        End If
    End If

End If

End Sub

Private Sub CommandButton1_Click()
CajaTexto = Label1.Caption
Comentario = ComentarioAnt
UserForm3.Hide
Ejecutando2 = False
End Sub

Private Sub CommandButton2_Click()
CajaTexto = ComboBox1.Value
If ComboBox2.Visible = True Then
    TextBox1 = ComboBox2.Value
    Comentario = TextBox1
Else
    Comentario = TextBox1
End If

UserForm3.Hide
Ejecutando2 = False
End Sub

Private Sub CommandButton3_Click()
Workbooks(LibroActivo).Worksheets(HojaActiva).Activate
Call MACRO_1.Cancelar
Ejecutando2 = False

End Sub

Private Sub CommandButton4_Click()
TextBox1.Visible = False
ComboBox2.Visible = True
CommandButton6.Visible = True

Call MACRO_1.ListaComentarios

ComboBox1.Value = Label1.Caption
ComboBox2.Value = Label5.Caption

End Sub

Private Sub CommandButton5_Click()
Dim UltimaFila, i, A As Long
Dim Nombre As String
Dim CeldaActividad, ActividadRang As Range
Dim MatrizActividad() As String
Dim TemMatrizActividad() As String

UserForm7.Show

'COMBOLIST TOPICO

UltimaFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, 46).End(xlUp).Row
Set ActividadRang = Worksheets("INDICADORES").Range(Cells(6, 46), Cells(UltimaFila, 46))

For Each CeldaActividad In ActividadRang
        If IsEmpty(i) Then i = 0
            ReDim Preserve MatrizActividad(0, i)
            MatrizActividad(0, i) = CeldaActividad.Value 'TOPICO
        A = i
        i = i + 1

Next CeldaActividad

    For i = 0 To A
        ReDim Preserve TemMatrizActividad(A, 0)
        TemMatrizActividad(i, 0) = MatrizActividad(0, i)
    Next i

MatrizActividad = TemMatrizActividad

ComboBox1.ColumnCount = A + 1
ComboBox1.List = MatrizActividad

End Sub

Private Sub CommandButton6_Click()
ComboBox2.Value = ""

End Sub

Private Sub UserForm_Activate()
Dim UltFila As Long
Dim ActividadesCelda, ActividadesPrinRang As Range
Dim A, i As Long
Dim MatrizActividad() As String
Dim TemMatrizActividad() As String
Application.ScreenUpdating = False

TextBox1.Visible = True
ComboBox2.Visible = False
CommandButton6.Visible = False

TextBox1 = ""
Label5 = ""

ThisWorkbook.Activate

UserForm3.Caption = "Descripción de la Actividad de las .." & Worksheets("BITACORA").Cells(FilaDia, 2)

'TEXTO DE ACTIVIDAD ANTERIOR
If Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia).Interior.ColorIndex <> xlNone Then
    Label1.Caption = Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia)
    CommandButton1.Enabled = True
    If Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia).Comment Is Nothing Then
        ComentarioAnt = ""
    Else
        ComentarioAnt = Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia).Comment.Text
        Label5.Caption = ComentarioAnt
    End If
Else
    Label1.Caption = Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia)
    If Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia).Comment Is Nothing Then
        ComentarioAnt = ""
    Else
        ComentarioAnt = Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia).Comment.Text
        Label5.Caption = ComentarioAnt
    End If
End If

If Label1.Caption = "LUNES" Or Label1.Caption = "MARTES" Or Label1.Caption = "MIERCOLES" _
    Or Label1.Caption = "JUEVES" Or Label1.Caption = "VIERNES" Or Label1.Caption = "SABADO" _
    Or Label1.Caption = "DOMINGO" Then
    Label1.Caption = ""
End If

If RefHora2 <> "" Then
If RefHora2 = 14 And CajaTexto = "" And Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia) <> "" Then
    CommandButton1.Enabled = True
    Label1.Caption = Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia)
    If Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia).Comment Is Nothing Then
        ComentarioAnt = ""
    Else
        Label5.Caption = Worksheets("BITACORA").Cells(FilaDia - 5, ColumnaDia).Comment.Text
    End If
End If
End If

'COMBOLIST ACTIVIDAD PRINCIPAL

UltFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, 46).End(xlUp).Row
Worksheets("INDICADORES").Activate
Worksheets("INDICADORES").Cells(6, 46).Select
Set ActividadesPrinRang = Worksheets("INDICADORES").Range(Cells(6, 46), Cells(UltFila, 46))

For Each ActividadesCelda In ActividadesPrinRang
        If IsEmpty(i) Then i = 0
            ReDim Preserve MatrizActividad(0, i)
            MatrizActividad(0, i) = ActividadesCelda.Value 'ACTIVIDAD
        A = i
        i = i + 1
Next ActividadesCelda
    
    ReDim Preserve TemMatrizActividad(A, 0)
    For i = 0 To A
        TemMatrizActividad(i, 0) = MatrizActividad(0, i)
    Next i

MatrizActividad = TemMatrizActividad

ComboBox1.ColumnCount = A + 1
ComboBox1.List = MatrizActividad
ComboBox1.Value = "SELECCIONAR"

'VALIDACIÓN DE TEXTO EN LA CELDA
If CajaTexto <> "" Then
    ComboBox1 = CajaTexto
    If CajaTexto = Worksheets("BITACORA").Cells(FilaDia, ColumnaDia) Then
        ComboBox2 = Comentario
    Else
        ComboBox2 = ComentarioAnt
    End If
Else
    CommandButton2.Enabled = False
End If

'ACTIVAR BOTON ANTERIOR
If CajaTexto = "" Or Worksheets("BITACORA").Cells(FilaDia - 1, ColumnaDia) = "" Then
    CommandButton1.Enabled = False
Else
    CommandButton1.Enabled = True
End If

'ACTIVAR BOTON ANTERIOR
If Label1.Caption <> "" Then
    CommandButton1.Enabled = True
End If

End Sub




