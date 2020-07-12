VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton2_Click()
Call Cancelar
End Sub

Private Sub CommandButton3_Click()
Dim UltimaFila, i, A As Long
Dim Nombre As String
Dim CeldaTopico, TopicoRang As Range
Dim MatrizTopico() As String
Dim TemMatrizTopico() As String

UserForm6.Show

'COMBOLIST TOPICO

UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, TOPICOS).End(xlUp).Row
Set TopicoRang = Worksheets("Solicitudes").Range(Cells(2, TOPICOS), Cells(UltimaFila, TOPICOS))

For Each CeldaTopico In TopicoRang
        If IsEmpty(i) Then i = 0
            ReDim Preserve MatrizTopico(0, i)
            MatrizTopico(0, i) = CeldaTopico.Value 'TOPICO
        A = i
        i = i + 1

Next CeldaTopico

    For i = 0 To A
        ReDim Preserve TemMatrizTopico(A, 0)
        TemMatrizTopico(i, 0) = MatrizTopico(0, i)
    Next i

MatrizTopico = TemMatrizTopico

ComboBox1.ColumnCount = A + 1
ComboBox1.List = MatrizTopico

End Sub

Private Sub UserForm_Activate()
Dim TopicoRang, CeldaTopico As Range
Dim DestinoRang, CeldaDestino As Range
Dim UltimaFila, A, i As Long
Dim MatrizTopico() As String
Dim TemMatrizTopico() As String
Dim MatrizDestino() As String
Dim TemMatrizDestino() As String

Worksheets("Solicitudes").Activate

DTPicker1 = Date
DTPicker2 = Date


'QUITAR FILTROS
If Worksheets("Solicitudes").FilterMode = True Then
    Worksheets("Solicitudes").ShowAllData
End If

'COMBOLIST TOPICO

UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, TOPICOS).End(xlUp).Row
Set TopicoRang = Worksheets("Solicitudes").Range(Cells(2, TOPICOS), Cells(UltimaFila, TOPICOS))

For Each CeldaTopico In TopicoRang
        If IsEmpty(i) Then i = 0
            ReDim Preserve MatrizTopico(0, i)
            MatrizTopico(0, i) = CeldaTopico.Value 'TOPICO
        A = i
        i = i + 1

Next CeldaTopico

    For i = 0 To A
        ReDim Preserve TemMatrizTopico(A, 0)
        TemMatrizTopico(i, 0) = MatrizTopico(0, i)
    Next i

MatrizTopico = TemMatrizTopico

ComboBox1.ColumnCount = A + 1
ComboBox1.List = MatrizTopico

'COMBOLIST DESTINO
A = 0
i = 0
UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, DESTINOS).End(xlUp).Row
Set DestinoRang = Worksheets("Solicitudes").Range(Cells(2, DESTINOS), Cells(UltimaFila, DESTINOS))

For Each CeldaDestino In DestinoRang
        If IsEmpty(i) Then i = 0
            ReDim Preserve MatrizDestino(0, i)
            MatrizDestino(0, i) = CeldaDestino.Value 'DESTINO
        A = i
        i = i + 1

Next CeldaDestino

    For i = 0 To A
        ReDim Preserve TemMatrizDestino(A, 0)
        TemMatrizDestino(i, 0) = MatrizDestino(0, i)
    Next i

MatrizDestino = TemMatrizDestino

ComboBox2.ColumnCount = A + 1
ComboBox2.List = MatrizDestino

End Sub

Private Sub CommandButton1_Click()
Dim UltimaFila As Long


If TextBox2 = "" Or TextBox3 = "" Or TextBox8 = "" Or ComboBox1 = "Seleccionar" Or ComboBox2 = "Seleccionar" Then
    MsgBox "Por favor completar los campos obligatorios"
    GoTo SIG1
End If

Worksheets("Solicitudes").Activate

UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 2).End(xlUp).Row + 1

'FORMULA MES ATENDIDO
Worksheets("Solicitudes").Cells(UltimaFila, AtendidoMes).FormulaLocal = "=SI(P" & UltimaFila & ">E" & UltimaFila & ";D" & UltimaFila & ";C" & UltimaFila & ")"

'FORMULA MES SOLICITADO
Worksheets("Solicitudes").Cells(UltimaFila, MesSolicitado).FormulaLocal = "=SI(E" & UltimaFila & "="""";"""";AÑO(E" & UltimaFila & ")&""-""&MES(E" & UltimaFila & "))"

'FORMULA MES REPUESTA
Worksheets("Solicitudes").Cells(UltimaFila, MesRepuesta).FormulaLocal = "=SI(P" & UltimaFila & "="""";"""";AÑO(P" & UltimaFila & ")&""-""&MES(P" & UltimaFila & "))"

'FORMULA TIEMPO DE MORA
Worksheets("Solicitudes").Cells(UltimaFila, TiempoMora).FormulaLocal = "=SI(P" & UltimaFila & "="""";SI(O" & UltimaFila & "="""";"""";DIAS.LAB(E" & UltimaFila & ";O" & UltimaFila & ";FERIADOS));SI(P" & UltimaFila & "="""";"""";DIAS.LAB(E" & UltimaFila & ";P" & UltimaFila & ";FERIADOS)))"

'FECHA DE SOLICITUD
Worksheets("Solicitudes").Cells(UltimaFila, FechaSolicitud) = DTPicker1

'SOLICITANTE
Worksheets("Solicitudes").Cells(UltimaFila, Solicitante) = TextBox2

'TITULO
Worksheets("Solicitudes").Cells(UltimaFila, Titulo) = TextBox3

'DOCUMENTO
Worksheets("Solicitudes").Cells(UltimaFila, Documento) = TextBox4
Worksheets("Solicitudes").Cells(UltimaFila, Documento).NumberFormat = "0"

'TOPICO
Worksheets("Solicitudes").Cells(UltimaFila, TOPICO) = ComboBox1

'DESTINO
Worksheets("Solicitudes").Cells(UltimaFila, DESTINO) = ComboBox2

'STATUS
Worksheets("Solicitudes").Cells(UltimaFila, Status) = "PENDIENTE"

'FECHA DE ENVIO CORREO
Worksheets("Solicitudes").Cells(UltimaFila, FechaEnvioCorreo) = DTPicker2

'SOLICITUD
Worksheets("Solicitudes").Cells(UltimaFila, TextoSolicitud) = TextBox8.Text

'OBSERVACIONES
Worksheets("Solicitudes").Cells(UltimaFila, Observaciones) = TextBox9

Worksheets("Solicitudes").Rows(UltimaFila - 1).Select
Selection.Copy
Worksheets("Solicitudes").Rows(UltimaFila).Select
Selection.PasteSpecial Paste:=xlPasteFormats
Selection.PasteSpecial Paste:=xlPasteValidation
Application.CutCopyMode = False

Worksheets("Solicitudes").Range("A" & UltimaFila).Select

Call AsignarNumero
Call Cancelar

SIG1:
End Sub


Sub Cancelar()

TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
ComboBox1.Clear
ComboBox2.Clear
TextBox8.Value = ""
TextBox9.Value = ""

UserForm5.Hide
End Sub


Sub AsignarNumero()
Dim UltimaFila, A As Long
Dim NUMERO As String
Dim MesSoliRang, MesSoliCel As Range

UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, 2).End(xlUp).Row
Set MesSoliRang = Worksheets("Solicitudes").Range(Cells(2, MesSolicitado), Cells(UltimaFila - 1, MesSolicitado))

A = 0

For Each MesSoliCel In MesSoliRang
    If MesSoliCel.Value = Cells(UltimaFila, MesSolicitado) Then
     A = A + 1
    End If
Next MesSoliCel

A = A + 1

NUMERO = "SI-" & Format(DTPicker1, "YY") & Format(DTPicker1, "MM") & Format(DTPicker1, "DD") & "-" & A
Worksheets("Solicitudes").Cells(UltimaFila, 1) = NUMERO

End Sub
