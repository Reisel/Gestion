VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "UserForm7"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim Nombre As String
Dim UltimaFila As Long
Application.ScreenUpdating = False

Nombre = TextBox1
Workbooks("GESTIÓN.XLS").Worksheets("INDICADORES").Activate
UltimaFila = Worksheets("INDICADORES").Cells(Cells.Rows.Count, 46).End(xlUp).Row + 1
Worksheets("INDICADORES").Cells(UltimaFila, 46) = Nombre
Worksheets("INDICADORES").Cells(UltimaFila - 1, 46).Select
Selection.Copy
Worksheets("INDICADORES").Cells(UltimaFila, 46).Select
Selection.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

Worksheets("INDICADORES").Range(Cells(6, 46), Cells(UltimaFila, 46)).Select
    Selection.Sort Key1:=Worksheets("INDICADORES").Range(Cells(1, 46), Cells(UltimaFila, 46)), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers

UserForm7.Hide
End Sub
