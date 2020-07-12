VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim Nombre As String
Dim UltimaFila As Long

Nombre = UCase(TextBox1)

UltimaFila = Worksheets("Solicitudes").Cells(Cells.Rows.Count, TOPICOS).End(xlUp).Row + 1
Worksheets("Solicitudes").Cells(UltimaFila, TOPICOS) = Nombre
Worksheets("Solicitudes").Cells(UltimaFila - 1, TOPICOS).Select
Selection.Copy
Worksheets("Solicitudes").Cells(UltimaFila, TOPICOS).Select
Selection.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

Worksheets("Solicitudes").Range(Cells(2, TOPICOS), Cells(UltimaFila, TOPICOS)).Select
    Selection.Sort Key1:=Worksheets("Solicitudes").Range(Cells(1, TOPICOS), Cells(UltimaFila, TOPICOS)), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers

TextBox1 = ""
UserForm6.Hide

End Sub

