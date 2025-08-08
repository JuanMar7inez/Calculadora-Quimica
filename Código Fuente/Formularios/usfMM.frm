VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfMM 
   Caption         =   "Peso Moleculares Elementos"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "usfMM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim FormX As Double, FormY As Double
Dim numero As String
Dim numeroDatos As Variant
Dim clear As Variant
Dim y As Variant
Dim fila As Double
Dim descrip As String

Private Sub cmdElegir_Click()
    numero = lbx1.Text
    Me.lbl1.Caption = numero
End Sub

Private Sub cmdLimpiar_Click()
    Me.lbl1.Caption = ""
    Me.txtEscribir.Text = ""
    Me.txtEscribir.SetFocus
End Sub

Private Sub cmdVolver_Click()
    usfMM.Hide
End Sub

Private Sub CommandButton1_Click()
    MsgBox "¡Bienvenido/a a la tabla de Pesos Molares!. Sólo tienes que escribir el símbolo del elemento que buscas en el cuadro de texto, luego seleccionarlo con el mouse y posteriormente apretar el botón Elegir. Si quieres seleccionar otro elemento, puedes limpiar los formularios con el Botón Limpiar. Para ir al Menú Principal, aprieta el botón Salir. "
End Sub

Private Sub txtEscribir_Change()
   numeroDatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
    Hoja1.AutoFilterMode = False
    Me.lbx1 = clear
    Me.lbx1.RowSource = clear
    y = 0
    For fila = 4 To numeroDatos
    descrip = Hoja1.Cells(fila, 2).Value

    If UCase(descrip) Like "" & UCase(txtEscribir.Value) & "" Then
    Me.lbx1.AddItem
    Me.lbx1.List(y, 0) = Hoja1.Cells(fila, 1).Value
    Me.lbx1.List(y, 1) = Hoja1.Cells(fila, 2).Value
    Me.lbx1.List(y, 2) = Hoja1.Cells(fila, 3).Value
    Me.lbx1.List(y, 3) = Hoja1.Cells(fila, 4).Value

    y = y + 1

    End If
Next
End Sub

Private Sub UserForm_Activate()
    Me.lbx1.RowSource = "Tabla1"
    Me.lbx1.ColumnCount = 3

End Sub
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
If Button = 1 Then
FormX = X
FormY = y
End If
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
If Button = 1 Then
Me.Left = Me.Left + (X - FormX)
Me.Top = Me.Top + (y - FormY)
End If
End Sub




Private Sub UserForm_Initialize()
Me.txtEscribir.SetFocus
Dim lngWindow As Long, lFrmHdl As Long
lFrmHdl = FindWindowA(vbNullString, Me.Caption)
lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
lngWindow = lngWindow And (Not WS_CAPTION)
Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
Call DrawMenuBar(lFrmHdl)
End Sub
