VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfRemdimientoPorcentual 
   Caption         =   "Calculadora de Rendimiento Porcentual"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12405
   OleObjectBlob   =   "usfRemdimientoPorcentual.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfRemdimientoPorcentual"
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

Private Sub cmdCalcular_Click()

Dim Moles1, ReglaDe3, RendimientoFinal As Double

If Me.txtMasa.Text = "" Then
    MsgBox "Debe insertar un valor"
ElseIf Me.txtMolesRl.Text = "" Then
    MsgBox "Debe insertar un valor"
ElseIf Me.txtNumeroE.Text = "" Then
    MsgBox "Debe insertar un valor"
ElseIf Me.txtNumeroERL.Text = "" Then
    MsgBox "Debe insertar un valor"
ElseIf Me.txtPM.Text = "" Then
    MsgBox "Debe insertar un valor"
ElseIf Me.txtRL.Text = "" Then
    MsgBox "Debe insertar un valor"
Else

Moles1 = Me.txtMasa.Value / Me.txtPM.Value
Me.lblMolesResultado.Caption = Format(Moles1, "##,##0.000")

ReglaDe3 = (Me.txtMolesRl.Value * Me.txtNumeroE.Value) / Me.txtNumeroERL.Value



If ReglaDe3 >= Me.lblMolesResultado.Caption Then
    RendimientoFinal = ((Me.lblMolesResultado.Caption) / (ReglaDe3))
    Me.lblResultado.Caption = Format(RendimientoFinal, "##,##0.00%")
ElseIf ReglaDe3 < Me.lblMolesResultado.Caption Then
   RendimientoFinal = ((ReglaDe3) / (Me.lblMolesResultado.Caption))
   Me.lblResultado.Caption = Format(RendimientoFinal, "##,##0.00%")
End If

End If




End Sub

Private Sub cmdInstrucciones_Click()
    MsgBox "¡Bienvenido/a a la calculadora de Rendimiento Porcentual!. Antes de empezar, necesitarás los siguientes datos para operar correctamente: El reactivo limitante de tu reacción, número estequiométrico del Reactivo Limitante, Moles Reactivo Limitante, el coeficiente Estequiométrico, la masa en gramos, y el peso Molecular del producto a calcular. Después de obtener y rellenar con estos datos, debes presionar el botón Calcular para realizar la operación y obtener los moles del producto calculado y su rendimiento porcentual. Para realizar una nueva operación, debes limpiar el formulario con el botón Limpiar. Para regresar al Menú Principal, solo debes presional el botón Salir."
End Sub

Private Sub cmdLimpiar_Click()
    Me.txtMasa.Text = ""
    Me.txtMolesRl.Text = ""
    Me.txtNumeroE.Text = ""
    Me.txtNumeroERL.Text = ""
    Me.txtPM.Text = ""
    Me.txtRL.Text = ""
    Me.lblMolesResultado.Caption = ""
    Me.lblResultado.Caption = ""

    
End Sub

Private Sub cmdMM_Click()
    usfMM.Show
End Sub

Private Sub cmdSalir_Click()
    usfRemdimientoPorcentual.Hide
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
Dim lngWindow As Long, lFrmHdl As Long
lFrmHdl = FindWindowA(vbNullString, Me.Caption)
lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
lngWindow = lngWindow And (Not WS_CAPTION)
Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
Call DrawMenuBar(lFrmHdl)
End Sub
