VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfPureza 
   Caption         =   "UserForm1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310.001
   OleObjectBlob   =   "usfPureza.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfPureza"
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
Dim PorcentajePureza, FactorMolar, Moles, Resultado1, ResultadoFinal As Double

If Me.txtGramos.Text = "" Then
    MsgBox "Debe ingresar datos"
ElseIf Me.txtNEP.Text = "" Then
    MsgBox "Debe ingresar datos"
ElseIf Me.txtNER.Text = "" Then
    MsgBox "Debe ingresar datos"
ElseIf Me.txtPM.Text = "" Then
    MsgBox "Debe ingresar datos"
ElseIf Me.txtPM2.Text = "" Then
    MsgBox "Debe ingresar datos"
ElseIf Me.txtPureza.Text = "" Then
    MsgBox "Debe ingresar datos"




Else
PorcentajePureza = (Me.txtGramos.Value * Me.txtPureza.Value) / 100
FactorMolar = Me.txtNEP.Value / Me.txtNER.Value
Moles = PorcentajePureza / Me.txtPM.Value
Resultado1 = Moles * FactorMolar
Me.lblMoles.Caption = Format(Resultado1, "##,##0.00")
ResultadoFinal = Moles * Me.txtPM2.Value
Me.lblResultado = Format(ResultadoFinal, "##,##0.00")
End If

End Sub

Private Sub CommandButton4_Click()
    usfCalculadoraMyG.Show
End Sub

Private Sub cmdInstrucciones_Click()
    MsgBox "¡Bienvenido/a la calculadora de Pureza!. Para empezar a calcular, necesitarás algunos datos: Gramos del reactante, Porcentaje de Pureza (sólo ingresar el número sin símbolos), Peso molecular del Reactante a Calcular, el coeficiente estequiométrico del Reactante y del producto; y el peso molecular del compuesto o producto a calcular. Después de tener todos estos valores, rellena los campos según corresponda y aprieta el botón Calcular. Si deseas realizar otra operación, utiliza el botón Limpiar y rellena nuevamente. Si necesitas saber el Peso Molecular de un elemento, y no lo conoces, puedes utilizar la tabla Peso Molecular. Para ir al Menú Principal, solo presiona el botón Salir."
End Sub

Private Sub cmdLimpiar_Click()
Me.txtGramos.Text = ""
Me.txtNEP.Text = ""
Me.txtNER.Text = ""
Me.txtPM.Text = ""
Me.txtPM2.Text = ""
Me.txtPureza.Text = ""
Me.lblMoles.Caption = ""
Me.lblResultado.Caption = ""
Me.txtGramos.SetFocus

End Sub

Private Sub cmdSalir_Click()
    usfPureza.Hide
End Sub

Private Sub CommandButton1_Click()
    usfMM.Show
End Sub



Private Sub UserForm_Initialize()
    Me.txtGramos.SetFocus
    

    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
    
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

