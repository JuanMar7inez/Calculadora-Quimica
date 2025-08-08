VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfRLyRE 
   Caption         =   "Calculadora de Reactivo Limitante y Reactivo en Exceso "
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   OleObjectBlob   =   "usfRLyRE.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfRLyRE"
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
Dim Moles1, Moles2 As String
Dim NumeroEstequiometrico1, NumeroEstequiometrico2, Division1, Division2, SubDivision1, SubDivision2 As Double

If Me.txtMasa1.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtMasa2.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtNumeroEstequiometrico.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtNumeroEstequiometrico2.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtReactivo1.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtReactivo2.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtPM1.Text = "" Then
    MsgBox "Debe ingresar un valor"
ElseIf Me.txtPM2.Text = "" Then
    MsgBox "Debe ingresar un valor"
Else
'Calcular Moles
Moles1 = Me.txtMasa1.Value / Me.txtPM1.Value
Me.lblMoles1.Caption = Format(Moles1, "##,##0.00")

Moles2 = Me.txtMasa2.Value / Me.txtPM2.Value
Me.lblMoles2.Caption = Format(Moles2, "##,##0.00")

NumeroEstequiometrico1 = Me.txtNumeroEstequiometrico.Value
NumeroEstequiometrico2 = Me.txtNumeroEstequiometrico2.Value
Division1 = NumeroEstequiometrico1 / NumeroEstequiometrico2
Division2 = NumeroEstequiometrico2 / NumeroEstequiometrico1
SubDivision1 = Me.lblMoles1.Caption / Me.lblMoles2.Caption
SubDivision2 = Me.lblMoles2.Caption / Me.lblMoles1.Caption




If Me.txtNumeroEstequiometrico.Value >= Me.txtNumeroEstequiometrico2.Value Then
        If Division1 > SubDivision1 Then
        Me.lblRLletra.Caption = Me.txtReactivo1.Text
        Me.lblREletra.Caption = Me.txtReactivo2.Text
        Else
        Me.lblREletra.Caption = Me.txtReactivo1.Text
        Me.lblRLletra.Caption = Me.txtReactivo2.Text
          End If
Else
      If Division2 > SubDivision2 Then
        Me.lblRLletra.Caption = Me.txtReactivo2.Text
        Me.lblREletra.Caption = Me.txtReactivo1.Text
        Else
        Me.lblREletra.Caption = Me.txtReactivo2.Text
        Me.lblRLletra.Caption = Me.txtReactivo1.Text
            End If
    
End If

    
End If

    
    



End Sub

Private Sub cmdInstrucciones_Click()
    MsgBox "Bienvenido/a a la calculadora de Reactivo Limitante y Reactivo en Exceso!. Antes de Empezar, necesitarás los siguientes valores para empezar a calcular: Coeficiente Estequiométrico, Reactivo, Masa en gramos y Peso Molecular; todo estos datos del Reactivo 1 y 2. Una vez completados los datos, se procede a calcular presionando el botón Calcular, dando en el mismo instante, el Compuesto con el Reactivo Limitante y Exceso, además de los moles del Reactivo 1 y 2. Para rellenar con otros datos, debes limpiar el formulario con el botón Limpiar. Si no sabes el Peso Molecular de algún elemento, puedes presionar el botón Calculadora Peso Molecular y hacer los cálculos. Para ir al Menú Principal, presiona el botón Salir. "
End Sub

Private Sub cmdLimpiar_Click()
    Me.txtNumeroEstequiometrico.Text = ""
    Me.txtNumeroEstequiometrico2.Text = ""
    Me.txtReactivo1.Text = ""
    Me.txtReactivo2.Text = ""
    Me.txtMasa1.Text = ""
    Me.txtMasa2.Text = ""
    Me.txtPM1.Text = ""
    Me.txtPM2.Text = ""
    Me.lblMoles1.Caption = ""
    Me.lblMoles2.Caption = ""
    Me.lblREletra.Caption = ""
    Me.lblRLletra.Caption = ""
    
    
    
End Sub

Private Sub cmdPM_Click()
    usfPesoMolecularTotal.Show
End Sub

Private Sub cmdSalir_Click()
    usfRLyRE.Hide
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
