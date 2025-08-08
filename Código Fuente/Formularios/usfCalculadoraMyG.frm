VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCalculadoraMyG 
   Caption         =   "Conversor de Gramos y Moles"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "usfCalculadoraMyG.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfCalculadoraMyG"
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
Dim Gramos, Moles As Double


    
    If Not Me.txtMoles1.Value = "" Then
    
    
    Gramos = Me.txtMoles1.Value * Me.txtPesoAtómico1.Value
    Me.lblGramos1.Caption = Gramos
    ElseIf Me.txtGramos1.Text = "" Then
        MsgBox "Debe ingresar un dato"
    ElseIf Me.txtPesoAtómico2.Text = "" Then
         MsgBox "Debe ingresar un dato"
    Else
    Moles = Me.txtGramos1.Value / Me.txtPesoAtómico2.Value
    
    Me.lblMoles1.Caption = Moles
    End If
    
End Sub

Private Sub cmdInstrucciones_Click()
    MsgBox "¡Bienvenido/a al conversor de gramos y moles!. Sólo puedes convertir un valor a la vez (Mol a Gramos, o Gramos a Mol). Usa el botón calcular para conocer el valor deseado, el botón limpiar para limpiar los formularios y el botón salir para volver al Menú Principal."
End Sub

Private Sub cmdLimpiar_Click()
    Me.txtGramos1.Text = ""
    Me.txtMoles1.Text = ""
    Me.txtPesoAtómico1.Text = ""
    Me.txtPesoAtómico2.Text = ""
    Me.lblGramos1.Caption = ""
    Me.lblMoles1.Caption = ""
    Me.txtMoles1.SetFocus
End Sub

Private Sub cmdSalir_Click()
    usfCalculadoraMyG.Hide
End Sub


Private Sub UserForm_Initialize()
    Me.txtMoles1.SetFocus
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

