VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfMenúPrincipal 
   Caption         =   "Menú Principal"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "usfMenúPrincipal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfMenúPrincipal"
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
Private Sub cmdConversor_Click()
    usfCalculadoraMyG.Show
End Sub

Private Sub cmdCPM_Click()
    usfPesoMolecularTotal.Show
End Sub

Private Sub cmdPDR_Click()
    usfRemdimientoPorcentual.Show
End Sub

Private Sub cmdPureza_Click()
    usfPureza.Show
End Sub

Private Sub cmdRLyRE_Click()
    usfRLyRE.Show
End Sub

Private Sub cmdSalir_Click()
    MsgBox "¡Gracias por su Preferencia!"""
    End
End Sub

Private Sub cmdTMM_Click()
    usfMM.Show
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

