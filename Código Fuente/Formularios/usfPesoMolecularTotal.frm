VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfPesoMolecularTotal 
   Caption         =   "Calculadora de Peso Molecular"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   OleObjectBlob   =   "usfPesoMolecularTotal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfPesoMolecularTotal"
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
    Dim Resultado1 As String




Private Sub cmdCalcular_Click()
    If Me.txtNumeroAtómico1.Text = "" Then
        MsgBox "Debe ingresar datos"
   
    ElseIf Me.txtPM.Text = "" Then
        MsgBox "Debe ingresar datos"
   
    Else
    
    
    
    Resultado1 = Me.txtNumeroAtómico1.Value * Me.txtPM.Value
    Me.lblResultado1.Caption = Resultado1
    
    End If
End Sub

Private Sub cmdInstrucciones_Click()
    MsgBox "¡Bienvenido/a a la calculadora de Peso Molecular!. Para calcular el peso molecular de un elemento, necesitarás un elemento, su número atómico y su peso molecular. Si no conoces el peso molecular de un elemento, puedes encontrar en el botón Tabla Peso Molecular todos los elementos con sus respectivos pesos moleculares. Para Calcular y teniendo todos los datos anteriormente mencionados, solo se necesita presionar el botón Calcular. Si quieres calcular otro elemento, presiona el botón limpiar y rellena nuevamente con los datos solicitados. El botón Salir te llevará al Menú Principal."
End Sub

Private Sub cmdLimpiar_Click()
    Me.txtNumeroAtómico1.Text = ""
    Me.txtPM.Text = ""
    cboElemento1.ListIndex = "0"
    lblResultado1.Caption = ""
   
End Sub

Private Sub cmdPM_Click()
    usfMM.Show
End Sub



Private Sub cmdSalir_Click()
    End
End Sub

Private Sub cmdVolver_Click()
    usfPesoMolecularTotal.Hide
End Sub

Private Sub UserForm_Initialize()
    
    cboElemento1.AddItem ("Ac")
    cboElemento1.AddItem ("Al")
    cboElemento1.AddItem ("Am")
    cboElemento1.AddItem ("Sb")
    cboElemento1.AddItem ("Ar")
    cboElemento1.AddItem ("As")
    cboElemento1.AddItem ("At")
    cboElemento1.AddItem ("S")
    cboElemento1.AddItem ("Ba")
    cboElemento1.AddItem ("Be")
    cboElemento1.AddItem ("Bk")
    cboElemento1.AddItem ("Bi")
    cboElemento1.AddItem ("Bh")
    cboElemento1.AddItem ("B")
    cboElemento1.AddItem ("Br")
    cboElemento1.AddItem ("Cd")
    cboElemento1.AddItem ("Ca")
    cboElemento1.AddItem ("Cf")
    cboElemento1.AddItem ("C")
    cboElemento1.AddItem ("Ce")
    cboElemento1.AddItem ("Cs")
    cboElemento1.AddItem ("Cl")
    cboElemento1.AddItem ("Co")
    cboElemento1.AddItem ("Cu")
    cboElemento1.AddItem ("Cn")
    cboElemento1.AddItem ("Cr")
    cboElemento1.AddItem ("Cm")
    cboElemento1.AddItem ("Ds")
    cboElemento1.AddItem ("Dy")
    cboElemento1.AddItem ("Db")
    cboElemento1.AddItem ("Es")
    cboElemento1.AddItem ("Er")
    cboElemento1.AddItem ("Sc")
    cboElemento1.AddItem ("Sn")
    cboElemento1.AddItem ("Sr")
    cboElemento1.AddItem ("Eu")
    cboElemento1.AddItem ("Fm")
    cboElemento1.AddItem ("Fl")
    cboElemento1.AddItem ("F")
    cboElemento1.AddItem ("P")
    cboElemento1.AddItem ("Fr")
    cboElemento1.AddItem ("Gd")
    cboElemento1.AddItem ("Ga")
    cboElemento1.AddItem ("Ge")
    cboElemento1.AddItem ("Hf")
    cboElemento1.AddItem ("Hs")
    cboElemento1.AddItem ("He")
    cboElemento1.AddItem ("H")
    cboElemento1.AddItem ("Fe")
    cboElemento1.AddItem ("Ho")
    cboElemento1.AddItem ("In")
    cboElemento1.AddItem ("I")
    cboElemento1.AddItem ("Ir")
    cboElemento1.AddItem ("Yb")
    cboElemento1.AddItem ("Y")
    cboElemento1.AddItem ("Kr")
    cboElemento1.AddItem ("La")
    cboElemento1.AddItem ("Lr")
    cboElemento1.AddItem ("Li")
    cboElemento1.AddItem ("Lv")
    cboElemento1.AddItem ("Lu")
    cboElemento1.AddItem ("Mg")
    cboElemento1.AddItem ("Mn")
    cboElemento1.AddItem ("Mt")
    cboElemento1.AddItem ("Md")
    cboElemento1.AddItem ("Hg")
    cboElemento1.AddItem ("Mo")
    cboElemento1.AddItem ("Mc")
    cboElemento1.AddItem ("Nd")
    cboElemento1.AddItem ("Ne")
    cboElemento1.AddItem ("Np")
    cboElemento1.AddItem ("Nh")
    cboElemento1.AddItem ("Nb")
    cboElemento1.AddItem ("Ni")
    cboElemento1.AddItem ("N")
    cboElemento1.AddItem ("No")
    cboElemento1.AddItem ("Og")
    cboElemento1.AddItem ("Au")
    cboElemento1.AddItem ("Os")
    cboElemento1.AddItem ("O")
    cboElemento1.AddItem ("Pd")
    cboElemento1.AddItem ("Ag")
    cboElemento1.AddItem ("Pt")
    cboElemento1.AddItem ("Pb")
    cboElemento1.AddItem ("Pu")
    cboElemento1.AddItem ("Po")
    cboElemento1.AddItem ("K")
    cboElemento1.AddItem ("Pr")
    cboElemento1.AddItem ("Pm")
    cboElemento1.AddItem ("Pa")
    cboElemento1.AddItem ("Ra")
    cboElemento1.AddItem ("Rn")
    cboElemento1.AddItem ("Re")
    cboElemento1.AddItem ("Rh")
    cboElemento1.AddItem ("Rg")
    cboElemento1.AddItem ("Rb")
    cboElemento1.AddItem ("Ru")
    cboElemento1.AddItem ("Rf")
    cboElemento1.AddItem ("Sm")
    cboElemento1.AddItem ("Sg")
    cboElemento1.AddItem ("Se")
    cboElemento1.AddItem ("Si")
    cboElemento1.AddItem ("Na")
    cboElemento1.AddItem ("Tl")
    cboElemento1.AddItem ("Ta")
    cboElemento1.AddItem ("Tc")
    cboElemento1.AddItem ("Te")
    cboElemento1.AddItem ("Ts")
    cboElemento1.AddItem ("Tb")
    cboElemento1.AddItem ("Ti")
    cboElemento1.AddItem ("Th")
    cboElemento1.AddItem ("Tm")
    cboElemento1.AddItem ("U")
    cboElemento1.AddItem ("V")
    cboElemento1.AddItem ("W")
    cboElemento1.AddItem ("Xe")
    cboElemento1.AddItem ("Zn")
    cboElemento1.AddItem ("Zr")
    
    
    
    cboElemento1.ListIndex = "0"
  
    
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

