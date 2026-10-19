VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculadora"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Calculadora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4440
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   2160
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   230
      TabIndex        =   34
      Top             =   550
      Width           =   255
   End
   Begin VB.TextBox Consola 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   230
      TabIndex        =   32
      Top             =   135
      Width           =   2505
   End
   Begin VB.CommandButton Foco 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Foco"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Number0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1040
   End
   Begin VB.CommandButton Coma 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton Suma 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton Calcular 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2340
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton Number1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton Number2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton Number3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton Resta 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton Number4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2400
      Width           =   500
   End
   Begin VB.CommandButton Number5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2400
      Width           =   500
   End
   Begin VB.CommandButton Number6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Width           =   500
   End
   Begin VB.CommandButton Multiplicación 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   500
   End
   Begin VB.CommandButton Recíproco 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2400
      Width           =   500
   End
   Begin VB.CommandButton Number7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1920
      Width           =   500
   End
   Begin VB.CommandButton Number8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1920
      Width           =   500
   End
   Begin VB.CommandButton Number9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1920
      Width           =   500
   End
   Begin VB.CommandButton División 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1920
      Width           =   500
   End
   Begin VB.CommandButton Porcentaje 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Width           =   500
   End
   Begin VB.CommandButton Atrás 
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Picture         =   "Calculadora.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Clean 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Opuesto 
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Raíz 
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      Picture         =   "Calculadora.frx":108C3
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton MLess 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Resta el número mostrado al número en memoria (CTRL+ Q)"
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton MPlus 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Suma el número mostrado al número en memoria (CTRL+ P)"
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton MS 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Almacena en memoria el número mostrado (CTRL+ M)"
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton MR 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Recupera el número almacenado (CTRL+ R)"
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton MC 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Limpia el número almacenado en memoria (CTRL+ L)"
      Top             =   960
      Width           =   500
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   420
      MaxLength       =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   360
      Width           =   2390
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      MaxLength       =   17
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   180
      MaxLength       =   17
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Width           =   2650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num1 As Double
Dim num2 As Double
Dim numM As Double
Dim numT As Double
Dim numR As Double
Dim sig As Double
Dim buffo As Double
Dim buffoS As String
Dim buffoR As Double
Dim buffoRaiz As String
Dim temp As Double
Dim op As String
Dim opRaiz As Boolean
Dim textval As String
Dim numval As String
Dim tempCon As String
Dim error As Boolean
Dim errorZ As Boolean
Dim isCalculado As Boolean
Dim isSemiCalculado As Boolean
Dim limit As Boolean
Dim isMulti As Boolean
Dim textTemp As Boolean
Dim textT1 As String
Dim textT2 As String
Dim changed As Boolean
Dim c As Integer
 


Private Sub Atrás_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Atrasmethod

Foco.SetFocus
End Sub

Private Sub Calcular_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Calcula(True)

Foco.SetFocus

End Sub

Private Sub Clean_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then

Consola.Text = ""
buffoS = ""
tempCon = ""
isMulti = False
Text1.Text = "0"
isCalculado = False
op = ""

End If

Foco.SetFocus

End Sub

Private Sub Coma_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Punto

Foco.SetFocus
End Sub

Private Sub Command1_GotFocus()
Foco.SetFocus
End Sub

Private Sub Command2_GotFocus()
Foco.SetFocus
End Sub

Private Sub Command3_GotFocus()
Foco.SetFocus
End Sub

Private Sub Command4_GotFocus()
Foco.SetFocus
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then

If (error = True Or Text1.Text = "No se puede dividir entre cero" Or Text1.Text = "Entrada no válida") Then
Consola.Text = ""
buffoS = ""
num1 = 0
num2 = 0
op = ""
End If

Text1.Text = "0"
isCalculado = True
End If

Foco.SetFocus
End Sub


Private Sub Consola_Change()
If (errorZ = True) Then
error = True
Text1.Text = "No se puede dividir entre cero"
Else
sig = 0
For i = 1 To Len(Consola.Text)
If Mid(Consola.Text, i, 1) = ")" Then sig = sig + 1
Next

Consola.Text = Right(Consola.Text, 26 + sig)
End If

End Sub

Sub División_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Operacion("/", True)
End Sub

Private Sub Foco_Click()

If Not (KeyAscii = vbKeySpace) Then Call Calcula(True)

End Sub

Private Sub Foco_KeyPress(KeyAscii As Integer)

If (KeyAscii = vbKeyNumpad1 Or KeyAscii = vbKey1) Then Call Num("1", False)

If (KeyAscii = vbKeyNumpad2 Or KeyAscii = vbKey2) Then Call Num("2", False)

If (KeyAscii = vbKeyNumpad3 Or KeyAscii = vbKey3) Then Call Num("3", False)

If (KeyAscii = vbKeyNumpad4 Or KeyAscii = vbKey4) Then Call Num("4", False)

If (KeyAscii = vbKeyNumpad5 Or KeyAscii = vbKey5) Then Call Num("5", False)

If (KeyAscii = vbKeyNumpad6 Or KeyAscii = vbKey6) Then Call Num("6", False)

If (KeyAscii = vbKeyNumpad7 Or KeyAscii = vbKey7) Then Call Num("7", False)

If (KeyAscii = vbKeyNumpad8 Or KeyAscii = vbKey8) Then Call Num("8", False)

If (KeyAscii = vbKeyNumpad9 Or KeyAscii = vbKey9) Then Call Num("9", False)

If (KeyAscii = vbKeyNumpad0 Or KeyAscii = vbKey0) Then Call Zero

If (KeyAscii = 44 Or KeyAscii = 46) Then Call Punto

If (KeyAscii = vbKeyBack And Not Text1.Text = "") Then Call Atrasmethod

If (KeyAscii = vbMenuAccelCtrlL) Then Call MCmethod

If (KeyAscii = vbMenuAccelCtrlR) Then Call MRmethod

If (KeyAscii = vbMenuAccelCtrlM) Then Call MSmethod

If (KeyAscii = vbMenuAccelCtrlP) Then Call MPlusmethod

If (KeyAscii = vbMenuAccelCtrlQ) Then Call MLessmethod

If (KeyAscii = 47) Then Call Operacion("/", False)

If (KeyAscii = 42) Then Call Operacion("*", False)

If (KeyAscii = 45) Then Call Operacion("-", False)

If (KeyAscii = 43) Then Call Operacion("+", False)

If (KeyAscii = vbMenuAccelCtrlZ And isCalculado = False) Then

If (textTemp) Then
    Text1.Text = textT1
Else:
    Text1.Text = textT2
End If

End If

If (KeyAscii = vbMenuAccelCtrlC) Then
Clipboard.Clear
Clipboard.SetText (Text1.Text)
End If

If (KeyAscii = vbMenuAccelCtrlX) Then
Clipboard.Clear
Clipboard.SetText (Text1.Text)
Text1.Text = ""
isCalculado = False
End If

If (KeyAscii = vbMenuAccelCtrlV) Then

If (IsNumeric(Clipboard.GetText) And Not Clipboard.GetText Like "*.*") Then Text1.Text = Clipboard.GetText

End If

If (KeyAscii = vbMenuAccelCtrlO) Then

If Not (c = 2) Then
c = c + 1
Else: If (c = 2) Then c = 1
End If
If (c = 1) Then Form1.BackColor = &HE0E0E0
If (c = 2) Then Form1.BackColor = &H404040

End If

End Sub

Private Sub Form_Load()

num1 = 0
textT1 = Text1.Text
c = 1

End Sub

Private Sub MC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call MCmethod
Foco.SetFocus
End Sub

Private Sub MLess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call MLessmethod
Foco.SetFocus
End Sub

Private Sub MPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call MPlusmethod
Foco.SetFocus
End Sub

Private Sub MR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call MRmethod
Foco.SetFocus
End Sub

Private Sub MS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call MSmethod
Foco.SetFocus
End Sub

Private Sub Multiplicación_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Operacion("*", True)
End Sub

Private Sub Number0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Zero
Foco.SetFocus
End Sub

Private Sub Number1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("1", True)
End Sub

Private Sub Number2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("2", True)
End Sub

Private Sub Number3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("3", True)
End Sub

Private Sub Number4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("4", True)
End Sub

Private Sub Number5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("5", True)
End Sub

Private Sub Number6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("6", True)
End Sub

Private Sub Number7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("7", True)
End Sub

Private Sub Number8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("8", True)
End Sub

Private Sub Number9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Num("9", True)
End Sub

Private Sub Opuesto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then
If (error = False) Then
If Not (Text1.Text = "" Or Text1.Text = "0") Then

If (Text1.Text) > 0 Then
numT = CDbl(Text1.Text)
Text1.Text = "-"
Text1.Text = Text1.Text & numT

Else
Text1.Text = Replace(Text1.Text, "-", "")

End If
End If
End If
End If
Foco.SetFocus

End Sub

Private Sub Porcentaje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
If (error = False) Then
If (op = "*") Then

If Not (Text1.Text = "" Or Text1.Text = "0") Then

    num2 = CDec(Text1.Text)
    Text1.Text = num1 * num2 / 100
    Consola.Text = Consola.Text & Text1.Text
    If Not (op = "*") Then
    isCalculado = True
    isSemiCalculado = False
    op = ""
    
End If
End If
End If
End If
End If
Foco.SetFocus
End Sub

Private Sub Raíz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then
If (error = False) Then
If Not (Text1.Text = "" Or Text1.Text = "-") Then

    If (Text1.Text > 0) Then
    
    numR = CDec(Text1.Text)
    If (opRaiz = False) Then
    buffoS = Consola.Text
    buffoR = num1
    buffoRaiz = "sqrt(" & numR & ")"
    If (Consola.Text = "") Then
    isCalculado = True
    Else: isCalculado = False
    End If
    Consola.Text = Consola.Text & "sqrt(" & numR & ")"
    Else
    
    Consola.Text = buffoS & "sqrt(" & buffoRaiz & ")"
    buffoRaiz = "sqrt(" & buffoRaiz & ")"
    End If
    numR = Math.Sqr(numR)
    Text1.Text = numR
    isCalculado = True
    opRaiz = True
    End If
    
    If (Text1.Text < 0) Then
    numR = CDec(Text1.Text)
    Consola.Text = "sqrt(" & numR & ")"
    Text1.Text = "Entrada no válida"
    End If
End If
End If
End If
Foco.SetFocus

End Sub

Private Sub Recíproco_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
If (error = False) Then
If Not (Text1.Text = "" Or Text1.Text = "0") Then

    num1 = CDec(Text1.Text)
    num1 = 1 / num1
    Text1.Text = num1
    isCalculado = True
    
End If
End If
End If
Foco.SetFocus

End Sub

Private Sub Resta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Operacion("-", True)
End Sub

Private Sub Suma_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Operacion("+", True)
End Sub

Private Sub Text1_Change()

If (Len(Text1.Text) >= 16) Then
limit = True
Else: limit = False
End If

If (Not Text1.Text = "-" And Not Text1.Text = "" And Not Text1.Text = "" And Not Text1.Text = "Entrada no válida" And Not Text1.Text = "No se puede dividir entre cero") Then
If (Text1.Text < 0 And Len(Text1.Text) = 16) Then limit = False
End If

If (textTemp = False) Then
textT2 = Text1.Text
textTemp = True
Else:
textT1 = Text1.Text
textTemp = False

End If

If (Len(Text1.Text) > 11) Then
Text1.FontSize = 14
Text1.Height = 360
Text1.Top = 480
End If
If (Len(Text1.Text) > 14) Then
Text1.FontSize = 11
Text1.Height = 320
Text1.Top = 530
End If
If (Len(Text1.Text) > 16) Then
Text1.FontSize = 10
Text1.Height = 320
Text1.Top = 530
End If
If (Len(Text1.Text) > 19) Then
Text1.FontSize = 9
Text1.Height = 320
Text1.Top = 530
End If
If (Len(Text1.Text) > 24) Then
Text1.FontSize = 7
Text1.Height = 320
Text1.Top = 530
End If
If (Len(Text1.Text) < 12) Then
Text1.FontSize = 18
Text1.Height = 450
Text1.Top = 390
End If

changed = True

If (Text1.Text = "Entrada no válida" Or Text1.Text = "No se puede dividir entre cero") Then
error = True
errorZ = True
Else:
error = False
errorZ = False
End If

End Sub

Private Sub Calcula(Final As Boolean)
If (error = False And Not Text1.Text = "No se puede dividir entre cero") Then
If (isCalculado = False Or opRaiz = True) Then

If Not (op = "" Or Text1.Text = "") Then

num2 = Text1.Text

If (op = "+") Then
Text1.Text = num1 + num2
End If
If (op = "-") Then
Text1.Text = num1 - num2
End If
If (op = "/" And Not num2 = 0) Then
Text1.Text = num1 / num2
End If
If (op = "/" And num2 = 0) Then
erroZ = True
Consola.Text = Consola.Text & "0"
Text1.Text = "No se puede dividir entre cero"
errorZ = True
error = True
End If
If (op = "*") Then
Text1.Text = num1 * num2
End If


If (Final = True And Not Text1.Text = "No se puede dividir entre cero") Then
num1 = 0
num2 = 0
op = ""
Consola.Text = ""
buffoS = ""

End If

isCalculado = True

End If
End If
End If
End Sub

Private Sub Num(st As String, bo As Boolean)
If (error = False) Then
If (limit = False Or isCalculado = True) Then
    If (Text1.Text = "0" Or isCalculado = True) Then
    If (opRaiz = True) Then
    Consola.Text = ""
    End If
    Text1.Text = st
    isCalculado = False
    opRaiz = False
    
Else
    Text1.Text = Text1.Text & st
End If
End If
opRaiz = False
End If
    If (bo = True) Then Foco.SetFocus

End Sub

Private Sub Operacion(st As String, bo As Boolean)
If (error = False) Then
If Not (Text1.Text = "" Or Text1.Text = " " And op = "") Then

If (op = "/" And Text1.Text = "0" And Not isCalculado = True) Then
errorZ = True
error = True
End If

    
    
    If (Not isCalculado = True And Not buffoS = "" And Not buffoS = " ") Then
    isSemiCalculado = True
    buffoS = Consola.Text & CDbl(Text1.Text) & " "
    
    Consola.Text = Consola.Text & CDbl(Text1.Text) & " " & st & " "
    Call Calcula(False)
    If (error = False) Then num1 = CDbl(Text1.Text)
    op = st
    
    ElseIf (isCalculado = True And Not buffoS = "" And Not buffoS = " ") Then
    isSemiCalculado = True
    op = st
    Consola.Text = buffoS & op & " "
    
    Else
    op = st
    num1 = CDbl(Text1.Text)
    If Not (opRaiz = True) Then Consola.Text = num1 & " " & op & " "
    If (opRaiz = True) Then Consola.Text = Consola.Text & " " & op & " "
    buffoS = num1 & " "
    End If
    isCalculado = True
    opRaiz = False
   
    
ElseIf (Text1.Text = "" Or Text1.Text = " " And Not op = "") Then op = st

End If
End If
    If (bo = True) Then Foco.SetFocus

End Sub

Private Sub Zero()
If (error = False) Then
If (Text1.Text = "" Or isCalculado = True) Then

    Text1.Text = 0
    isCalculado = False

ElseIf (Text1.Text Like "*,*" Or Not Text1.Text = "0") Then

    Text1.Text = Text1.Text & 0

End If
End If
End Sub

Private Sub MCmethod()
If (error = False) Then
numM = 0
Text2.Text = ""
End If
End Sub

Private Sub MLessmethod()
If (error = False) Then
If Not (Text1.Text = "") Then
numM = numM - Text1.Text
Text2.Text = "M"
End If
End If
End Sub

Private Sub MPlusmethod()
If (error = False) Then
If Not (Text1.Text = "") Then
numM = numM + Text1.Text
Text2.Text = "M"
End If
End If
End Sub

Private Sub MRmethod()
If (error = False) Then
Text1.Text = numM
If (Consola.Text = "") Then
isCalculado = True
Else: isCalculado = False
End If
End If
End Sub

Private Sub MSmethod()
If (error = False) Then
If Not (Text1.Text = "" Or Text1.Text = " ") Then
numM = Text1.Text
If (Consola.Text = "") Then
isCalculado = True
Else: isCalculado = False
End If
Text2.Text = "M"
End If
End If
End Sub

Private Sub Atrasmethod()

If (isCalculado = False And Not Text1.Text = "" And error = False) Then

If Not (Mid(Text1.Text, 1, Len(Text1.Text) - 1) = "-" Or Mid(Text1.Text, 1, Len(Text1.Text) - 1) = "-0" Or Mid(Text1.Text, 1, Len(Text1.Text) - 1) = "") Then

Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)

Else: Text1.Text = 0
End If

End If
End Sub

Private Sub Punto()

If (error = False) Then
If Not (Text1.Text Like "*,*" Or Text1.Text = "" Or isCalculado = True) Then
 
Text1.Text = Text1.Text & ","

End If
End If
End Sub

