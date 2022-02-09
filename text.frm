VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm3 
   Caption         =   "Cuadro de texto"
   ClientHeight    =   6120
   ClientLeft      =   3135
   ClientTop       =   2925
   ClientWidth     =   8100
   LinkTopic       =   "Form3"
   ScaleHeight     =   6120
   ScaleWidth      =   8100
   Begin VB.CheckBox chksub 
      Caption         =   "Check1"
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   5520
      Width           =   135
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   5400
      Width           =   135
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   4920
      Width           =   135
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   4440
      Width           =   135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   4080
      Value           =   -1  'True
      Width           =   135
   End
   Begin VB.CommandButton cmdpequeño 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "text.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdgrande 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Picture         =   "text.frx":DCE3
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "limpiar"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.Timer tmrhorafecha 
      Interval        =   1
      Left            =   5640
      Top             =   3720
   End
   Begin VB.CommandButton cmdenviar 
      Caption         =   "Enviar texto"
      Default         =   -1  'True
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "Volver al Menú"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CheckBox chk3 
      Caption         =   "Check3"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   135
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Check2"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   135
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   4080
      Width           =   135
   End
   Begin VB.TextBox txthoja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Subrayado"
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "color verde"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "color rojo"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "normal"
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2355
      _cy             =   873
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Caption         =   "hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6240
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "color azul"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "MAYUSCULA"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   5040
      Width           =   4095
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cursiva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Negrita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe un texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -1800
      Picture         =   "text.frx":1B9C6
      Top             =   -2760
      Width           =   15360
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chk1_Click()
If chk1.Value = 1 Then
txthoja.Font.Bold = True
Else
txthoja.Font.Bold = False
End If
End Sub

Private Sub chk2_Click()
If chk2.Value = 1 Then
txthoja.Font.Italic = True
Else
txthoja.Font.Italic = False
End If
End Sub

Private Sub chk3_Click()
If chk3.Value = 1 Then
txthoja = UCase(txthoja.Text)
Else
txthoja = LCase(txthoja.Text)
End If
End Sub

Private Sub chk4_Click()
If chk4.Value = 1 Then
txthoja.ForeColor = RGB(0, 0, 255)
Else
txthoja.ForeColor = RGB(0, 0, 0)
End If
End Sub

Private Sub chksub_Click()
If chksub.Value = 1 Then
txthoja.Font.Underline = True
Else
txthoja.Font.Underline = False
End If
End Sub

Private Sub cmdenviar_Click()
mensaje = "¿Seguro que desea enviar el mensaje?"
titulo = "Texto"
estilo = vbYesNo + vbQuestion
rpta = MsgBox(mensaje, estilo, titulo)
If rpta = vbYes Then
MsgBox "mensaje enviado a Luis"
txthoja.Text = ""
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\mensaje enviado.wma"
wmp1.URL = sonido
wmp1.Visible = False
ElseIf rpta = vbNo Then
MsgBox "No se envió el mensaje"
End If
End Sub

Private Sub cmdgrande_Click()
If cmdgrande.Value = 1 Then
txthoja.Font.Size = 20
Else
txthoja.Font.Size = 25
End If
End Sub

Private Sub cmdlimpiar_Click()
txthoja.Text = ""
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\clic.wma"
wmp1.URL = sonido
wmp1.Visible = False
End Sub

Private Sub cmdmenu_Click()
frm3.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
wmp1.Visible = False
End Sub

Private Sub cmdpequeño_Click()
If cmdpequeño.Value = 1 Then
txthoja.Font.Size = 10
Else
txthoja.Font.Size = 10
End If
End Sub

Private Sub Label3_Click()
frm3.Hide
frmcalender.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = 1 Then
txthoja.ForeColor = RGB(0, 0, 0)
Else
txthoja.ForeColor = RGB(0, 0, 0)
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = 1 Then
txthoja.ForeColor = RGB(0, 0, 255)
Else
txthoja.ForeColor = RGB(0, 0, 255)
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = 1 Then
txthoja.ForeColor = RGB(255, 0, 0)
Else
txthoja.ForeColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = 1 Then
txthoja.ForeColor = RGB(0, 255, 0)
Else
txthoja.ForeColor = RGB(0, 255, 0)
End If
End Sub

Private Sub tmrhorafecha_Timer()
lblhora.Caption = Time
lblfecha.Caption = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
