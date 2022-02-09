VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm1 
   Caption         =   "cuenta de usuario"
   ClientHeight    =   8970
   ClientLeft      =   5850
   ClientTop       =   1005
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   4080
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingresar nueva cuenta"
      Height          =   315
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdvolver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdingresar 
      Caption         =   "Ingresar"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpcuenta 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
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
      _cx             =   1720
      _cy             =   873
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -5160
      Picture         =   "cuen.frx":0000
      Top             =   -2640
      Width           =   15360
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mcerrar 
         Caption         =   "Cerrar"
      End
      Begin VB.Menu mvolver 
         Caption         =   "Volver"
      End
   End
   Begin VB.Menu mayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mtutorial 
         Caption         =   "Videotutorial"
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdingresar_Click()
'Falta datos
If txt1.Text = "" And txt2.Text = "" Then
MsgBox "usted no ha escrito nada por favor ingrese el nombre de usuario y contraseña", vbInformation, "Falta rellenar"
End If
'Contraseña para validar
If txt1.Text = "Luis" And txt2.Text = "jceluis" Then
frmcargar.Timerpase.Enabled = True
frm1.Hide
frmcargar.Show
frmcargar.Shape1.FillColor = &HC0FFC0
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ingresando.wma"
wmpcuenta.URL = sonido
wmpcuenta.Visible = False
txt1.Text = ""
txt2.Text = ""
Else
MsgBox "contraseña o nombre invalido", vbInformation, "Corregir"
End If
End Sub

Private Sub cmdnuevo_Click()
txt1.Text = ""
txt2.Text = ""
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\usuario visual\audio\clic.wma"
wmpcuenta.URL = sonido
wmpcuenta.Visible = False
End Sub

Private Sub cmdvolver_Click()
frmInicio.cmdespecial.Visible = True
frmInicio.cmdcuenta.Visible = False
frm1.Hide
frmInicio.Show
End Sub

Private Sub mcerrar_Click()
End
End Sub

Private Sub mtutorial_Click()
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\tutorial visual.wma"
wmpcuenta.URL = sonido
wmpcuenta.Visible = False
MsgBox ("mira el tutorial que esta hecho para explicar paso a paso lo que debes hacer para cada situacion con este programa creado por Luis"), vbInformation, "Ayuda a Luis"
frmtutor.Show
End Sub

Private Sub mvolver_Click()
frmInicio.Show
frm1.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
