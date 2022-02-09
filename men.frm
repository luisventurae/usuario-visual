VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm2 
   Caption         =   "Menú"
   ClientHeight    =   3840
   ClientLeft      =   4230
   ClientTop       =   3495
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5625
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   3360
      Picture         =   "men.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   360
      Width           =   1695
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.ComboBox cmbopcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "men.frx":6234
      Left            =   960
      List            =   "men.frx":624D
      TabIndex        =   4
      Text            =   "Escoge uno de las aplicaciones"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdcerrarsesion 
      Caption         =   "cerrar sesion"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1815
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
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
      _cx             =   2143
      _cy             =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escoger una aplicacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "men.frx":62D0
      Top             =   0
      Width           =   15360
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mcerrar 
         Caption         =   "Cerrar sesion..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mimagen 
         Caption         =   "Cambiar imagen"
      End
      Begin VB.Menu mestado 
         Caption         =   "Estado"
         Begin VB.Menu mdisponible 
            Caption         =   "Disponible"
         End
         Begin VB.Menu mocupado 
            Caption         =   "Ocupado"
         End
         Begin VB.Menu mausente 
            Caption         =   "Ausente"
         End
         Begin VB.Menu mdesconectado 
            Caption         =   "Desconectado"
         End
      End
   End
   Begin VB.Menu mayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mayudaluis 
         Caption         =   "Ayuda a Luis..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdaceptar_Click()
'Texto
If cmbopcion.Text = "Ir a texto" Then
frm2.Hide
frm3.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a texto.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Calculadora
If cmbopcion.Text = "Ir a calculadora" Then
frm2.Hide
frm4.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a calculadora.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Reproductor
If cmbopcion.Text = "Ir a reproductor" Then
frm2.Hide
frm5.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a reproductor.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Traductor
If cmbopcion.Text = "Ir a traductor" Then
frm2.Hide
frm6.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a traductor.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Registro
If cmbopcion.Text = "Ir a registro" Then
frm2.Hide
frm7.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a registrador.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Internet
If cmbopcion.Text = "Ir a navegador de internet" Then
frm2.Hide
frm8.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ir a navegador.wma"
wmp1.URL = sonido
wmp1.Visible = False
End If
'Otros
If cmbopcion.Text = "Otras Aplicaciones" Then
frm2.Hide
frm9.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\otros.wma"
wmp1.URL = sonido
End If
End Sub


Private Sub cmdcerrarsesion_Click()
frm2.Hide
frm1.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\cerrando.wma"
wmp1.URL = sonido
wmp1.Visible = False
End Sub

Private Sub mausente_Click()
Shape1.BorderColor = RGB(231, 120, 19)
Shape1.FillColor = RGB(231, 120, 19)
End Sub

Private Sub mayudaluis_Click()
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\tutorial visual.wma"
wmp1.URL = sonido
wmp1.Visible = False
MsgBox ("mira el tutorial que esta hecho para explicar paso a paso lo que debes hacer para cada situacion con este programa creado por Luis"), vbInformation, "Ayuda a Luis"
frmtutor.Show
End Sub

Private Sub mcerrar_Click()
frm2.Hide
frm1.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\cerrando.wma"
wmp1.URL = sonido
wmp1.Visible = False
End Sub

Private Sub mdesconectado_Click()
Shape1.BorderColor = RGB(190, 233, 254)
Shape1.FillColor = RGB(190, 233, 254)
End Sub

Private Sub mdisponible_Click()
Shape1.BorderColor = RGB(0, 255, 0)
Shape1.FillColor = RGB(0, 255, 0)
End Sub

Private Sub mimagen_Click()
 With CommonDialog1
       .DialogTitle = "Abrir un archivo de imagen"
       .Filter = "Archivos de imagenes *.jpg;*jpeg;*png;*bmp"
       .ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
 End With
End Sub

Private Sub mocupado_Click()
Shape1.BorderColor = RGB(255, 0, 0)
Shape1.FillColor = RGB(255, 0, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
