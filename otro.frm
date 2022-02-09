VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Cuenta Básica"
   ClientHeight    =   8955
   ClientLeft      =   4290
   ClientTop       =   1005
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   4575
   Begin VB.ComboBox cmbestado 
      Height          =   315
      ItemData        =   "otro.frx":0000
      Left            =   2880
      List            =   "otro.frx":0010
      TabIndex        =   8
      Text            =   "Conectado"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   240
      Picture         =   "otro.frx":003F
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdfoto 
      Caption         =   "cambiar foto"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   7920
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdmusica 
      Caption         =   "escoger musica o video para reproducir"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton cmdingresar 
      Caption         =   "ingresar"
      Default         =   -1  'True
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtnombre 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpestado 
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   855
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
      _cx             =   1508
      _cy             =   450
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   2415
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpmusica 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4335
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
      _cx             =   7646
      _cy             =   7858
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      X1              =   120
      X2              =   4320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su nombre"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -1320
      Picture         =   "otro.frx":2504
      Top             =   -360
      Width           =   15360
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mcerrar 
         Caption         =   "Cerrar sesion"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbestado_Click()
Select Case cmbestado.ListIndex
Case 0
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\conectado.wma"
wmpestado.URL = sonido
Shape1.BorderColor = &HFF00&
Shape1.FillColor = &HFF00&
Case 1
Shape1.BorderColor = RGB(255, 0, 0)
Shape1.FillColor = RGB(255, 0, 0)
Case 2
Shape1.BorderColor = RGB(231, 120, 19)
Shape1.FillColor = RGB(231, 120, 19)
Case 3
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\desconectado.wma"
wmpestado.URL = sonido
Shape1.BorderColor = RGB(190, 233, 254)
Shape1.FillColor = RGB(190, 233, 254)
End Select
End Sub

Private Sub cmdfoto_Click()
 With CommonDialog1
       .DialogTitle = "Abrir un archivo de imagen"
       .Filter = "Archivos de imagenes *.jpg;*jpeg;*png;*bmp"
       .ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
 End With
End Sub

Private Sub cmdingresar_Click()
lblnombre.Caption = txtnombre.Text
End Sub

Private Sub cmdmusica_Click()
 With CommonDialog1
       .DialogTitle = "Abrir un archivo de video"
       .Filter = "Archivos de imagenes *.wmv;*mp4;*mp3;*wma;*jpg"
       .ShowOpen
wmpmusica.URL = CommonDialog1.FileName
 End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mcerrar_Click()
Form1.Hide
frm0.Show
End Sub
