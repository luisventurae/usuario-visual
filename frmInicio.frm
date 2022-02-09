VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmInicio 
   Caption         =   "Cuenta de usuario"
   ClientHeight    =   4965
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7275
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcuenta 
      BackColor       =   &H80000003&
      Caption         =   "Cuenta especial"
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdespecial 
      BackColor       =   &H00808000&
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar Cuenta"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdingresar 
      Caption         =   "Ingresar"
      Default         =   -1  'True
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdcrear 
      Caption         =   "Crear Cuenta"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpcuenta 
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
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
      _cy             =   873
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "cuenta especial"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -3360
      Picture         =   "frmInicio.frx":0000
      Top             =   -1920
      Width           =   15360
   End
   Begin VB.Menu mayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mvideo 
         Caption         =   "Video tutorial"
      End
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcrear_Click()
frmInicio.Hide
frmcrear.Show
End Sub

Private Sub cmdcuenta_Click()
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\especial.wma"
wmpcuenta.URL = sonido
frmInicio.Hide
frm1.Show
cmdespecial.Visible = False
End Sub

Private Sub cmdeliminar_Click()
Call eliminar
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\eliminado.wma"
wmpcuenta.URL = sonido
End Sub

Private Sub cmdespecial_Click()
cmdcuenta.Visible = True
End Sub

Private Sub cmdingresar_Click()
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\ingresarcomun.wma"
wmpcuenta.URL = sonido
frmInicio.Hide
frm0.Show
End Sub

Private Sub cmdsalir_Click()
End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mvideo_Click()
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\tutorial visual.wma"
wmpcuenta.URL = sonido
wmpcuenta.Visible = False
MsgBox ("mira el tutorial que esta hecho para explicar paso a paso lo que debes hacer para cada situacion con este programa creado por Luis"), vbInformation, "Ayuda a Luis"
frmtutor.Show
End Sub
