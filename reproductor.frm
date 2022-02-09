VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm5 
   Caption         =   "Reproductor de vaudio y video"
   ClientHeight    =   6165
   ClientLeft      =   2205
   ClientTop       =   1800
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8160
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
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
      _cx             =   1931
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpvideo 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
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
      _cx             =   14420
      _cy             =   11033
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mabrir 
         Caption         =   "Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mmenu 
         Caption         =   "Volver a menú"
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu medicion 
      Caption         =   "Edición"
      Begin VB.Menu mcompleto 
         Caption         =   "Pantalla Completa"
      End
      Begin VB.Menu mcerrar 
         Caption         =   "Cerrar video..."
      End
   End
End
Attribute VB_Name = "frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub centrar_repro()
With wmpvideo
.Left = (Me.ScaleWidth - .Width) / 2
.Top = (Me.ScaleHeight - .Height) / 2
End With
End Sub

Private Sub Form_Resize()
centrar_repro
End Sub

Private Sub mabrir_Click()
With CommonDialog1
       .DialogTitle = "Abrir un archivo de video"
       .Filter = "Archivos de imagenes *.wmv;*wma;*mp4;*mp3;*jpeg;*jpg"
       .ShowOpen
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\abrir.wma"
wmp1.URL = sonido
wmpvideo.URL = CommonDialog1.FileName
End With
End Sub

Private Sub mcerrar_Click()
MsgBox ("puedes seleccionar otro video y reproducirlo"), vbInformation, "LUIS"
wmpvideo.Close
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\clic.wma"
wmp1.URL = sonido
End Sub

Private Sub mcompleto_Click()
wmpvideo.fullScreen = True
End Sub

Private Sub mmenu_Click()
frm5.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
