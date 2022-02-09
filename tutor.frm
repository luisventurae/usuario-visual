VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmtutor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tutorial"
   ClientHeight    =   6150
   ClientLeft      =   3690
   ClientTop       =   1395
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin WMPLibCtl.WindowsMediaPlayer wmpt 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
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
      _cx             =   14208
      _cy             =   10821
   End
   Begin VB.Menu mejecutar 
      Caption         =   "Ejecutar"
      Begin VB.Menu mreproducir 
         Caption         =   "Reproducir"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmtutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mreproducir_Click()
video = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\tutorial.wmv"
wmpt.URL = video
wmpt.Visible = True
End Sub

