VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm8 
   Caption         =   "Internet Venz"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   1575
   ClientWidth     =   16980
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   16980
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   11055
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   16695
      ExtentX         =   29448
      ExtentY         =   19500
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdlista 
      Caption         =   "lista"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox txtweb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "inter.frx":0000
      Left            =   3000
      List            =   "inter.frx":0019
      TabIndex        =   6
      Text            =   "www.google.com.pe"
      Top             =   240
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   12480
      Top             =   120
   End
   Begin VB.CommandButton cmdactualizar 
      Caption         =   "Actualizar  página"
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdir 
      Caption         =   "Ir"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdadelante 
      BackColor       =   &H8000000D&
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   1320
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdatras 
      BackColor       =   &H8000000D&
      Caption         =   "Atras"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
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
      Height          =   495
      Left            =   13440
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "inter.frx":009A
      Top             =   0
      Width           =   240
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   600
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
      _cy             =   661
   End
   Begin VB.Image Image2 
      Height          =   12960
      Left            =   -120
      Picture         =   "inter.frx":03DC
      Top             =   0
      Width           =   17280
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mvolver 
         Caption         =   "Volver a menú"
      End
   End
End
Attribute VB_Name = "frm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactualizar_Click()
WebBrowser1.Refresh
End Sub

Private Sub cmdadelante_Click()
WebBrowser1.GoForward
End Sub

Private Sub cmdatras_Click()
On Error GoTo errSub
WebBrowser1.GoBack
errSub:
If Err.Number = -2147467259 Then
WebBrowser1.GoForward
End If
End Sub

Private Sub cmdir_Click()
WebBrowser1.Navigate2 txtweb.Text
End Sub

Private Sub cmdlista_Click()
cmbweb.AddItem txtweb.Text
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate2 ("www.google.com.pe")
End Sub

Private Sub mvolver_Click()
mensaje = "Esta seguro que desea salir de Internet Venz?"
titulo = "¿Seguro?"
estilo = vbYesNo + vbQuestion
rpta = MsgBox(mensaje, estilo, titulo)
If rpta = vbYes Then
frm8.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
Else
frm8.Show
frm2.Hide
End If
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
