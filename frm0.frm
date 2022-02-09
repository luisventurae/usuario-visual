VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm0 
   Caption         =   "Cuenta de usuario"
   ClientHeight    =   9375
   ClientLeft      =   4545
   ClientTop       =   1005
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   ScaleHeight     =   9375
   ScaleWidth      =   4290
   Begin VB.CommandButton cmdingresar 
      Caption         =   "ingresar"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   17
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdvolver 
      Caption         =   "Volver"
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingresar nueva cuenta"
      Height          =   315
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingresar nueva cuenta"
      Height          =   315
      Index           =   0
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "salir"
      Height          =   495
      Index           =   0
      Left            =   9360
      TabIndex        =   2
      Top             =   10800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   3
      Top             =   10800
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
      Index           =   0
      Left            =   7800
      TabIndex        =   4
      Top             =   8160
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
      Index           =   0
      Left            =   7800
      TabIndex        =   5
      Top             =   5880
      Width           =   2535
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
      Index           =   1
      Left            =   1320
      TabIndex        =   15
      Top             =   360
      Width           =   2055
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
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   14
      Top             =   1680
      Width           =   2655
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
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpcuenta 
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   5640
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
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -720
      Picture         =   "frm0.frx":0000
      Top             =   -480
      Width           =   15360
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
      Index           =   0
      Left            =   8280
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
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
      Index           =   0
      Left            =   7680
      TabIndex        =   8
      Top             =   5160
      Width           =   2655
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
      Index           =   0
      Left            =   8160
      TabIndex        =   7
      Top             =   7560
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpcuenta 
      Height          =   495
      Index           =   0
      Left            =   8520
      TabIndex        =   6
      Top             =   9120
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
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mvolver 
         Caption         =   "Volver"
      End
   End
End
Attribute VB_Name = "frm0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdingresar_Click()
Form1.lblnombre.Caption = frm0.Text1.Text

If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Ingrese los datos", vbInformation
End If

If Text1.Text <> "" And Text2.Text <> "" Then
frmcargar.Timerpase.Enabled = True
Call usuario
Text1.Text = ""
Text2.Text = ""
Else
cmdingresar.SetFocus
End If

End Sub

Private Sub cmdnuevo_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub cmdvolver_Click(Index As Integer)
frm0.Hide
frmInicio.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub mvolver_Click()
frm0.Hide
frmInicio.Show
End Sub
