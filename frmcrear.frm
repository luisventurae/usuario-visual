VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmcrear 
   Caption         =   "Crear cuenta"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "cancelar"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdcrear 
      Caption         =   "crear ahora"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
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
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblestado 
      BackStyle       =   0  'Transparent
      Caption         =   "/ Coinciden contraseñas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   2655
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpcuenta 
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
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
      _cx             =   2778
      _cy             =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -2760
      Picture         =   "frmcrear.frx":0000
      Top             =   -600
      Width           =   15360
   End
End
Attribute VB_Name = "frmcrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcancelar_Click()
frmcrear.Hide
frmInicio.Show
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub cmdcrear_Click()
If Text1.Text <> "" And Text2.Text <> "" Then
Call Datos
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\creado.wma"
wmpcuenta.URL = sonido
Else
Text1.SetFocus
End If
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
cmdcrear.Enabled = True
Else
cmdcrear.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
cmdcrear.Enabled = True
Else
cmdcrear.Enabled = False
End If
If Text2.Text = Text3.Text Then
lblestado.Caption = "X No coinciden contraseñas"
lblestado.ForeColor = RGB(255, 0, 0)
cmdcrear.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
cmdcrear.Enabled = True
Else
cmdcrear.Enabled = False
End If
If Text3.Text = Text2.Text Then
lblestado.Visible = True
lblestado.Caption = "/ Coinciden contraseñas"
lblestado.ForeColor = &HFF00&
cmdcrear.Enabled = True
Else
lblestado.Visible = True
lblestado.Caption = "X No coinciden contraseñas"
lblestado.ForeColor = RGB(255, 0, 0)
cmdcrear.Enabled = True
End If
If Text3.Text = "" Then
cmdcrear.Enabled = False
lblestado.Visible = False
ElseIf Text3.Text <> Text2.Text Then
cmdcrear.Enabled = False
lblestado.Visible = True
lblestado.Caption = "X No coinciden contraseñas"
lblestado.ForeColor = RGB(255, 0, 0)
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
