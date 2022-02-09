VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm7 
   BackColor       =   &H8000000B&
   Caption         =   "Registradora de Datos"
   ClientHeight    =   6900
   ClientLeft      =   3240
   ClientTop       =   2025
   ClientWidth     =   9330
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9330
   Begin VB.Frame frasiono 
      Caption         =   "Si ó No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2400
      TabIndex        =   38
      Top             =   2280
      Width           =   2535
      Begin VB.OptionButton optno 
         Caption         =   "Option6"
         Height          =   195
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Value           =   -1  'True
         Width           =   135
      End
      Begin VB.OptionButton optsi 
         Caption         =   "Option5"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label21 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   6240
   End
   Begin VB.OptionButton opttrabyest 
      Caption         =   "Option4"
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   35
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optnoche 
      Caption         =   "Option3"
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   34
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton opttarde 
      Caption         =   "Option2"
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   33
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optmañana 
      Caption         =   "Option1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   32
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkcasado 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2400
      TabIndex        =   30
      Top             =   2040
      Width           =   135
   End
   Begin VB.ComboBox cboedad 
      Height          =   315
      ItemData        =   "registro.frx":0000
      Left            =   1800
      List            =   "registro.frx":003A
      TabIndex        =   29
      Text            =   "- Seleccione su edad -"
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtapellido 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Top             =   840
      Width           =   4695
   End
   Begin VB.ListBox lstapellido 
      Height          =   2400
      Left            =   2040
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdvolver 
      Caption         =   "Volver a menú"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdquitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin VB.ListBox lstcasado 
      Height          =   2400
      Left            =   7800
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox lstturno 
      Height          =   2400
      Left            =   6240
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ListBox lsttrabajo 
      Height          =   2400
      Left            =   5400
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.ListBox lstedad 
      Height          =   2400
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox lstnombre 
      Height          =   2400
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   4920
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
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
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
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
      Left            =   240
      TabIndex        =   36
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Casado"
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
      Left            =   2640
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabaja y estudia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7440
      TabIndex        =   26
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Noche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7440
      TabIndex        =   25
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Mañana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabaja:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado civil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   1680
      Width           =   855
   End
   Begin VB.Line Line20 
      BorderColor     =   &H8000000B&
      X1              =   3360
      X2              =   3960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line19 
      BorderColor     =   &H8000000B&
      X1              =   7200
      X2              =   7680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line18 
      BorderColor     =   &H8000000B&
      X1              =   6120
      X2              =   6360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line17 
      BorderColor     =   &H8000000B&
      X1              =   4680
      X2              =   5280
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line16 
      BorderColor     =   &H8000000B&
      X1              =   1440
      X2              =   2160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Casado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabaja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Line Line15 
      BorderColor     =   &H8000000B&
      X1              =   9240
      X2              =   8880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000B&
      X1              =   120
      X2              =   480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000B&
      X1              =   9240
      X2              =   9240
      Y1              =   6720
      Y2              =   3120
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000B&
      X1              =   9240
      X2              =   120
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000B&
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   6720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000B&
      X1              =   8880
      X2              =   9120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000B&
      X1              =   9120
      X2              =   9120
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000B&
      X1              =   6840
      X2              =   9120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Turno de trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000B&
      X1              =   7080
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000B&
      X1              =   6840
      X2              =   6840
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000B&
      X1              =   6720
      X2              =   6720
      Y1              =   2880
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000B&
      X1              =   1320
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos"
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
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000B&
      X1              =   360
      X2              =   600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000B&
      X1              =   360
      X2              =   6720
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   360
      X2              =   360
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -2640
      Picture         =   "registro.frx":0098
      Top             =   -2280
      Width           =   15360
   End
End
Attribute VB_Name = "frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdagregar_Click()
'agregar nombre
lstnombre.AddItem (txtnombre.Text)

'agregar apellido
lstapellido.AddItem (txtapellido.Text)

'agregar edad
lstedad.AddItem (cboedad.Text)

'agregar turno
If optmañana.Value = True Then
lstturno.AddItem "Mañana"
End If
If opttarde.Value = True Then
lstturno.AddItem "Tarde"
End If
If optnoche.Value = True Then
lstturno.AddItem "Noche"
End If
If opttrabyest.Value = True Then
lstturno.AddItem "Trabaja y estudia"
End If
If optmañana.Value = False And opttarde.Value = False And optnoche.Value = False And opttrabyest.Value = False Then
lstturno.AddItem " ------------------ "
End If

'agregar estado civil
If chkcasado.Value = Checked Then
lstcasado.AddItem "Si"
Else
lstcasado.AddItem "No"
End If

'agregar trabajo
If optsi.Value = True Then
lsttrabajo.AddItem "Si"
End If
If optno.Value = True Then
lsttrabajo.AddItem "No"
End If

'Dejar campo en blanco
txtnombre.Text = ""
txtapellido.Text = ""
cboedad.Text = "- Seleccione su edad -"
chkcasado.Value = False
optno.Value = True
optmañana.Value = False
opttarde.Value = False
optnoche.Value = False
opttrabyest.Value = False
End Sub

Private Sub cmdlimpiar_Click()
lstnombre.Clear
lstapellido.Clear
lstedad.Clear
lsttrabajo.Clear
lstturno.Clear
lstcasado.Clear

optmañana.Value = False
opttarde.Value = False
optnoche.Value = False
opttrabyest.Value = False
End Sub

Private Sub cmdquitar_Click()
If lstnombre.ListIndex <> -1 Then
lstnombre.RemoveItem lstnombre.ListIndex
End If
If lstapellido.ListIndex <> -1 Then
lstapellido.RemoveItem lstapellido.ListIndex
End If
If lstedad.ListIndex <> -1 Then
lstedad.RemoveItem lstedad.ListIndex
End If
If lsttrabajo.ListIndex <> -1 Then
lsttrabajo.RemoveItem lsttrabajo.ListIndex
End If
If lstturno.ListIndex <> -1 Then
lstturno.RemoveItem lstturno.ListIndex
End If
If lstcasado.ListIndex <> -1 Then
lstcasado.RemoveItem lstcasado.ListIndex
End If
End Sub

Private Sub cmdvolver_Click()
frm7.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
End Sub

Private Sub optno_Click()
If optno.Value = True Then
optmañana.Enabled = False
opttarde.Enabled = False
optnoche.Enabled = False
opttrabyest.Enabled = False
End If
If optno.Value = True Then
optmañana.Value = False
opttarde.Value = False
optnoche.Value = False
opttrabyest.Value = False
End If
End Sub

Private Sub optsi_Click()
If optsi.Value = True Then
optmañana.Enabled = True
opttarde.Enabled = True
optnoche.Enabled = True
opttrabyest.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
lblhora.Caption = Time
End Sub

Private Sub txtapellido_Change()
If txtapellido.Text = "" And txtnombre.Text = "" Then
cmdagregar.Enabled = False
ElseIf txtapellido.Text <> "" And txtnombre.Text <> "" Then
cmdagregar.Enabled = True
End If
If txtapellido.Text = "" Then
cmdagregar.Enabled = False
End If
End Sub

Private Sub txtnombre_Change()
If txtapellido.Text = "" And txtnombre.Text = "" Then
cmdagregar.Enabled = False
ElseIf txtapellido.Text <> "" And txtnombre.Text <> "" Then
cmdagregar.Enabled = True
End If
If txtnombre.Text = "" Then
cmdagregar.Enabled = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
