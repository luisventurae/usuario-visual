VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm9 
   Caption         =   "Otras Aplicaciones"
   ClientHeight    =   10260
   ClientLeft      =   3240
   ClientTop       =   450
   ClientWidth     =   9705
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   Picture         =   "frmotros.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   9705
   Begin VB.Frame fraopt 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame1"
      Height          =   975
      Left            =   7800
      TabIndex        =   40
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton optrespuesta 
         Caption         =   "Respuesta"
         Height          =   375
         Left            =   0
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton optnuevo 
         Caption         =   "Nuevo"
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   39
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   38
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   37
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton opt4 
      Caption         =   "Option1"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   8160
      Width           =   255
   End
   Begin VB.CommandButton cmdvolver 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Volver a menú"
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdultimo 
      Caption         =   ">>l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdprimero 
      Caption         =   "l<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "Borrar Registro"
      Height          =   435
      Left            =   7440
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar Registro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Nuevo Registro"
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdsiguiente 
      Caption         =   "Siguiente Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdanterior 
      Caption         =   "Registro Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "Edad"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtapellido 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7560
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmotros.frx":240042
      OLEDBString     =   $"frmotros.frx":2400EE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Datos"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdquitar 
      Caption         =   "Quitar"
      Height          =   615
      Left            =   6000
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   6000
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmddetener 
      Caption         =   "Detener"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdpausa 
      Caption         =   "Pausar"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   1920
   End
   Begin VB.CommandButton cmdempezar 
      Caption         =   "Empezar"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lsttiempos 
      Height          =   2985
      Left            =   7080
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.OptionButton Opt3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   255
   End
   Begin VB.OptionButton Opt2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdcarpeta 
      Caption         =   "Crear Carpeta"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape tapa4 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   9015
   End
   Begin VB.Shape tapa3 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   9015
   End
   Begin VB.Shape tapa2 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   9015
   End
   Begin VB.Shape tapa1 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H000080FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   5
      X1              =   3840
      X2              =   4560
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   5
      X1              =   3840
      X2              =   4560
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label lblrespuesta 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5040
      TabIndex        =   36
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Aspa Simple"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      TabIndex        =   35
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   2
      X1              =   360
      X2              =   9360
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Desbloquear"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   240
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   375
      Left            =   8280
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   615
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
      _cx             =   1085
      _cy             =   661
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "Orden"
      DataSource      =   "Adodc1"
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
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label txtedad 
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
      Left            =   2640
      TabIndex        =   19
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Orden:"
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
      Top             =   5400
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   1
      X1              =   360
      X2              =   9360
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblsegundo 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lbldospuntos 
      BackStyle       =   0  'Transparent
      Caption         =   ":       :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblcentecima 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   5160
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblminuto 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Crear carpeta en el escritorio >>"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdanterior_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub cmdborrar_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub cmdcarpeta_Click()
On Error GoTo errSub
Dim anjes As Object
Set anjes = CreateObject("scripting.filesystemobject")
anjes.createfolder "C:\Users\Casa\Desktop\Nueva Carpeta"
MsgBox "Su carpeta ha sido creada con exito", vbInformation, "Titulo"
Exit Sub
errSub:
If Err.Number = 58 Then
   MsgBox "solo se puede crear una carpetacon ese nombre", vbCritical
End If
End Sub

Private Sub cmddetener_Click()
Timer1.Enabled = False
cmdempezar.Enabled = True
cmdempezar.Caption = "Empezar"
lblminuto = "00"
lblsegundo = "00"
lblcentecima = "00"
End Sub

Private Sub cmdempezar_Click()
Timer1.Enabled = True
cmdempezar.Caption = "continuar"
cmdempezar.Enabled = False
End Sub

Private Sub cmdgrabar_Click()
lsttiempos.AddItem (lblminuto.Caption & ":" & lblsegundo.Caption & ":" & lblcentecima.Caption)
End Sub

Private Sub cmdguardar_Click()
Adodc1.Recordset.Update
cmdguardar.Enabled = False
cmdnuevo.Enabled = True
End Sub

Private Sub cmdlimpiar_Click()
lsttiempos.Clear
End Sub

Private Sub cmdnuevo_Click()
Adodc1.Recordset.AddNew
cmdguardar.Enabled = True
cmdnuevo.Enabled = False
End Sub

Private Sub cmdpausa_Click()
Timer1.Enabled = False
cmdempezar.Enabled = True
End Sub

Private Sub cmdprimero_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdquitar_Click()
If lsttiempos.ListIndex <> -1 Then
lsttiempos.RemoveItem lsttiempos.ListIndex
End If
End Sub

Private Sub cmdsiguiente_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub cmdultimo_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdvolver_Click()
frm9.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
End Sub

Private Sub Opt1_Click()
tapa1.Visible = False
cmdcarpeta.Visible = True
tapa2.Visible = True
tapa3.Visible = True
tapa4.Visible = True

cmdempezar.Visible = False
cmdpausa.Visible = False
cmddetener.Visible = False
cmdgrabar.Visible = False
cmdquitar.Visible = False
cmdlimpiar.Visible = False
lsttiempos.Visible = False
cmddetener.Value = True

txtnombre.Visible = False
txtapellido.Visible = False
Text3.Visible = False
cmdprimero.Visible = False
cmdanterior.Visible = False
cmdsiguiente.Visible = False
cmdultimo.Visible = False
cmdnuevo.Visible = False
cmdguardar.Visible = False
cmdborrar.Visible = False

txt1.Visible = False
txt2.Visible = False
txt3.Visible = False
fraopt.Visible = False
End Sub

Private Sub Opt2_Click()
tapa1.Visible = True
cmdcarpeta.Visible = False
tapa2.Visible = False
tapa3.Visible = True
tapa4.Visible = True

cmdempezar.Visible = True
cmdpausa.Visible = True
cmddetener.Visible = True
cmdgrabar.Visible = True
cmdquitar.Visible = True
cmdlimpiar.Visible = True
lsttiempos.Visible = True

txtnombre.Visible = False
txtapellido.Visible = False
Text3.Visible = False
cmdprimero.Visible = False
cmdanterior.Visible = False
cmdsiguiente.Visible = False
cmdultimo.Visible = False
cmdnuevo.Visible = False
cmdguardar.Visible = False
cmdborrar.Visible = False

txt1.Visible = False
txt2.Visible = False
txt3.Visible = False
fraopt.Visible = False
End Sub

Private Sub Opt3_Click()
tapa1.Visible = True
cmdcarpeta.Visible = False
tapa2.Visible = True
tapa3.Visible = False
tapa4.Visible = True

cmdempezar.Visible = False
cmdpausa.Visible = False
cmddetener.Visible = False
cmdgrabar.Visible = False
cmdquitar.Visible = False
cmdlimpiar.Visible = False
lsttiempos.Visible = False
cmddetener.Value = True

txtnombre.Visible = True
txtapellido.Visible = True
Text3.Visible = True
cmdprimero.Visible = True
cmdanterior.Visible = True
cmdsiguiente.Visible = True
cmdultimo.Visible = True
cmdnuevo.Visible = True
cmdguardar.Visible = True
cmdborrar.Visible = True

txt1.Visible = False
txt2.Visible = False
txt3.Visible = False
fraopt.Visible = False
End Sub

Private Sub opt4_Click()
tapa1.Visible = True
cmdcarpeta.Visible = False
tapa2.Visible = True
tapa3.Visible = True
tapa4.Visible = False

cmdempezar.Visible = False
cmdpausa.Visible = False
cmddetener.Visible = False
cmdgrabar.Visible = False
cmdquitar.Visible = False
cmdlimpiar.Visible = False
lsttiempos.Visible = False
cmddetener.Value = True

txtnombre.Visible = False
txtapellido.Visible = False
Text3.Visible = False
cmdprimero.Visible = False
cmdanterior.Visible = False
cmdsiguiente.Visible = False
cmdultimo.Visible = False
cmdnuevo.Visible = False
cmdguardar.Visible = False
cmdborrar.Visible = False

txt1.Visible = True
txt2.Visible = True
txt3.Visible = True
fraopt.Visible = True
End Sub

Private Sub optnuevo_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
lblrespuesta.Caption = "x"
End Sub

Private Sub optrespuesta_Click()
On Error GoTo errSub
   lblrespuesta.Caption = "x = " & Val((txt3.Text) * Val(txt2.Text)) / Val(txt1.Text)
Exit Sub
errSub:
If Err.Number = 13 Then
   lblrespuesta.Caption = "Rellene los casilleros!"
End If
End Sub

Private Sub Timer1_Timer()
lblcentecima.Caption = Val(lblcentecima.Caption) + 1
If lblcentecima.Caption = "100" Then
lblcentecima.Caption = "00"
lblsegundo.Caption = "0" & Val(lblsegundo.Caption) + 1
  If lblsegundo.Caption > 9 Then
  lblsegundo.Caption = Val(lblsegundo.Caption)
  End If
End If
If lblsegundo.Caption = "60" Then
lblsegundo.Caption = "00"
lblminuto.Caption = "0" & Val(lblminuto.Caption) + 1
  If lblminuto.Caption > 9 Then
  lblminuto.Caption = Val(lblminuto.Caption)
  End If
End If
End Sub

Private Sub Timer3_Timer()
lbldospuntos.Visible = False
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
