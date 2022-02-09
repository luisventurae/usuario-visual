VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm4 
   Caption         =   "calculadora"
   ClientHeight    =   5805
   ClientLeft      =   3765
   ClientTop       =   1755
   ClientWidth     =   8955
   LinkTopic       =   "Form4"
   ScaleHeight     =   5805
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Borrar resultado"
      Default         =   -1  'True
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdcien 
      Caption         =   "Sacar Porcentaje"
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdentre 
      Caption         =   "/"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdpor 
      Caption         =   "x"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdmenos 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdmas 
      Caption         =   "+"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdcuadrado 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      Picture         =   "calcu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdcubo 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      Picture         =   "calcu.frx":0380
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdraiz2 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      Picture         =   "calcu.frx":06DD
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdraiz3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      Picture         =   "calcu.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton optporcentaje 
      Caption         =   "Option1"
      Height          =   195
      Left            =   6600
      TabIndex        =   23
      Top             =   2400
      Width           =   135
   End
   Begin VB.OptionButton optmayor 
      Caption         =   "Option2"
      Height          =   195
      Left            =   2040
      TabIndex        =   18
      Top             =   3120
      Width           =   255
   End
   Begin VB.OptionButton optsigno 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   2520
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Timer tmrhorafecha 
      Interval        =   1
      Left            =   240
      Top             =   4800
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "volver a Menú"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtr 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Text            =   "0"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "S/."
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
      Left            =   5880
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar Porcentajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6840
      TabIndex        =   24
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar (al cuadrado) y (raiz cuadrada)"
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
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar +, -, x, /"
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
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
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
      _cx             =   2355
      _cy             =   873
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Caption         =   "hora"
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
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado :"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2° Numero :"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1° Numero :"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculadora Semicientífica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -1920
      Picture         =   "calcu.frx":1031
      Top             =   -4440
      Width           =   15360
   End
End
Attribute VB_Name = "frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdborrar_Click()
txtr.Text = "0"
End Sub

Private Sub cmdcien_Click()
txtr.Text = (Val(txt1.Text) * Val(txt2.Text)) / 100
End Sub

Private Sub cmdcuadrado_Click()
txtr.Text = (Val(txt1.Text) ^ (2))
End Sub

Private Sub cmdcubo_Click()
txtr.Text = (Val(txt1.Text) ^ (3))
End Sub

Private Sub cmdentre_Click()
On Error GoTo errSub
   txtr.Text = Val(txt1.Text) / Val(txt2.Text)
Exit Sub
errSub:
If Err.Number = 11 Then
   txtr.Text = "No se puede dividir entre 0"
End If
End Sub

Private Sub cmdmas_Click()
txtr.Text = Val(txt1.Text) + Val(txt2.Text)
End Sub


Private Sub cmdmenos_Click()
txtr.Text = Val(txt1.Text) - Val(txt2.Text)
End Sub



Private Sub cmdmenu_Click()
'limpiar campo
txt1.Text = ""
txt2.Text = ""
txtr.Text = ""

'cambiar campo
frm4.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
End Sub

Private Sub cmdpor_Click()
txtr.Text = Val(txt1.Text) * Val(txt2.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub cmdraiz2_Click()
txtr.Text = (Val(txt1.Text) ^ (1 / 2))
End Sub

Private Sub cmdraiz3_Click()
txtr.Text = (Val(txt1.Text) ^ (1 / 3))
End Sub


Private Sub Command5_Click()

End Sub

Private Sub Label6_Click()
frm4.Hide
frmcalender.Show
End Sub

Private Sub Option1_Click()

End Sub

Private Sub optmayor_Click()
If optmayor.Value = True Then
cmdmas.Enabled = False
cmdmenos.Enabled = False
cmdpor.Enabled = False
cmdentre.Enabled = False
cmdcuadrado.Enabled = True
cmdcubo.Enabled = True
cmdraiz2.Enabled = True
cmdraiz3.Enabled = True
txt2.Enabled = False
txt2.Text = ""
cmdcien.Visible = False
Label2.Left = 360
Label2.Top = 1080
Label3.Left = 360
Label3.Top = 1680
Label10.Visible = False
Label11.Visible = False
Label2.Caption = "1° Numero:"
Label3.Caption = "2° Numero:"
Else
cmdmas.Enabled = True
cmdmenos.Enabled = True
cmdpor.Enabled = True
cmdentre.Enabled = True
cmdcuadrado.Enabled = False
cmdcubo.Enabled = False
cmdraiz2.Enabled = False
cmdraiz3.Enabled = False
txt2.Enabled = True
End If
End Sub

Private Sub optporcentaje_Click()
If optporcentaje.Value = True Then
Label2.Caption = "El .."
Label2.Left = 1440
Label2.Top = 1080
Label3.Caption = "De .. "
Label3.Left = 1440
Label3.Top = 1560
Label10.Visible = True
Label11.Visible = True
cmdcien.Visible = True
cmdmas.Enabled = False
cmdmenos.Enabled = False
cmdpor.Enabled = False
cmdentre.Enabled = False
cmdcuadrado.Enabled = False
cmdcubo.Enabled = False
cmdraiz2.Enabled = False
cmdraiz3.Enabled = False
txt2.Enabled = True
End If
End Sub

Private Sub optsigno_Click()
If optsigno.Value = True Then
cmdmas.Enabled = True
cmdmenos.Enabled = True
cmdpor.Enabled = True
cmdentre.Enabled = True
cmdcuadrado.Enabled = False
cmdcubo.Enabled = False
cmdraiz2.Enabled = False
cmdraiz3.Enabled = False
txt2.Enabled = True
cmdcien.Visible = False
Label2.Left = 360
Label2.Top = 1080
Label3.Left = 360
Label3.Top = 1680
Label10.Visible = False
Label11.Visible = False
Label2.Caption = "1° Numero:"
Label3.Caption = "2° Numero:"
Else
cmdmas.Enabled = False
cmdmenos.Enabled = False
cmdpor.Enabled = False
cmdentre.Enabled = False
cmdcuadrado.Enabled = True
cmdcubo.Enabled = True
cmdraiz2.Enabled = True
cmdraiz3.Enabled = True
txt2.Enabled = False
End If
End Sub

Private Sub tmrhorafecha_Timer()
lblhora.Caption = Time
lblfecha.Caption = Date
End Sub
