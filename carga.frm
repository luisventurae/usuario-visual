VERSION 5.00
Begin VB.Form frmcargar 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cargando..."
   ClientHeight    =   6960
   ClientLeft      =   2595
   ClientTop       =   2310
   ClientWidth     =   9105
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcerrar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "cerrar"
      Height          =   255
      Left            =   7440
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   720
      Top             =   7440
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1680
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1200
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   720
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   4440
   End
   Begin VB.Timer Timerpase 
      Interval        =   950
      Left            =   7680
      Top             =   5400
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   7560
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   1680
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   1440
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   4200
      X2              =   4200
      Y1              =   1440
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   10
      X1              =   3720
      X2              =   3720
      Y1              =   1680
      Y2              =   2520
   End
   Begin VB.Line Linesubrayado 
      BorderColor     =   &H00FFFFFF&
      X1              =   6960
      X2              =   8880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblcancelar 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "ZapfEllipt BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7200
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblConectando 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando..."
      BeginProperty Font 
         Name            =   "Zurich BlkEx BT"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   360
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   0
      Top             =   3000
      Width           =   8655
   End
   Begin VB.Image image 
      Height          =   11520
      Left            =   -1800
      MousePointer    =   13  'Arrow and Hourglass
      Picture         =   "carga.frx":0000
      Top             =   -2880
      Width           =   15360
   End
End
Attribute VB_Name = "frmcargar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcerrar_Click()
End
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub image_Click()
Timerpase.Enabled = False
frmcargar.Hide
frm1.Show
End Sub

Private Sub lblcancelar_Click()
Timerpase.Enabled = False
Shape1.FillColor = RGB(0, 0, 0)
If Shape1.FillColor = RGB(0, 255, 255) Then
frmcargar.Hide
frm0.Show
End If
If Shape1.FillColor = &HC0FFC0 Then
frmcargar.Hide
Form1.Show
End If
End Sub

Private Sub Timer1_Timer()
Line1.Visible = False
Line2.Visible = True
Line3.Visible = False
Line4.Visible = False
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = False
Timer4.Enabled = False
If frm1.cmdingresar.Value = True Then
Timerpase.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Line1.Visible = False
Line2.Visible = False
Line3.Visible = True
Line4.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = False
If frm0.cmdingresar.Value = True Then
Timerpase.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Line1.Visible = True
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Shape1.FillColor = RGB(0, 0, 1)
End Sub

Private Sub Timerpase_Timer()
If Shape1.FillColor = &HC0FFC0 Then
frmcargar.Hide
frm2.Show
End If
If Shape1.FillColor = RGB(0, 255, 255) Then
frmcargar.Hide
Form1.Show
End If
Shape1.FillColor = RGB(0, 0, 0)
End Sub
