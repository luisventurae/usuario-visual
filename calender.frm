VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmcalender 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario"
   ClientHeight    =   6600
   ClientLeft      =   3330
   ClientTop       =   2505
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdmenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "volver al men�"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      _Version        =   524288
      _ExtentX        =   14420
      _ExtentY        =   11668
      _StockProps     =   1
      BackColor       =   12648384
      Year            =   2012
      Month           =   10
      Day             =   18
      DayLength       =   2
      MonthLength     =   2
      DayFontColor    =   8421376
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483646
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmcalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmenu_Click()
frmcalender.Hide
frm2.Show
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmcalender.Hide
frm2.Show
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
