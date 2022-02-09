VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frm6 
   Caption         =   "Traductor"
   ClientHeight    =   5565
   ClientLeft      =   3510
   ClientTop       =   2580
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdlimpiar 
      Caption         =   "limpiar"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdtraducir 
      Caption         =   "Traducir"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3135
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(usar principalmente verbos)"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   6600
      TabIndex        =   6
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A Ingles"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De español"
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
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -3360
      Picture         =   "traduc.frx":0000
      Top             =   -2880
      Width           =   15360
   End
   Begin VB.Menu marchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mvolver 
         Caption         =   "volver a escoger"
      End
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdlimpiar_Click()
Text2.Text = ""
Text1.Text = ""
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\clic.wma"
wmp1.URL = sonido
wmp1.Visible = False
End Sub

Private Sub cmdtraducir_Click()
'De Español a Inglés
If Text1.Text = "yo" Then
Text2.Text = "I"
End If
If Text1.Text = "tú" Then
Text2.Text = "you"
End If
If Text1.Text = "él" Then
Text2.Text = "he"
End If
If Text1.Text = "ella" Then
Text2.Text = "she"
End If
If Text1.Text = "nosotros" Then
Text2.Text = "we"
End If
If Text1.Text = "ellos" Then
Text2.Text = "they"
End If
If Text1.Text = "nosotras" Then
Text2.Text = "we"
End If
If Text1.Text = "ellas" Then
Text2.Text = "they"
End If
If Text1.Text = "mi nombre es luis" Then
Text2.Text = "my name is Luis"
End If
If Text1.Text = "tu nombre es gerson" Then
Text2.Text = "your name is Gerson"
End If
If Text1.Text = "tu nombre es teo" Then
Text2.Text = "your name is Teo"
End If
If Text1.Text = "tu nombre es gloria" Then
Text2.Text = "your name is Gloria"
End If
If Text1.Text = "tu nombre es jose" Then
Text2.Text = "your name is José"
End If
If Text1.Text = "tu nombre es ricardo" Then
Text2.Text = "your name is Ricardo"
End If
If Text1.Text = "tu nombre es melisa" Then
Text2.Text = "your name is Melisa"
End If
If Text1.Text = "tu nombre es eduardo" Then
Text2.Text = "your name is Eduardo"
End If
If Text1.Text = "tu nombre es orlando" Then
Text2.Text = "your name is Orlandoo"
End If
If Text1.Text = "tu nombre es emilio" Then
Text2.Text = "your name is Emilio"
End If
If Text1.Text = "tu nombre es fernando" Then
Text2.Text = "your name is Fernando"
End If
If Text1.Text = "tu nombre es jackelin" Then
Text2.Text = "your name is Jackeline"
End If
If Text1.Text = "tu nombre es nataly" Then
Text2.Text = "your name is Nataly"
End If

If Text1.Text = "tu nombre es gianpier" Then
Text2.Text = "your name is Gianpier"
End If
If Text1.Text = "tu nombre es alexandra" Then
Text2.Text = "your name is Alexandra"
End If
If Text1.Text = "tu nombre es teresa" Then
Text2.Text = "your name is Teresa"
End If
If Text1.Text = "tu nombre es leoncio" Then
Text2.Text = "your name is Leoncio"
End If
If Text1.Text = "tu nombre es ines" Then
Text2.Text = "your name is Inés"
End If
If Text1.Text = "tu nombre es graciela" Then
Text2.Text = "your name is Graciela"
End If
If Text1.Text = "tu nombre es sandra" Then
Text2.Text = "your name is Sandra"
End If
If Text1.Text = "tu nombre es sebastian" Then
Text2.Text = "your name is Sebastian"
End If
If Text1.Text = "tu nombre es ellen" Then
Text2.Text = "your name is Ellen"
End If
If Text1.Text = "tu nombre es leslie" Then
Text2.Text = "your name is Leslie"
End If
If Text1.Text = "tu nombre es anthy" Then
Text2.Text = "your name is Anthy"
End If
If Text1.Text = "tu nombre es miriam" Then
Text2.Text = "your name is Miriam"
End If
If Text1.Text = "yo soy luis" Then
Text2.Text = "I'm Luis"
End If
If Text1.Text = "yo vivo en perú" Then
Text2.Text = "I'm live in Perú"
End If
If Text1.Text = "mi colegio es julio cesar escobar" Then
Text2.Text = "my school is Julio Cesar Escobar"
End If
If Text1.Text = "tengo 15 años" Then
Text2.Text = "I'm 15 years old"
End If
If Text1.Text = "mi salon es el 5°D" Then
Text2.Text = "my classroom is 5°D"
End If
If Text1.Text = "somos la promocion 2012" Then
Text2.Text = "we are the promotion 2012"
End If
If Text1.Text = "hola" Then
Text2.Text = "hello"
End If
If Text1.Text = "hola" Then
Text2.Text = "hi"
End If
If Text1.Text = "buenos dias" Then
Text2.Text = "good morning"
End If
If Text1.Text = "buenas tardes" Then
Text2.Text = "good afternoon"
End If
If Text1.Text = "buenas noches" Then
Text2.Text = "good night"
End If
If Text1.Text = "tu eres jose" Then
Text2.Text = "you are Jose"
End If
If Text1.Text = "chau" Then
Text2.Text = "bye"
End If
If Text1.Text = "buena suerte" Then
Text2.Text = "good luck"
End If
If Text1.Text = "¿como estas?" Then
Text2.Text = "how are you?"
End If
If Text1.Text = "hasta luego" Then
Text2.Text = "so long"
End If
If Text1.Text = "hasta mañana" Then
Text2.Text = "till tomorrow"
End If
If Text1.Text = "¿cual es tu nombre?" Then
Text2.Text = "what is your name?"
End If
If Text1.Text = "la" Then
Text2.Text = "the"
End If
If Text1.Text = "el" Then
Text2.Text = "the"
End If
If Text1.Text = "las" Then
Text2.Text = "the"
End If
If Text1.Text = "los" Then
Text2.Text = "the"
End If
If Text1.Text = "la mesa" Then
Text2.Text = "the table"
End If
If Text1.Text = "el calenadario" Then
Text2.Text = "the calender"
End If
If Text1.Text = "la ventana" Then
Text2.Text = "the window"
End If
If Text1.Text = "el perro" Then
Text2.Text = "the dog"
End If
If Text1.Text = "la computadora" Then
Text2.Text = "the computer"
End If
If Text1.Text = "el colegio" Then
Text2.Text = "the school"
End If
If Text1.Text = "el lapicero" Then
Text2.Text = "the pen"
End If
If Text1.Text = "el cuaderno" Then
Text2.Text = "the notebook"
End If
If Text1.Text = "la biblioteca" Then
Text2.Text = "the library"
End If
If Text1.Text = "el borrador" Then
Text2.Text = "the eraser"
End If
If Text1.Text = "el lapiz" Then
Text2.Text = "the pencil"
End If
If Text1.Text = "el salon" Then
Text2.Text = "the classroom"
End If
If Text1.Text = "muy bien" Then
Text2.Text = "very well"
End If
If Text1.Text = "gracias" Then
Text2.Text = "thank you"
End If
If Text1.Text = "amigo" Then
Text2.Text = "friend"
End If
If Text1.Text = "profesor" Then
Text2.Text = "teacher"
End If
If Text1.Text = "bienvenido" Then
Text2.Text = "welcome"
End If
If Text1.Text = "¿que tal?" Then
Text2.Text = "how are you?"
End If
If Text1.Text = "hasta pronto" Then
Text2.Text = "till"
End If

If Text1.Text = "correr" Then
Text2.Text = "run"
End If
If Text1.Text = "caminar" Then
Text2.Text = "walk"
End If
If Text1.Text = "hablar" Then
Text2.Text = "speak"
End If
If Text1.Text = "nadar" Then
Text2.Text = "swim"
End If
If Text1.Text = "escuchar" Then
Text2.Text = "listen"
End If
If Text1.Text = "mirar" Then
Text2.Text = "look"
End If
If Text1.Text = "estudiar" Then
Text2.Text = "study"
End If
If Text1.Text = "empezar" Then
Text2.Text = "start"
End If
If Text1.Text = "escribir" Then
Text2.Text = "write"
End If
If Text1.Text = "leer" Then
Text2.Text = "read"
End If
If Text1.Text = "observar" Then
Text2.Text = "observe"
End If
If Text1.Text = "poner" Then
Text2.Text = "put"
End If
If Text1.Text = "colocar" Then
Text2.Text = "place"
End If
If Text1.Text = "comprobar" Then
Text2.Text = "check"
End If
If Text1.Text = "ver" Then
Text2.Text = "see"
End If
If Text1.Text = "comer" Then
Text2.Text = "eat"
End If
If Text1.Text = "beber" Then
Text2.Text = "drink"
End If
If Text1.Text = "tomar" Then
Text2.Text = "take"
End If
If Text1.Text = "escoger" Then
Text2.Text = "choose"
End If
If Text1.Text = "grabar" Then
Text2.Text = "record"
End If
If Text1.Text = "lavar" Then
Text2.Text = "wash"
End If
If Text1.Text = "cocinar" Then
Text2.Text = "cook"
End If
If Text1.Text = "molestar" Then
Text2.Text = "disturb"
End If
If Text1.Text = "salir" Then
Text2.Text = "leave"
End If
If Text1.Text = "vivir" Then
Text2.Text = "live"
End If
If Text1.Text = "viajar" Then
Text2.Text = "travel"
End If
If Text1.Text = "cantar" Then
Text2.Text = "sing"
End If
If Text1.Text = "cerrar" Then
Text2.Text = "close"
End If
If Text1.Text = "minimizar" Then
Text2.Text = "minimize"
End If
If Text1.Text = "maximizar" Then
Text2.Text = "maximize"
End If
If Text1.Text = "hacer" Then
Text2.Text = "do"
End If
If Text1.Text = "tener" Then
Text2.Text = "have"
End If
If Text1.Text = "lanzar" Then
Text2.Text = "launch"
End If
If Text1.Text = "trabajar" Then
Text2.Text = "work"
End If
If Text1.Text = "bailar" Then
Text2.Text = "dance"
End If
If Text1.Text = "cenar" Then
Text2.Text = "dine"
End If
If Text1.Text = "vender" Then
Text2.Text = "sell"
End If
If Text1.Text = "comprar" Then
Text2.Text = "buy"
End If
If Text1.Text = "mover" Then
Text2.Text = "move"
End If
If Text1.Text = "cambiar" Then
Text2.Text = "change"
End If
If Text1.Text = "alumbrar" Then
Text2.Text = "light"
End If
If Text1.Text = "crear" Then
Text2.Text = "maker"
End If
If Text1.Text = "excluir" Then
Text2.Text = "exclude"
End If
If Text1.Text = "borrar" Then
Text2.Text = "delete"
End If
If Text1.Text = "botar" Then
Text2.Text = "throw"
End If
If Text1.Text = "votar" Then
Text2.Text = "vote"
End If
If Text1.Text = "alcanzar" Then
Text2.Text = "achieve"
End If
If Text1.Text = "extraer" Then
Text2.Text = "extract"
End If
If Text1.Text = "amar" Then
Text2.Text = "love"
End If
If Text1.Text = "querer" Then
Text2.Text = "want"
End If
If Text1.Text = "ir" Then
Text2.Text = "go"
End If
If Text1.Text = "seleccionar" Then
Text2.Text = "select"
End If
If Text1.Text = "aceptar" Then
Text2.Text = "accept"
End If
If Text1.Text = "oscurecer" Then
Text2.Text = "obscure"
End If
If Text1.Text = "decir" Then
Text2.Text = "say"
End If
If Text1.Text = "cargar" Then
Text2.Text = "load"
End If
If Text1.Text = "montar" Then
Text2.Text = "mount"
End If
If Text1.Text = "pegar" Then
Text2.Text = "paste"
End If
If Text1.Text = "golpear" Then
Text2.Text = "hit"
End If
If Text1.Text = "condtruir" Then
Text2.Text = "build"
End If
If Text1.Text = "romper" Then
Text2.Text = "break"
End If
If Text1.Text = "felicitar" Then
Text2.Text = "congratulate"
End If
If Text1.Text = "comentar" Then
Text2.Text = "comment"
End If
If Text1.Text = "talar" Then
Text2.Text = "cut"
End If
If Text1.Text = "cortar" Then
Text2.Text = "cut"
End If
If Text1.Text = "quemar" Then
Text2.Text = "burn"
End If
If Text1.Text = "dormir" Then
Text2.Text = "sleep"
End If
If Text1.Text = "soñar" Then
Text2.Text = "dream"
End If
If Text1.Text = "echar" Then
Text2.Text = "cast"
End If
If Text1.Text = "sacar" Then
Text2.Text = "get"
End If
If Text1.Text = "atar" Then
Text2.Text = "tie"
End If
If Text1.Text = "amarrar" Then
Text2.Text = "moor"
End If
If Text1.Text = "huir" Then
Text2.Text = "escape"
End If
If Text1.Text = "escapar" Then
Text2.Text = "escape"
End If
If Text1.Text = "ganar" Then
Text2.Text = "win"
End If
If Text1.Text = "desayunar" Then
Text2.Text = "dreakfast"
End If
If Text1.Text = "enseñar" Then
Text2.Text = "tech"
End If
If Text1.Text = "deletrear" Then
Text2.Text = "spell"
End If
If Text1.Text = "presumir" Then
Text2.Text = "presume"
End If
If Text1.Text = "imprimir" Then
Text2.Text = "print"
End If
If Text1.Text = "dibujar" Then
Text2.Text = "sketch"
End If
If Text1.Text = "pintar" Then
Text2.Text = "paint"
End If
If Text1.Text = "pelar" Then
Text2.Text = "peel"
End If
If Text1.Text = "inflar" Then
Text2.Text = "inflate"
End If
If Text1.Text = "rayar" Then
Text2.Text = "scratch"
End If
If Text1.Text = "limpiar" Then
Text2.Text = "clean"
End If
If Text1.Text = "ordenar" Then
Text2.Text = "order"
End If
If Text1.Text = "desordenar" Then
Text2.Text = "mess"
End If
If Text1.Text = "robar" Then
Text2.Text = "steal"
End If
If Text1.Text = "asustar" Then
Text2.Text = "scare"
End If
If Text1.Text = "devolver" Then
Text2.Text = "return"
End If
If Text1.Text = "barrer" Then
Text2.Text = "sweep"
End If
If Text1.Text = "raspar" Then
Text2.Text = "scrape"
End If
If Text1.Text = "arrancar" Then
Text2.Text = "start"
End If
If Text1.Text = "terminar" Then
Text2.Text = "end"
End If
If Text1.Text = "calificar" Then
Text2.Text = "quality"
End If
If Text1.Text = "extrañar" Then
Text2.Text = "surprise"
End If
If Text1.Text = "bienestar" Then
Text2.Text = "wilfare"
End If
If Text1.Text = "cuidar" Then
Text2.Text = "care"
End If
If Text1.Text = "contar" Then
Text2.Text = "count"
End If
If Text1.Text = "pagar" Then
Text2.Text = "pay"
End If
If Text1.Text = "orar" Then
Text2.Text = "pray"
End If
If Text1.Text = "rezar" Then
Text2.Text = "pray"
End If
If Text1.Text = "explotar" Then
Text2.Text = "exploit"
End If
If Text1.Text = "anotar" Then
Text2.Text = "annotate"
End If
If Text1.Text = "ayudar" Then
Text2.Text = "help"
End If
If Text1.Text = "estudiar" Then
Text2.Text = "stuty"
End If
If Text1.Text = "pensar" Then
Text2.Text = "think"
End If
If Text1.Text = "arreglar" Then
Text2.Text = "fix"
End If
If Text1.Text = "abrazar" Then
Text2.Text = "embrace"
End If
If Text1.Text = "resolver" Then
Text2.Text = "solve"
End If
If Text1.Text = "preguntar" Then
Text2.Text = "ask"
End If
If Text1.Text = "importante" Then
Text2.Text = "important"
End If
If Text1.Text = "importar" Then
Text2.Text = "import"
End If
If Text1.Text = "exportar" Then
Text2.Text = "export"
End If
If Text1.Text = "llevar" Then
Text2.Text = "lead"
End If
If Text1.Text = "responder" Then
Text2.Text = "answer"
End If
If Text1.Text = "conseguir" Then
Text2.Text = "get"
End If
If Text1.Text = "llamar" Then
Text2.Text = "call"
End If
If Text1.Text = "jugar" Then
Text2.Text = "play"
End If
If Text1.Text = "mentir" Then
Text2.Text = "lie"
End If
If Text1.Text = "conversar" Then
Text2.Text = "talk"
End If
If Text1.Text = "chatear" Then
Text2.Text = "chat"
End If
If Text1.Text = "abrir" Then
Text2.Text = "open"
End If
If Text1.Text = "buscar" Then
Text2.Text = "search"
End If
If Text1.Text = "detallar" Then
Text2.Text = "detail"
End If
If Text1.Text = "escanear" Then
Text2.Text = "scan"
End If
If Text1.Text = "estar" Then
Text2.Text = "be"
End If
If Text1.Text = "entrar" Then
Text2.Text = "enter"
End If
If Text1.Text = "examinar" Then
Text2.Text = "review"
End If
If Text1.Text = "aprender" Then
Text2.Text = "learn"
End If
If Text1.Text = "elegir" Then
Text2.Text = "choose"
End If
If Text1.Text = "reparar" Then
Text2.Text = "repair"
End If
If Text1.Text = "quejar" Then
Text2.Text = "complain"
End If
If Text1.Text = "opacar" Then
Text2.Text = "overshadow"
End If
If Text1.Text = "reir" Then
Text2.Text = "laugh"
End If
If Text1.Text = "sonreir" Then
Text2.Text = "smile"
End If
If Text1.Text = "jugar" Then
Text2.Text = "play"
End If
If Text1.Text = "malograr" Then
Text2.Text = "spoil"
End If
If Text1.Text = "curar" Then
Text2.Text = "cure"
End If
If Text1.Text = "accionar" Then
Text2.Text = "actuate"
End If
If Text1.Text = "jalar" Then
Text2.Text = "pull"
End If
If Text1.Text = "enfermar" Then
Text2.Text = "sicken"
End If
If Text1.Text = "sanar" Then
Text2.Text = "heal"
End If
If Text1.Text = "enfocar" Then
Text2.Text = "focus"
End If
If Text1.Text = "desenfocar" Then
Text2.Text = "blur"
End If
If Text1.Text = "pedir" Then
Text2.Text = "ask"
End If
If Text1.Text = "caer" Then
Text2.Text = "fall"
End If

End Sub


Private Sub mvolver_Click()
frm6.Hide
frm2.Show
sonido = "D:\Fotos de Luis\word\Tarea de Luis (EPT)\calculadora a cuenta\Usuario-visual\audio\regresando.wma"
wmp1.URL = sonido
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
