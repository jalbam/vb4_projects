VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Millennium Time's"
   ClientHeight    =   8055
   ClientLeft      =   255
   ClientTop       =   375
   ClientWidth     =   11400
   ForeColor       =   &H0000FFFF&
   Height          =   8460
   Icon            =   "milltime.frx":0000
   Left            =   195
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   8055
   ScaleWidth      =   11400
   Top             =   30
   Width           =   11520
   Begin VB.CommandButton adelante 
      Caption         =   "----->"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   44
      Top             =   7200
      Width           =   615
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Mapa"
      ForeColor       =   &H00404080&
      Height          =   855
      Left            =   6840
      TabIndex        =   40
      Top             =   7080
      Width           =   2055
      Begin VB.CommandButton vmapa 
         Caption         =   "&Ver Mapa"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   "Situacion:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Estado Ajeno"
      ForeColor       =   &H0080C0FF&
      Height          =   6735
      Left            =   6840
      TabIndex        =   28
      Top             =   240
      Width           =   2055
      Begin VB.CommandButton Amplia 
         Caption         =   "A&mpliar"
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         ForeColor       =   &H00004080&
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   4080
         Width           =   1815
         Begin VB.PictureBox enemigo 
            BackColor       =   &H00000000&
            Height          =   1695
            Left            =   0
            ScaleHeight     =   1635
            ScaleWidth      =   1755
            TabIndex        =   43
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fuerza:"
         ForeColor       =   &H0080FF80&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "Pertenencias:"
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         Caption         =   "* Valoracion:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Agresividad:"
         ForeColor       =   &H0080FF80&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "Experiencia:"
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Resistencia Fisica:"
         ForeColor       =   &H00FF80FF&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Resistencia Magia:"
         ForeColor       =   &H00FF80FF&
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Nivel: "
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Clasificacion:"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   9120
      TabIndex        =   7
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   6495
   End
   Begin VB.CommandButton ayuda 
      Caption         =   "Ayu&da"
      Height          =   495
      Left            =   9480
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   7200
      Width           =   5415
   End
   Begin VB.Frame fravision 
      BackColor       =   &H00000000&
      Caption         =   "Vision"
      ForeColor       =   &H000040C0&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6495
      Begin VB.PictureBox vision 
         BackColor       =   &H00404040&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   6195
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.TextBox txtenemigo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   1695
      Left            =   240
      Locked          =   -1  'True
      MouseIcon       =   "milltime.frx":0442
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4800
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Estado"
      ForeColor       =   &H000080FF&
      Height          =   6735
      Left            =   9120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton Equipamiento 
         Caption         =   "&Equipamiento..."
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4320
         Width           =   1695
      End
      Begin VB.OptionButton equilibrado 
         BackColor       =   &H00000000&
         Caption         =   "&Equilibrado"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   4800
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton defensa 
         BackColor       =   &H00000000&
         Caption         =   "&Defensa"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   5640
         Width           =   975
      End
      Begin VB.OptionButton violencia 
         BackColor       =   &H00000000&
         Caption         =   "&Ofensiva"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Dinero:"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Hechizo:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   3480
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proteccion:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Arma:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Equipamiento Actual"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFC0&
         X1              =   0
         X2              =   2160
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Hechizos: "
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Proteccion: "
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Armas: "
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "* Global: "
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Experiencia: "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Magia: "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Defensa: "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Fuerza: "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Tus Niveles"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Tu Nombre: "
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3165
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3165
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   4080
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub adelante_Click()
clicks_flecha = clicks_flecha + 1
End Sub
Private Sub Amplia_Click()
Ampliado.Show
End Sub


Private Sub Equipamiento_Click()
Equipa.Show
End Sub

Private Sub Form_Activate()
If clicks_flecha = 1 Then
Form1.txtenemigo.ForeColor = "&H00FFaaFF"
Form1.txtenemigo.Text = "Tu eres El Elegido; has sido seleccionado para combatir contra El Ejercito del Bien y destruir El Diamante Perdido."
End If

If clicks_flecha = 2 Then
Form1.txtenemigo.ForeColor = "&H00bbaaFF"
Form1.txtenemigo.Text = "dsfd"
End If

End Sub

Private Sub Form_Load()
Dim tenemigo, clicks_flecha As Single
clicks_flecha = 0
'Form1.vision.Picture = LoadPicture("c:\windows\prueba.bmp")
MsgBox "Millennium Time's [0.00331 Alpha Version]", 64, "Millennium Time's 0.00331Av"

Form1.vision.Picture = LoadPicture("c:\milltime\milltime.bmp")
Form1.enemigo.Picture = LoadPicture("c:\milltime\ene1.bmp")

tenemigo = "Grafismo, Ideologo y Apoyo por DaTaR. "
tenemigo = tenemigo & "Programacion e Historia por A-Kristo. "
tenemigo = tenemigo & "Millennium Time's (c) -> "
tenemigo = tenemigo & "Una produccioncilla de: "
tenemigo = tenemigo & "Ultratumbaoz Productions (R)"
Form1.txtenemigo.Text = tenemigo
'Form1.txtenemigo.ForeColor = "&H00FF00FF"
Form1.txtenemigo.Text = tenemigo & " ...Presiona  '--->'  para Comenzar una partidilla (;"
'Ok
txt.Text = "Comando a Introducir"

End Sub


Private Sub StatusBar1_PanelClick(ByVal Panel As Panel)

End Sub


Private Sub mapa_Click()
End Sub

Private Sub salir_Click()
End
End Sub


Private Sub vmapa_Click()
mapa.Show
End Sub


