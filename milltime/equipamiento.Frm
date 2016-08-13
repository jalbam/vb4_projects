VERSION 4.00
Begin VB.Form Equipa 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipamiento"
   ClientHeight    =   8055
   ClientLeft      =   1215
   ClientTop       =   570
   ClientWidth     =   10590
   Height          =   8460
   Icon            =   "equipamiento.frx":0000
   Left            =   1155
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10590
   Top             =   225
   Width           =   10710
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Valoracion Total"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   735
      Left            =   8040
      TabIndex        =   48
      Top             =   7200
      Width           =   2415
   End
   Begin VB.PictureBox Picture23 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Height          =   855
      Left            =   3240
      ScaleHeight     =   795
      ScaleWidth      =   675
      TabIndex        =   45
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox Picture22 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Height          =   855
      Left            =   1560
      ScaleHeight     =   795
      ScaleWidth      =   675
      TabIndex        =   44
      Top             =   5640
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "MAGIA"
      ForeColor       =   &H00FFFF00&
      Height          =   6375
      Left            =   8040
      TabIndex        =   37
      Top             =   600
      Width           =   2415
      Begin VB.PictureBox Picture21 
         BackColor       =   &H00404040&
         ForeColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   240
         ScaleHeight     =   1755
         ScaleWidth      =   1875
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "MEDIA:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   240
         TabIndex        =   42
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Eficacia:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   300
         Left            =   240
         TabIndex        =   41
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Desgaste:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   300
         Left            =   240
         TabIndex        =   40
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Destruccion:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   300
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   5160
      ScaleHeight     =   675
      ScaleWidth      =   2715
      TabIndex        =   19
      Top             =   7200
      Width           =   2775
   End
   Begin VB.PictureBox Picture19 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      Height          =   2895
      Left            =   5400
      ScaleHeight     =   2835
      ScaleWidth      =   2235
      TabIndex        =   18
      Top             =   3960
      Width           =   2295
   End
   Begin VB.PictureBox Picture18 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      Height          =   2895
      Left            =   5400
      ScaleHeight     =   2835
      ScaleWidth      =   2235
      TabIndex        =   17
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox Picture17 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   3120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   16
      Top             =   4920
      Width           =   735
   End
   Begin VB.PictureBox Picture16 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   4920
      Width           =   735
   End
   Begin VB.PictureBox Picture15 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004080&
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   14
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox Picture14 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004080&
      Height          =   1215
      Left            =   1320
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   3120
      Width           =   615
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   1080
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2295
      Left            =   1920
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00332200&
      Height          =   735
      Left            =   3360
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00332200&
      Height          =   735
      Left            =   1200
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004040&
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Calzado Interior"
      ForeColor       =   &H00FF80FF&
      Height          =   195
      Left            =   4080
      TabIndex        =   47
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Calzado Interior"
      ForeColor       =   &H00FF80FF&
      Height          =   195
      Left            =   360
      TabIndex        =   46
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Municion"
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   6240
      TabIndex        =   36
      Top             =   6960
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Brazo Izquierdo"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   6000
      TabIndex        =   35
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Brazo Derecho"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   6000
      TabIndex        =   34
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Armamento"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6120
      TabIndex        =   33
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Calzado"
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   32
      Top             =   7080
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Calzado"
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   600
      TabIndex        =   31
      Top             =   7080
      Width           =   570
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodilla"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodilla"
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   1800
      TabIndex        =   29
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hombro"
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   3480
      TabIndex        =   28
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Hombro"
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   1440
      TabIndex        =   27
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cabeza"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Codo"
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   3960
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Codo"
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   1080
      TabIndex        =   24
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mano"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   23
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mano"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Anillos"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4200
      TabIndex        =   21
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Anillos"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   20
      Top             =   4560
      Width           =   450
   End
End
Attribute VB_Name = "Equipa"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
