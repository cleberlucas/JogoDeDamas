VERSION 5.00
Begin VB.Form frmDamas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Damas"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblO 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   69
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   68
      Top             =   1080
      Width           =   375
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   3600
      X2              =   4920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "X   -   O"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   67
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "PLACAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   66
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Rodada do Jogador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   360
      X2              =   5160
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   360
      X2              =   5160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   5160
      X2              =   5160
      Y1              =   1560
      Y2              =   6360
   End
   Begin VB.Line Line 
      BorderWidth     =   5
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   1560
      Y2              =   6360
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   63
      Left            =   4560
      TabIndex        =   64
      Tag             =   "direita"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   62
      Left            =   3960
      TabIndex        =   63
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   61
      Left            =   3360
      TabIndex        =   62
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   60
      Left            =   2760
      TabIndex        =   61
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   59
      Left            =   2160
      TabIndex        =   60
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   58
      Left            =   1560
      TabIndex        =   59
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   57
      Left            =   960
      TabIndex        =   58
      Tag             =   "centro"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   56
      Left            =   360
      TabIndex        =   57
      Tag             =   "esquerda"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   55
      Left            =   4560
      TabIndex        =   56
      Tag             =   "direita"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   54
      Left            =   3960
      TabIndex        =   55
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   53
      Left            =   3360
      TabIndex        =   54
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   52
      Left            =   2760
      TabIndex        =   53
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   51
      Left            =   2160
      TabIndex        =   52
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   50
      Left            =   1560
      TabIndex        =   51
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   49
      Left            =   960
      TabIndex        =   50
      Tag             =   "centro"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   48
      Left            =   360
      TabIndex        =   49
      Tag             =   "esquerda"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   47
      Left            =   4560
      TabIndex        =   48
      Tag             =   "direita"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   46
      Left            =   3960
      TabIndex        =   47
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   45
      Left            =   3360
      TabIndex        =   46
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   44
      Left            =   2760
      TabIndex        =   45
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   43
      Left            =   2160
      TabIndex        =   44
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   42
      Left            =   1560
      TabIndex        =   43
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   41
      Left            =   960
      TabIndex        =   42
      Tag             =   "centro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   40
      Left            =   360
      TabIndex        =   41
      Tag             =   "esquerda"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   39
      Left            =   4560
      TabIndex        =   40
      Tag             =   "direita"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   38
      Left            =   3960
      TabIndex        =   39
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   37
      Left            =   3360
      TabIndex        =   38
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   36
      Left            =   2760
      TabIndex        =   37
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   35
      Left            =   2160
      TabIndex        =   36
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   34
      Left            =   1560
      TabIndex        =   35
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   33
      Left            =   960
      TabIndex        =   34
      Tag             =   "centro"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   32
      Left            =   360
      TabIndex        =   33
      Tag             =   "esquerda"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   31
      Left            =   4560
      TabIndex        =   32
      Tag             =   "direita"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   30
      Left            =   3960
      TabIndex        =   31
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   29
      Left            =   3360
      TabIndex        =   30
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   28
      Left            =   2760
      TabIndex        =   29
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   27
      Left            =   2160
      TabIndex        =   28
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   26
      Left            =   1560
      TabIndex        =   27
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   25
      Left            =   960
      TabIndex        =   26
      Tag             =   "centro"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   24
      Left            =   360
      TabIndex        =   25
      Tag             =   "esquerda"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   23
      Left            =   4560
      TabIndex        =   24
      Tag             =   "direita"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   22
      Left            =   3960
      TabIndex        =   23
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   21
      Left            =   3360
      TabIndex        =   22
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   20
      Left            =   2760
      TabIndex        =   21
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   19
      Left            =   2160
      TabIndex        =   20
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   18
      Left            =   1560
      TabIndex        =   19
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   17
      Left            =   960
      TabIndex        =   18
      Tag             =   "centro"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   16
      Left            =   360
      TabIndex        =   17
      Tag             =   "esquerda"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   15
      Left            =   4560
      TabIndex        =   16
      Tag             =   "direita"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   14
      Left            =   3960
      TabIndex        =   15
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   13
      Left            =   3360
      TabIndex        =   14
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   12
      Left            =   2760
      TabIndex        =   13
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   11
      Left            =   2160
      TabIndex        =   12
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   10
      Left            =   1560
      TabIndex        =   11
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   9
      Left            =   960
      TabIndex        =   10
      Tag             =   "centro"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Tag             =   "esquerda"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Tag             =   "direita"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   6
      Left            =   3960
      TabIndex        =   7
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   5
      Left            =   3360
      TabIndex        =   6
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Tag             =   "centro"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Tag             =   "esquerda"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmDamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vJogador, _
vDir, vEsq, _
pontos_j1, pontos_j2, _
posivel_posicao_superior_vEsq, _
posivel_posicao_superior_vDir, _
posivel_posicao_inferior_vEsq, _
posivel_posicao_inferior_vDir _
As Integer

Dim vPeca_removida, _
vPeçaEsq_pode_ser_removida_superior, _
vPeçaDir_pode_ser_removida_superior, _
vPeçaEsq_pode_ser_removida_inferior, _
vPeçaDir_pode_ser_removida_inferior, _
vProseguir_peça, _
return_verificacao_dama, _
vJogada_permitida, _
vProseguir_dama, inf, _
proseguir_combo, combo, dama, _
return_p As Boolean

Dim vCaption As String
Private Sub Form_Load()
    lblX.Caption = 0
    lblO.Caption = 0
    vJogador = 1
    iniciar
    Call jogador("O", False)
End Sub

Private Sub lbl_Click(Index As Integer)
    If lbl(Index) = "X" Or lbl(Index) = "O" Then
        Call possivel_jogada(Index, vJogador, False, False)
        If vJogada_permitida Then remocao (Index)
        If Not vJogada_permitida Then Call jogador(lbl(Index), True)
        vProseguir_peça = True
        
    ElseIf proseguir_combo = True Then
        Call insercao(Index, False)
        remocao_peça (Index)
        Call combo_remocao_possivel(Index, vJogador, proseguir_combo, False)
        verificacao_dama
        vProseguir_peça = True
        
    ElseIf vProseguir_peça = True Then
        Call insercao(Index, False)
        If vPeçaDir_pode_ser_removida_superior = True Or vPeçaEsq_pode_ser_removida_superior = True Then
            If posivel_posicao_superior_vEsq = Index Or posivel_posicao_superior_vDir = Index Then
                remocao_peça (Index)
                conta_pontos (vJogador)
                Call combo_remocao_possivel(Index, vJogador, proseguir_combo, False)
            Else: Proximo_Jogador
            End If
            
        Else: Proximo_Jogador
        End If
        
        vProseguir_peça = False
        verificacao_dama
            
    Else:
        If (lbl(Index) = "Y" Or lbl(Index) = "Z") And return_verificacao_dama = False Then
            Call possivel_jogada(Index, vJogador, False, True)
            If vJogada_permitida Then remocao (Index)
            If Not vJogada_permitida Then Call jogador(lbl(Index), True)
            vProseguir_dama = True
        
        ElseIf proseguir_combo = True Then
            Call insercao(Index, True)
            remocao_peça (Index)
            Call combo_remocao_possivel(Index, vJogador, proseguir_combo, True)
            vProseguir_dama = True
        
        ElseIf vProseguir_dama = True Then
            Call insercao(Index, True)
            remocao_peça (Index)
            If vPeçaDir_pode_ser_removida_superior = True Or vPeçaEsq_pode_ser_removida_superior = True _
            Or vPeçaDir_pode_ser_removida_inferior = True Or vPeçaEsq_pode_ser_removida_inferior = True Then
            
                If posivel_posicao_superior_vEsq = Index Or posivel_posicao_superior_vDir = Index _
                Or posivel_posicao_inferior_vEsq = Index Or posivel_posicao_inferior_vDir = Index Then
                    remocao_peça (Index)
                    conta_pontos (vJogador)
                    If lbl(Index).tag = "centro" Then Call combo_remocao_possivel(Index, vJogador, proseguir_combo, True)
                   
                Else: Proximo_Jogador
                End If
                
            Else: Proximo_Jogador
            End If
            
            
            vProseguir_dama = False
            verificacao_dama
        Else: return_verificacao_dama = False
        End If
        
    End If
End Sub

Sub remocao_peça(Index)
        If posivel_posicao_superior_vEsq Or posivel_posicao_superior_vDir = Index Then
            If vPeçaDir_pode_ser_removida_superior = True And (lbl(Index - vDir).BackColor = &H80000010 Or &H8080FF) Then lbl(Index - vDir).Caption = ""
            If vPeçaEsq_pode_ser_removida_superior = True And (lbl(Index - vEsq).BackColor = &H80000010 Or &H8080FF) Then lbl(Index - vEsq).Caption = ""
        End If
         
        If posivel_posicao_inferior_vEsq Or posivel_posicao_inferior_vDir = Index Then
            Call troca_direcao(vDir, vEsq)
            If vPeçaDir_pode_ser_removida_inferior = True And (lbl(Index - vDir).BackColor = &H80000010 Or &H8080FF) Then lbl(Index - vDir).Caption = ""
            If vPeçaEsq_pode_ser_removida_inferior = True And (lbl(Index - vEsq).BackColor = &H80000010 Or &H8080FF) Then lbl(Index - vEsq).Caption = "     "
            Call troca_direcao(vDir, vEsq)
        End If
        
        vPeçaDir_pode_ser_removida_superior = False
        vPeçaEsq_pode_ser_removida_superior = False
        vPeçaDir_pode_ser_removida_inferior = False
        vPeçaEsq_pode_ser_removida_inferior = False
        
End Sub


Sub combo_remocao_possivel(Index, vJogador, proseguir_combo, dama)
    Call possivel_jogada(Index, vJogador, True, dama)
    
            If vPeçaDir_pode_ser_removida_superior Or vPeçaEsq_pode_ser_removida_superior = True _
            Or vPeçaDir_pode_ser_removida_inferior Or vPeçaEsq_pode_ser_removida_inferior = True Then
                remocao (Index)
                proseguir_combo = True
   
            Else
                voltar_cores
                Proximo_Jogador
                proseguir_combo = False
            End If
End Sub

Sub conta_pontos(vJogador)
    If vJogador = 1 Then pontos_j1 = pontos_j1 + 1
    If vJogador = 2 Then pontos_j2 = pontos_j2 + 1
    lblX.Caption = pontos_j1
    lblO.Caption = pontos_j2
    
End Sub

Sub verificacao_dama()
    For I = 0 To 63
        If lbl(I).Caption = "X" And I <= 7 Then
            lbl(I).Caption = "Y"
            return_verificacao_dama = True
        End If
        If lbl(I).Caption = "O" And I >= 56 Then
            lbl(I).Caption = "Z"
            return_verificacao_dama = True
        End If
    Next I
    
End Sub


Sub jogador_EsqDir(vJogador)
For I = 0 To 63
        lbl(I).Enabled = False
    Next I
   
    If vJogador = 1 Then
        vDir = (-7)
        vEsq = (-9)
        vCaption = "X"
    ElseIf vJogador = 2 Then
        vDir = (9)
        vEsq = (7)
        vCaption = "O"
    End If
    
End Sub

Sub troca_direcao(vDir, vEsq)
    If vDir = (9) Then
        vDir = (-7)
    ElseIf vDir = (-7) Then vDir = (9)
    End If
    
    If vEsq = (7) Then
        vEsq = (-9)
    ElseIf vEsq = (-9) Then vEsq = (7)
    End If
End Sub

Sub possivel_jogada(Index, vJogador, combo, dama)
    Call jogador_EsqDir(vJogador)

    On Error Resume Next
    If lbl(Index).tag = "centro" Then Call movimentacao("direita", vDir, Index, vCaption, combo, False)
    If lbl(Index).tag = "centro" Then Call movimentacao("esquerda", vEsq, Index, vCaption, combo, False)
    If lbl(Index).tag = "esquerda" Then Call movimentacao("direita", vDir, Index, vCaption, combo, False)
    If lbl(Index).tag = "direita" Then Call movimentacao("esquerda", vEsq, Index, vCaption, combo, False)
    
    If dama = True Then
        Call troca_direcao(vDir, vEsq)
        If lbl(Index).tag = "centro" Then Call movimentacao("direita", vDir, Index, vCaption, combo, True)
        If lbl(Index).tag = "centro" Then Call movimentacao("esquerda", vEsq, Index, vCaption, combo, True)
        If lbl(Index).tag = "esquerda" Then Call movimentacao("direita", vDir, Index, vCaption, combo, True)
        If lbl(Index).tag = "direita" Then Call movimentacao("esquerda", vEsq, Index, vCaption, combo, True)
        Call troca_direcao(vDir, vEsq)
    End If
End Sub

Function verifica_casa(posicao) As Boolean
    verifica_casa = True
    If lbl(posicao).Caption = "X" Then verifica_casa = False
    If lbl(posicao).Caption = "O" Then verifica_casa = False
    If lbl(posicao).Caption = "Z" Then verifica_casa = False
    If lbl(posicao).Caption = "Y" Then verifica_casa = False
End Function

Sub movimentacao(tag, esq_dir, Index, vCaption, combo, inf)
    If verifica_casa(Index + esq_dir) _
    And Not lbl(Index + esq_dir).BackColor = &HFFFFFF _
    And combo = False Then
        lbl(Index + esq_dir).BackColor = &H80000010
        lbl(Index + esq_dir).Enabled = True
        vJogada_permitida = True
    
    Else
        If verifica_casa(Index + esq_dir * 2) _
        And Not lbl(Index + esq_dir).Caption = vCaption _
        And Not lbl(Index + esq_dir).Caption = "" _
        And Not lbl(Index + esq_dir).tag = tag _
        And Not lbl(Index + esq_dir * 2).BackColor = &HFFFFFF Then
            lbl(Index + esq_dir).BackColor = &H8080FF
            lbl(Index + esq_dir * 2).BackColor = &H80000010
            lbl(Index + esq_dir * 2).Enabled = True
            vJogada_permitida = True
            
            If inf = False Then
                If esq_dir = vEsq Then
                    vPeçaEsq_pode_ser_removida_superior = True
                    posivel_posicao_superior_vEsq = (Index + (esq_dir * 2))
                ElseIf esq_dir = vDir Then
                    vPeçaDir_pode_ser_removida_superior = True
                    posivel_posicao_superior_vDir = (Index + (esq_dir * 2))
                End If
            End If
                
            If inf = True Then
                If esq_dir = vEsq Then
                    vPeçaEsq_pode_ser_removida_inferior = True
                    posivel_posicao_inferior_vEsq = (Index + (esq_dir * 2))
                ElseIf esq_dir = vDir Then
                    vPeçaDir_pode_ser_removida_inferior = True
                    posivel_posicao_inferior_vDir = (Index + (esq_dir * 2))
                End If
            End If
            
        End If
    End If
End Sub

Sub iniciar()
    For I = 0 To 63
            If lbl(I).BackColor = &H8000000F Then
                If I <= 23 Then lbl(I).Caption = "O"
                If I >= 40 Then lbl(I).Caption = "X"
            ElseIf lbl(I).BackColor = &HFFFFFF Then lbl(I).Enabled = False  'Cores brancas
            End If
        Next I
End Sub

Sub jogador(XO As String, atividade As Boolean)
    For I = 0 To 63
    If lbl(I).Caption = XO Then lbl(I).Enabled = atividade
    Next I
End Sub

Sub Proximo_Jogador()
    If vJogador = 1 Then
        vJogador = 2
        Label.Caption = vJogador
        Call jogador("X", False)
        Call jogador("Y", False)
        Call jogador("O", True)
        Call jogador("Z", True)
    Else
        vJogador = 1
        Label.Caption = vJogador
        Call jogador("O", False)
        Call jogador("Z", False)
        Call jogador("X", True)
        Call jogador("Y", True)
    End If
End Sub

Sub remocao(Index)
        lbl(Index).Caption = ""
End Sub

Sub voltar_cores()
    For I = 0 To 63
        If lbl(I).BackColor = &H80000010 Or lbl(I).BackColor = &H8080FF Then lbl(I).BackColor = &H8000000F 'cinza escuro e vermelho -> cinza claro
    Next I
End Sub

Sub insercao(Index, dama)
    If vJogador = 1 Then lbl(Index).Caption = "X"
    If vJogador = 2 Then lbl(Index).Caption = "O"
    If vJogador = 1 And dama = True Then lbl(Index).Caption = "Y"
    If vJogador = 2 And dama = True Then lbl(Index).Caption = "Z"
    voltar_cores
    vJogada_permitida = False
End Sub

