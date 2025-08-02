VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1990
   ClientLeft      =   50
   ClientTop       =   360
   ClientWidth     =   3100
   DrawMode        =   0  'Blackness
   LinkTopic       =   "Form2"
   ScaleHeight     =   8590
   ScaleWidth      =   14300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   370
      Left            =   4560
      TabIndex        =   4
      Top             =   3600
      Width           =   2770
   End
   Begin VB.TextBox Text1 
      Height          =   370
      Left            =   4560
      TabIndex        =   2
      Top             =   2760
      Width           =   2770
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LOG  IN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   490
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   1330
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   960
      TabIndex        =   3
      Top             =   3600
      Width           =   1810
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   960
      TabIndex        =   1
      Top             =   2760
      Width           =   1810
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "WELCOME  TO  DIAMOND   HOTEL"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   4570
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
