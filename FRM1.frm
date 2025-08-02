VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HOME PAGE"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   710
   ClientWidth     =   11230
   LinkTopic       =   "Form1"
   Picture         =   "FRM1.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   11230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000010&
      Caption         =   "CHECK OUT &DATA"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   2530
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000010&
      Caption         =   "BILL I&NFO"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1570
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000010&
      Caption         =   "G&UEST INFO"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1690
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000010&
      Caption         =   "&HOTEL STATUS"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   2040
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000010&
      Caption         =   "ADD &GUEST"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1570
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000010&
      Caption         =   "B&ILLING"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1450
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "&ADD ROOM"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1690
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "CHECK &OUT"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1570
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "&CHECK IN"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   1800
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1450
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "College :GOVT. Polytechnic , Yavatmal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   240
      TabIndex        =   13
      Top             =   7680
      Width           =   3610
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Guided by : Aziz Khan (Director IQRA SOLUTIONs)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   240
      TabIndex        =   12
      Top             =   7440
      Width           =   4930
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Industrial Training Centre: IQRA SOLUTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   240
      TabIndex        =   11
      Top             =   7200
      Width           =   4090
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Tasbiha khan "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   240
      TabIndex        =   10
      Top             =   6960
      Width           =   2530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HOTEL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   370
      Left            =   5520
      TabIndex        =   9
      Top             =   600
      Width           =   5410
   End
   Begin VB.Image Image10 
      Height          =   970
      Left            =   4680
      Picture         =   "FRM1.frx":29D42
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1450
   End
   Begin VB.Image Image9 
      Height          =   970
      Left            =   5160
      Picture         =   "FRM1.frx":525E3
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1450
   End
   Begin VB.Image Image8 
      Height          =   970
      Left            =   1680
      Picture         =   "FRM1.frx":55616
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1450
   End
   Begin VB.Image Image7 
      Height          =   970
      Left            =   8880
      Picture         =   "FRM1.frx":6BEC7
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1450
   End
   Begin VB.Image Image6 
      Height          =   970
      Left            =   10200
      Picture         =   "FRM1.frx":7770B
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1450
   End
   Begin VB.Image Image3 
      Height          =   970
      Left            =   6840
      Picture         =   "FRM1.frx":79241
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1450
   End
   Begin VB.Image Image1 
      Height          =   970
      Left            =   3360
      Picture         =   "FRM1.frx":7B4A6
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1450
   End
   Begin VB.Image Image5 
      Height          =   1090
      Left            =   11160
      Picture         =   "FRM1.frx":7DA3C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1450
   End
   Begin VB.Image Image4 
      Height          =   970
      Left            =   8040
      Picture         =   "FRM1.frx":7FF72
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1450
   End
   Begin VB.Image Image2 
      Height          =   490
      Left            =   720
      Top             =   2400
      Width           =   850
   End
   Begin VB.Image Image11 
      Height          =   8520
      Left            =   0
      Picture         =   "FRM1.frx":81999
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17020
   End
   Begin VB.Menu ll 
      Caption         =   "&log out"
   End
   Begin VB.Menu bb 
      Caption         =   "&back"
   End
   Begin VB.Menu ss 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bb_Click()
Form2.WindowState = 1
End Sub

Private Sub Command1_Click()
Form4.WindowState = 2
Form4.Show
End Sub

Private Sub Command10_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Command2_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub Command3_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub Command4_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub Command5_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub Command6_Click()
Form8.WindowState = 2
Form8.Show
End Sub



Private Sub Command7_Click()

Form10.WindowState = 2
Form10.Show
End Sub

Private Sub Command8_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub Command9_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub ll_Click()
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Form2.Text3.Text = ""
Form2.Text4.Text = ""
Unload Me
Form2.Show

End Sub

Private Sub ss_Click()
Form10.WindowState = 2
Form10.Show
End Sub
