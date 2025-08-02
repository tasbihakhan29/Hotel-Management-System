VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00808080&
   Caption         =   "SETTINGS"
   ClientHeight    =   1840
   ClientLeft      =   180
   ClientTop       =   2010
   ClientWidth     =   2980
   LinkTopic       =   "Form10"
   ScaleHeight     =   1840
   ScaleWidth      =   2980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "CHANGE SYSTEM USE&RNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6120
      Width           =   2170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "CHANG&E SETTINGS PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6720
      Width           =   2170
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "CHANGE SYSTEM &PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5520
      Width           =   2170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "SERVICE CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4930
      Left            =   9480
      TabIndex        =   37
      Top             =   480
      Width           =   4570
      Begin VB.TextBox Text20 
         Height          =   300
         Left            =   2760
         TabIndex        =   42
         Text            =   "5000"
         Top             =   1320
         Width           =   1570
      End
      Begin VB.TextBox Text19 
         Height          =   300
         Left            =   2760
         TabIndex        =   41
         Text            =   "7000"
         Top             =   1800
         Width           =   1570
      End
      Begin VB.TextBox Text18 
         Height          =   300
         Left            =   2760
         TabIndex        =   40
         Text            =   "10000"
         Top             =   2280
         Width           =   1570
      End
      Begin VB.TextBox Text17 
         Height          =   300
         Left            =   2760
         TabIndex        =   39
         Text            =   "15000"
         Top             =   2760
         Width           =   1570
      End
      Begin VB.TextBox Text16 
         Height          =   300
         Left            =   2760
         TabIndex        =   38
         Text            =   "30000"
         Top             =   3240
         Width           =   1570
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "SINGLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   49
         Top             =   1320
         Width           =   1570
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "DOUBLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1570
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "QUAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   47
         Top             =   2760
         Width           =   1570
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "TRIPLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   46
         Top             =   2280
         Width           =   1570
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "SUITE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   45
         Top             =   3240
         Width           =   1570
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Width           =   1570
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   1570
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "PERMISSIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1570
      Left            =   240
      TabIndex        =   34
      Top             =   5520
      Width           =   11530
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
         Caption         =   "   ALLOW DELETING DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   8520
         TabIndex        =   52
         Top             =   480
         Width           =   2650
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "   ALLOW EDITING DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   4800
         TabIndex        =   51
         Top             =   960
         Width           =   3730
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "   ADD PASSWORD TO SETTINGS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   4800
         TabIndex        =   50
         Top             =   480
         Value           =   1  'Checked
         Width           =   3730
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "  ALLOW CHANGING RATES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   1080
         TabIndex        =   36
         Top             =   960
         Width           =   3730
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "   ALLOW EXTRA SERVICES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   1080
         TabIndex        =   35
         Top             =   480
         Width           =   3730
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "GST AND DISCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4930
      Left            =   5400
      TabIndex        =   17
      Top             =   480
      Width           =   3970
      Begin VB.TextBox Text13 
         Height          =   300
         Left            =   2160
         TabIndex        =   31
         Text            =   "5"
         Top             =   4200
         Width           =   1570
      End
      Begin VB.TextBox Text12 
         Height          =   300
         Left            =   2160
         TabIndex        =   30
         Text            =   "10"
         Top             =   3720
         Width           =   1570
      End
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   2160
         TabIndex        =   29
         Text            =   "20"
         Top             =   3240
         Width           =   1570
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Left            =   2160
         TabIndex        =   25
         Text            =   "8"
         Top             =   2280
         Width           =   1570
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   2160
         TabIndex        =   24
         Text            =   "12"
         Top             =   1800
         Width           =   1570
      End
      Begin VB.TextBox Text8 
         Height          =   300
         Left            =   2160
         TabIndex        =   23
         Text            =   "18"
         Top             =   1320
         Width           =   1570
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "FOR DISCOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   1570
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "FOR GST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1570
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ">10000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   1570
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   ">20000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   27
         Top             =   3720
         Width           =   1570
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   ">30000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   26
         Top             =   3240
         Width           =   1570
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   ">2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   1570
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   ">7000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1570
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   ">10000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1570
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCENTAGE "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   1570
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CONDITION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "SERVICE CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4930
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4690
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   2760
         TabIndex        =   16
         Text            =   "400"
         Top             =   4200
         Width           =   1570
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   2760
         TabIndex        =   15
         Text            =   "1000"
         Top             =   3720
         Width           =   1570
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   2760
         TabIndex        =   14
         Text            =   "1000"
         Top             =   3240
         Width           =   1570
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   2760
         TabIndex        =   13
         Text            =   "2000"
         Top             =   2760
         Width           =   1570
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   2760
         TabIndex        =   12
         Text            =   "700"
         Top             =   2280
         Width           =   1570
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Text            =   "1000"
         Top             =   1800
         Width           =   1570
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2760
         TabIndex        =   10
         Text            =   "500"
         Top             =   1320
         Width           =   1570
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   3000
         TabIndex        =   9
         Top             =   720
         Width           =   1570
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   1570
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "VALET PARKING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   1570
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "CHILDCARE SERVICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   2170
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM SERVICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1570
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "GYM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1570
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SPA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   1570
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LAUNDRY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   1570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SWIMMING POOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   1570
      End
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose "
   End
   Begin VB.Menu bb 
      Caption         =   "&back"
   End
   Begin VB.Menu hh 
      Caption         =   "&home"
   End
   Begin VB.Menu dd 
      Caption         =   "a&dd guest"
   End
   Begin VB.Menu cc 
      Caption         =   "&check in"
   End
   Begin VB.Menu oo 
      Caption         =   "check &out"
   End
   Begin VB.Menu ii 
      Caption         =   "b&ill"
   End
   Begin VB.Menu gg 
      Caption         =   "&guest info"
   End
   Begin VB.Menu aa 
      Caption         =   "&add room"
   End
   Begin VB.Menu tt 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aa_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub bb_Click()
Form10.WindowState = 1
End Sub

Private Sub cc_Click()
Form4.WindowState = 2
Form4.Show
End Sub


Private Sub Check2_Click()
If Check2.Value = 1 Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text20.Enabled = True
End If
If Check2.Value = 0 Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text20.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM PASS", conn
a = InputBox("Enter old password", "Change System password")
If a = "" Then
Unload Me
Else
If a = rs.Fields("spass").Value Then
b = InputBox("Enter new password", "Change System password")
If b = "" Then
Unload Me
Else

Dim cmd As New ADODB.Command

cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "UPDATE PASS SET spass='" & b & "';"
cmd.Execute
conn.Close

Set conn = Nothing

End If
 End If
 End If

End Sub

Private Sub Command2_Click()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM PASS", conn
a = InputBox("Enter old username", "Change system username")
If a = "" Then
Unload Me
Else
If a = rs.Fields("users").Value Then
b = InputBox("Enter new username", "Change system username")
If b = "" Then
Unload Me
Else
Dim cmd As New ADODB.Command
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "UPDATE PASS SET users='" & b & "';"
cmd.Execute
conn.Close
Set conn = Nothing
End If
End If
End If
End Sub

Private Sub Command3_Click()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM PASS", conn
a = InputBox("Enter old password", "Change settings password")
If a = "" Then
Unload Me
Else
If a = rs.Fields("pass").Value Then
b = InputBox("Enter new password", "Change settings password")
If b = "" Then
Unload Me
Else
Dim cmd As New ADODB.Command
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "UPDATE PASS SET pass='" & b & "';"
cmd.Execute
conn.Close

Set conn = Nothing
End If
End If
End If
End Sub

Private Sub dd_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub ff_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Form_Activate()
If Check1.Value = 1 Then
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM PASS", conn

Do
a = InputBox("Enter password", "Settings")
If a = "" Then
Unload Me
Exit Do
End If
Loop Until a = rs.Fields("spass").Value
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
End If
If Check2.Value = 1 Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text20.Enabled = True
End If
If Check2.Value = 0 Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text20.Enabled = False
End If
End Sub



Private Sub gg_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub hh_Click()
Form1.WindowState = 2
Form1.Show
End Sub

Private Sub ii_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub oo_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub tt_Click()
Form8.WindowState = 2
Form8.Show
End Sub
