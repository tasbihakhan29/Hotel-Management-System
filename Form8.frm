VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00808080&
   Caption         =   "   HOTEL STATUS  "
   ClientHeight    =   1850
   ClientLeft      =   180
   ClientTop       =   1710
   ClientWidth     =   2980
   LinkTopic       =   "Form8"
   ScaleHeight     =   1850
   ScaleWidth      =   2980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "TO DO LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3610
      Left            =   9720
      TabIndex        =   16
      Top             =   1920
      Width           =   3850
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
         Caption         =   "  ROOM  CLEANING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   2170
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "  PAY PENDING BILLS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   2170
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "  ORDER INVENTORY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   19
         Top             =   1680
         Width           =   2890
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "  SWIMMING POOL CLEANING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   3250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "  ROOM  CLEANING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   2170
      End
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   3600
      TabIndex        =   12
      Top             =   3240
      Width           =   730
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   3600
      TabIndex        =   11
      Top             =   2640
      Width           =   730
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "GUEST INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3610
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   3970
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   2880
         TabIndex        =   15
         Top             =   2520
         Width           =   730
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2880
         TabIndex        =   14
         Top             =   1320
         Width           =   730
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   730
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   2880
         TabIndex        =   5
         Top             =   1920
         Width           =   730
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK OUT TODAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   1690
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CHILDREN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1690
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ADULTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1690
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKED IN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2530
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "ROOM INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3610
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   3490
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   1920
         Width           =   730
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ROOMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1690
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "OCCUPIED ROOMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2050
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE ROOMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   250
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   2170
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HOTEL STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   370
      Left            =   1080
      TabIndex        =   22
      Top             =   720
      Width           =   3250
   End
   Begin VB.Menu aq 
      Caption         =   "&back"
   End
   Begin VB.Menu ty 
      Caption         =   "&home"
   End
   Begin VB.Menu ag 
      Caption         =   "a&dd guest"
   End
   Begin VB.Menu ii 
      Caption         =   "&check in"
   End
   Begin VB.Menu qk 
      Caption         =   "check &out"
   End
   Begin VB.Menu ad 
      Caption         =   "b&ill"
   End
   Begin VB.Menu qe 
      Caption         =   "&guest info"
   End
   Begin VB.Menu qr 
      Caption         =   "&add room"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
   Begin VB.Menu qs 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ad_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub ag_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub aq_Click()
Unload Me
End Sub

Private Sub ff_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Form_Activate()
Dim i As Integer

Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT COUNT(*) AS rc FROM room", conn
Text6.Text = rs.Fields("rc").Value
rs.Close
rs.Open "SELECT COUNT(*) AS rv FROM checkin", conn
Text7.Text = rs.Fields("rv").Value
Text2.Text = rs.Fields("rv").Value
rs.Close
Dim rj As New ADODB.Recordset
rj.Open "SELECT addguest.adult, addguest.child FROM addguest INNER JOIN checkin ON addguest.CusID=checkin.cusid ;", conn
Dim ad As Integer
Dim ch As Integer
ad = 0
ch = 0
Do While Not rj.EOF
ad = ad + rj.Fields("adult").Value
ch = ch + rj.Fields("child").Value
rj.MoveNext
Loop
rj.Close
rs.Open "SELECT * FROM checkin", conn
Do While Not rs.EOF
If rs.Fields("checkoutdate").Value = Date Then
i = i + 1
End If
rs.MoveNext
Loop
Text4.Text = i
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Text5.Text = Val(Text6.Text) - Val(Text7.Text)
Text1.Text = ad
Text3.Text = ch
End Sub

Private Sub ii_Click()
Form4.WindowState = 2
Form4.Show
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub qe_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub qk_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub qr_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub qs_Click()
Form10.WindowState = 2
Form10.Show
End Sub

Private Sub ty_Click()
Form1.WindowState = 2
Form1.Show
End Sub
