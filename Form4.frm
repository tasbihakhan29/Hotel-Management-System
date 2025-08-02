VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00808080&
   Caption         =   "GUEST CHECK IN"
   ClientHeight    =   6440
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   13270
   LinkTopic       =   "Form4"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "&VIEW/EDIT"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6720
      Width           =   2170
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   260
      Left            =   2160
      Top             =   7440
      Visible         =   0   'False
      Width           =   1090
      _ExtentX        =   1923
      _ExtentY        =   459
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "extraser"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "ROOM INFO"
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
      Height          =   1090
      Left            =   600
      TabIndex        =   23
      Top             =   2760
      Width           =   12730
      Begin VB.ComboBox Combo2 
         DataField       =   "roomid"
         DataSource      =   "Adodc1"
         Height          =   280
         Left            =   9720
         TabIndex        =   29
         Top             =   360
         Width           =   2650
      End
      Begin VB.TextBox Text15 
         DataField       =   "roomtype"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   5640
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   490
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "roomtype"
         DataSource      =   "Adodc1"
         Height          =   280
         Left            =   2880
         TabIndex        =   26
         Top             =   360
         Width           =   2410
      End
      Begin VB.TextBox Text10 
         DataField       =   "roomno"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   8760
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM TYPE"
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
         Height          =   370
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1450
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM NO"
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
         Height          =   370
         Left            =   6600
         TabIndex        =   24
         Top             =   360
         Width           =   1450
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   260
      Left            =   8280
      Top             =   4080
      Visible         =   0   'False
      Width           =   2530
      _ExtentX        =   4463
      _ExtentY        =   459
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "data"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "CHECKOUT DETAILS"
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
      Height          =   970
      Left            =   600
      TabIndex        =   14
      Top             =   5640
      Width           =   12850
      Begin VB.TextBox Text14 
         DataField       =   "checkouttime"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   9000
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   490
      End
      Begin VB.TextBox Text13 
         DataField       =   "checkoutdate"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   5280
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   490
      End
      Begin VB.TextBox Text4 
         DataField       =   "checkouttime"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   9600
         TabIndex        =   18
         Top             =   360
         Width           =   2290
      End
      Begin VB.TextBox Text3 
         DataField       =   "checkoutdate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   2290
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKOUT TIME"
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
         Height          =   370
         Left            =   6240
         TabIndex        =   17
         Top             =   360
         Width           =   1690
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKOUT DATE"
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
         Height          =   370
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1810
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "CHECK IN DETAILS"
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
      Height          =   1690
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   12730
      Begin VB.ComboBox Combo3 
         DataField       =   "cusid"
         DataSource      =   "Adodc1"
         Height          =   280
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   2290
      End
      Begin VB.TextBox Text2 
         DataField       =   "cusid"
         DataSource      =   "Adodc3"
         Height          =   370
         Left            =   5400
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.TextBox Text12 
         DataField       =   "nofdayres"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   9000
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.TextBox Text9 
         DataField       =   "tofarrival"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   8880
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.TextBox Text8 
         DataField       =   "dofarrival"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   2280
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.TextBox Text7 
         DataField       =   "cusid"
         DataSource      =   "Adodc2"
         Height          =   370
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   610
      End
      Begin VB.TextBox Text5 
         DataField       =   "nofdayres"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   9720
         TabIndex        =   13
         Top             =   360
         Width           =   2290
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "NO  OF  DAYS  RESERVED"
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
         Height          =   370
         Left            =   6480
         TabIndex        =   12
         Top             =   360
         Width           =   2530
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "tofarrival"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   9720
         TabIndex        =   11
         Top             =   960
         Width           =   2290
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME OF ARRIVAL"
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
         Height          =   370
         Left            =   6480
         TabIndex        =   10
         Top             =   960
         Width           =   1930
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dofarrival"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3000
         TabIndex        =   9
         Top             =   960
         Width           =   2290
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF ARRIVAL"
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
         Height          =   370
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1810
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER  ID"
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
         Height          =   370
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1450
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "EXTRA SERVICES"
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
      Height          =   1450
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   12730
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808080&
         Caption         =   "     VALET PARKING"
         DataField       =   "valetparking"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   10440
         TabIndex        =   32
         Top             =   240
         Width           =   2170
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00808080&
         Caption         =   "     CHILDCARE SERVICES"
         DataField       =   "childcare"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   7800
         TabIndex        =   31
         Top             =   840
         Width           =   2890
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
         Caption         =   "     ROOM SERVICES"
         DataField       =   "roomservice"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   7800
         TabIndex        =   30
         Top             =   240
         Width           =   2170
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "     GYM"
         DataField       =   "gym"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   4800
         TabIndex        =   5
         Top             =   840
         Width           =   2170
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "     SPA SERVICES"
         DataField       =   "spa"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   2170
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "     LAUNDRY SERVICES"
         DataField       =   "laundry"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2410
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "     SWIMMING POOL"
         DataField       =   "swim"
         DataSource      =   "Adodc3"
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
         Height          =   490
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2170
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "SUB&MIT"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   2170
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   370
      Left            =   360
      Top             =   7800
      Visible         =   0   'False
      Width           =   2770
      _ExtentX        =   4886
      _ExtentY        =   653
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "checkin"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "GUEST CHECK IN"
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
      Left            =   720
      TabIndex        =   37
      Top             =   360
      Width           =   3130
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose"
   End
   Begin VB.Menu aa 
      Caption         =   "&back"
   End
   Begin VB.Menu dd 
      Caption         =   "&home"
   End
   Begin VB.Menu cc 
      Caption         =   "a&dd guest"
   End
   Begin VB.Menu rr 
      Caption         =   "check &out"
   End
   Begin VB.Menu ii 
      Caption         =   "b&ill"
   End
   Begin VB.Menu jj 
      Caption         =   "&guest info"
   End
   Begin VB.Menu uu 
      Caption         =   "&add room"
   End
   Begin VB.Menu mm 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
   Begin VB.Menu vv 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aa_Click()
Form4.WindowState = 1
End Sub

Private Sub cc_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
 check ("SINGLE")
 End If
If Combo1.ListIndex = 1 Then
 check ("DOUBLE")
 End If
  If Combo1.ListIndex = 2 Then
 check ("TRIPLE")
 End If
If Combo1.ListIndex = 3 Then
 check ("QUAD")
 End If
 If Combo1.ListIndex = 4 Then
 check ("SUITE")
 End If
End Sub

Sub check(J As String)
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT room.* FROM room LEFT JOIN checkin ON room.roomid=checkin.roomid WHERE checkin.roomid IS NULL AND room.roomtype= '" & J & "'", conn
Combo2.Clear
Do While Not rs.EOF
Combo2.AddItem rs.Fields("roomid").Value
rs.MoveNext
Loop
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Update
MsgBox ("Checked in successfully")
Text7.Text = Combo3.Text
Text8.Text = Label4.Caption
Text12.Text = Text5.Text
Text9.Text = Label13.Caption
Text2.Text = Combo3.Text
Text15.Text = Combo1.Text
Text13.Text = Text3.Text
Text14.Text = Text4.Text
Text10.Text = Combo2.Text
Adodc2.Recordset.Update
Adodc3.Recordset.Update
Adodc1.Recordset.MoveLast
Adodc3.Recordset.AddNew
Adodc2.Recordset.AddNew
Adodc1.Recordset.AddNew


Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT addguest.*,checkin.* FROM addguest LEFT JOIN checkin ON addguest.CusID=checkin.cusid WHERE checkin.cusid IS NULL OR addguest.CusID IS NULL  ", conn
Combo3.Clear
Do While Not rs.EOF
Combo3.AddItem rs.Fields("addguest.CusID").Value
rs.MoveNext
Loop
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Dim conn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
conn1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs1.Open "SELECT addguest.*,checkin.* FROM addguest LEFT JOIN checkin ON addguest.CusID=checkin.cusid WHERE checkin.cusid IS NULL OR addguest.CusID IS NULL  ", conn1
Combo3.Clear
Do While Not rs1.EOF
Combo3.AddItem rs1.Fields("addguest.CusID").Value
rs1.MoveNext
Loop
rs1.Close
conn1.Close
Set rs1 = Nothing
Set conn1 = Nothing

Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Label4.Caption = Date
Label13.Caption = Time

End Sub

Private Sub Command2_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub Command5_Click()

End Sub

Private Sub dd_Click()
Form1.WindowState = 2
Form1.Show

End Sub

Private Sub ff_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Form_Activate()
Combo3.SetFocus
If Form10.Check3.Value = 0 Then
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Else
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
End If
Combo3.SetFocus
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
Adodc2.Recordset.AddNew
Adodc3.Recordset.AddNew
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT addguest.*,checkin.* FROM addguest LEFT JOIN checkin ON addguest.CusID=checkin.cusid WHERE checkin.cusid IS NULL OR addguest.CusID IS NULL ", conn
Do While Not rs.EOF
Combo3.AddItem rs.Fields("addguest.CusID").Value
rs.MoveNext
Loop
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Combo1.AddItem "SINGLE"
Combo1.AddItem "DOUBLE"
Combo1.AddItem "TRIPLE"
Combo1.AddItem "QUAD"
Combo1.AddItem "SUITE"
If Form10.Check3.Value = 0 Then
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Else
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
End If

Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Label4.Caption = Date
Label13.Caption = Time

End Sub

Private Sub ii_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub jj_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub mm_Click()
Form8.WindowState = 2
Form8.Show
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub rr_Click()
Form7.WindowState = 2
Form7.Show
End Sub








Private Sub Text2_lostfocus()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
a = Val(Text2.Text)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT roomno FROM checkin WHERE roomno= " & a, conn
If rs.EOF Then
Else
MsgBox ("Already occupied")
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

End Sub

Private Sub Text3_GotFocus()
Text3.Text = Date + Val(Text5.Text)
Text4.Text = Label13.Caption
End Sub





Private Sub uu_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub vv_Click()
Form10.WindowState = 2
Form10.Show
End Sub
