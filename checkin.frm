VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   Caption         =   "ADD GUEST"
   ClientHeight    =   1850
   ClientLeft      =   180
   ClientTop       =   2010
   ClientWidth     =   2980
   LinkTopic       =   "Form3"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000010&
      Caption         =   "&UPDATE"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   1210
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "DE&LETE"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6120
      Width           =   1330
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "&EDIT"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   1090
   End
   Begin VB.TextBox Text9 
      Height          =   370
      Left            =   9600
      TabIndex        =   24
      Top             =   840
      Width           =   2650
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "checkin.frx":0000
      Height          =   4570
      Left            =   7200
      TabIndex        =   23
      Top             =   1440
      Width           =   6610
      _ExtentX        =   11659
      _ExtentY        =   8061
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8421504
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   24
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "ADDITIONAL INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2170
      Left            =   480
      TabIndex        =   15
      Top             =   4440
      Width           =   6490
      Begin VB.TextBox Text8 
         DataField       =   "phone"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1440
         Width           =   2170
      End
      Begin VB.TextBox Text5 
         DataField       =   "email"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   19
         Top             =   960
         Width           =   3730
      End
      Begin VB.TextBox Text4 
         DataField       =   "add"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   3730
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1810
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "PERSONAL INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3610
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   6490
      Begin VB.TextBox Text7 
         DataField       =   "child"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   14
         Top             =   2400
         Width           =   1570
      End
      Begin VB.TextBox Text6 
         DataField       =   "adult"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   12
         Top             =   1920
         Width           =   1570
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         Height          =   280
         Left            =   2520
         TabIndex        =   10
         Top             =   2880
         Width           =   1570
      End
      Begin VB.TextBox Text3 
         DataField       =   "lname"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   8
         Top             =   1440
         Width           =   2890
      End
      Begin VB.TextBox Text2 
         DataField       =   "fname"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   2890
      End
      Begin VB.TextBox Text1 
         DataField       =   "CusID"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   1570
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NO. OF CHILDREN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   1810
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "NO. OF ADULTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1810
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   1810
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1810
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   370
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER ID"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1810
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000010&
      Caption         =   "CHEC&K IN"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1330
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1210
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   610
      Left            =   2160
      Top             =   7680
      Visible         =   0   'False
      Width           =   2770
      _ExtentX        =   4886
      _ExtentY        =   1076
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
      RecordSource    =   "addguest"
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH CUSTOMER ID"
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
      Height          =   370
      Left            =   7080
      TabIndex        =   25
      Top             =   840
      Width           =   2410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD GUEST"
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
      TabIndex        =   22
      Top             =   120
      Width           =   2290
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose"
   End
   Begin VB.Menu hm 
      Caption         =   "&back"
   End
   Begin VB.Menu gt 
      Caption         =   "&home"
   End
   Begin VB.Menu df 
      Caption         =   "&check in"
   End
   Begin VB.Menu gl 
      Caption         =   "check &out"
   End
   Begin VB.Menu ss 
      Caption         =   "b&ill"
   End
   Begin VB.Menu gi 
      Caption         =   "&guest info"
   End
   Begin VB.Menu tt 
      Caption         =   "&add room"
   End
   Begin VB.Menu qw 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "Customer in&fo"
   End
   Begin VB.Menu st 
      Caption         =   "&Settings"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim J As Integer

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Combo1.Text = "" Then
MsgBox ("Can't update Enter all values")
Else
Adodc1.Recordset.Update

Adodc1.Recordset.AddNew
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM addguest ORDER BY CusID DESC", conn
If rs.EOF Then
Else
If J = 0 Then
MsgBox ("Record Submitted")
Text1.Text = rs.Fields("CusID").Value + 1
Else
Text1.Text = rs.Fields("CusID").Value + 1
J = 0
End If
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Command3.Enabled = False
End If
End Sub


Private Sub Command5_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Combo1.Text = "" Then
MsgBox ("Can't update Enter all values")
Else
Adodc1.Recordset.Update

Adodc1.Recordset.AddNew
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM addguest ORDER BY CusID DESC", conn
If rs.EOF Then
Else
If J = 0 Then
MsgBox ("Record Updated")
Text1.Text = rs.Fields("CusID").Value + 1
Else
Text1.Text = rs.Fields("CusID").Value + 1
J = 0
End If
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Command3.Enabled = False
End If
End Sub

Private Sub ff_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Form_Activate()
If Form10.Check4.Value = 0 Then
Text9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Else
Text9.Enabled = True
Command2.Enabled = True
End If
Text2.SetFocus
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub Text9_LostFocus()
Dim i As Integer
i = 0
Dim f As Boolean
f = False
Dim conn1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT CusID FROM addguest", conn1
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields("CusID").Value = Val(Text9.Text) Then
Adodc1.Recordset.AbsolutePosition = i + 1
J = 1
f = True
Exit Do
End If
rs.MoveNext
i = i + 1
Loop
End If
If Not f Then
MsgBox ("Data not found")
End If
rs.Close
conn1.Close
End Sub
Private Sub Command2_Click()
Command3.Enabled = True
Adodc1.Recordset.Delete
Text9.SetFocus
End Sub

Private Sub Command3_Click()
If Text9.Text = "" Then
MsgBox ("select a cusid")
Else
Adodc1.Recordset.Delete
Text9.Text = ""
End If
End Sub

Private Sub Command4_Click()
Form4.WindowState = 2
Form4.Show
End Sub

Private Sub df_Click()
Form4.WindowState = 2
Form4.Show
End Sub



Private Sub Form_Load()
Adodc1.Recordset.MoveLast
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM addguest ORDER BY CusID DESC", conn
If rs.EOF Then
Else
Text1.Text = rs.Fields("CusID").Value
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Combo1.AddItem "MALE"
Combo1.AddItem "FEMALE"
End Sub

Private Sub gi_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub gl_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub gt_Click()
Form1.WindowState = 2
Form1.Show
End Sub
Private Sub hm_Click()
Form3.WindowState = 1
End Sub

Private Sub qw_Click()
Form8.WindowState = 2
Form8.Show
End Sub

Private Sub ss_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub st_Click()
Form10.WindowState = 2
Form10.Show
End Sub





Private Sub tt_Click()
Form5.WindowState = 2
Form5.Show
End Sub
