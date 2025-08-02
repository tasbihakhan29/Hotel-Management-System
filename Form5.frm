VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00808080&
   Caption         =   "ADD ROOM"
   ClientHeight    =   1990
   ClientLeft      =   180
   ClientTop       =   1710
   ClientWidth     =   3100
   LinkTopic       =   "Form5"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000010&
      Caption         =   "SUB&MIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   1810
   End
   Begin VB.TextBox Text9 
      Height          =   370
      Left            =   5520
      TabIndex        =   13
      Top             =   3960
      Width           =   2650
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "&EDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1810
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1810
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   730
      Left            =   1320
      Top             =   7800
      Width           =   3130
      _ExtentX        =   5521
      _ExtentY        =   1288
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
      RecordSource    =   "room"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":0000
      Height          =   3370
      Left            =   2760
      TabIndex        =   9
      Top             =   4440
      Width           =   5890
      _ExtentX        =   10389
      _ExtentY        =   5944
      _Version        =   393216
      BackColor       =   -2147483642
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin VB.TextBox Text3 
      DataField       =   "roomrate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   370
      Left            =   6120
      TabIndex        =   8
      Top             =   3120
      Width           =   2290
   End
   Begin VB.TextBox Text2 
      DataField       =   "noofbeds"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   370
      Left            =   6120
      TabIndex        =   7
      Top             =   2400
      Width           =   2290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "&UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1810
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "roomtype"
      DataSource      =   "Adodc1"
      Height          =   280
      ItemData        =   "Form5.frx":0015
      Left            =   6120
      List            =   "Form5.frx":0017
      TabIndex        =   5
      Top             =   1800
      Width           =   2530
   End
   Begin VB.TextBox Text1 
      DataField       =   "roomid"
      DataSource      =   "Adodc1"
      Height          =   370
      Left            =   6120
      TabIndex        =   1
      Top             =   1080
      Width           =   2290
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH ROOM ID"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Width           =   2410
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD ROOM"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   480
      Width           =   2650
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM RATE"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   2650
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. OF BEDS"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   2650
   End
   Begin VB.Label Label2 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   2650
   End
   Begin VB.Label Label1 
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   2650
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose"
   End
   Begin VB.Menu bb 
      Caption         =   "&back "
   End
   Begin VB.Menu hh 
      Caption         =   "&home"
   End
   Begin VB.Menu bl 
      Caption         =   "&add guest"
   End
   Begin VB.Menu mm 
      Caption         =   "&check in"
   End
   Begin VB.Menu hl 
      Caption         =   "check &out"
   End
   Begin VB.Menu vv 
      Caption         =   "b&ill"
   End
   Begin VB.Menu gg 
      Caption         =   "&guest info"
   End
   Begin VB.Menu dd 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
   Begin VB.Menu oo 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim J As Integer
Private Sub bb_Click()
Form5.WindowState = 1
End Sub
Private Sub bl_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Text2.Text = 1
Text3.Text = Form10.Text20.Text
End If
If Combo1.ListIndex = 1 Then
Text2.Text = 2
Text3.Text = Form10.Text19.Text
End If
If Combo1.ListIndex = 2 Then
Text2.Text = 3
Text3.Text = Form10.Text18.Text
End If
If Combo1.ListIndex = 3 Then
Text2.Text = 4
Text3.Text = Form10.Text17.Text
End If
If Combo1.ListIndex = 4 Then
Text2.Text = 2
Text3.Text = Form10.Text16.Text
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text1.Text = "" Or Combo1.Text = "" Or Text3.Text = "" Then
MsgBox ("Can't update Insert all values")
Combo1.SetFocus
Else
Adodc1.Recordset.Update
Adodc1.Recordset.AddNew
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM room ORDER BY roomid DESC", conn
If rs.EOF Then
Else
If J = 0 Then
MsgBox ("Room added successfully")
Text1.Text = rs.Fields("roomid").Value + 1
Else
Text1.Text = rs.Fields("roomid").Value + 1
J = 0
End If
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Text9.Text = ""
Command3.Enabled = False
End If
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
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT roomid FROM checkin where roomid=" & Text1.Text, conn
If rs.EOF Then
Adodc1.Recordset.Delete
Else
MsgBox ("Occupied room cannot be deleted")
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
End If

End Sub


Private Sub Command4_Click()
If Text1.Text = "" Or Text1.Text = "" Or Combo1.Text = "" Or Text3.Text = "" Then
MsgBox ("Can't update Insert all values")
Combo1.SetFocus
Else
Adodc1.Recordset.Update
Adodc1.Recordset.AddNew
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM room ORDER BY roomid DESC", conn
If rs.EOF Then
Else
If J = 0 Then
MsgBox ("Room added successfully")
Text1.Text = rs.Fields("roomid").Value + 1
Else
Text1.Text = rs.Fields("roomid").Value + 1
J = 0
End If
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Text9.Text = ""
Command3.Enabled = False
End If
End Sub

Private Sub dd_Click()
Form8.WindowState = 2
Form8.Show
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

End Sub

Private Sub Form_Load()
Adodc1.Recordset.MoveLast
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM room ORDER BY roomid DESC", conn
If rs.EOF Then
Else
Text1.Text = rs.Fields("roomid").Value
End If
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Combo1.AddItem "Single"
Combo1.AddItem "Double"
Combo1.AddItem "Triple"
Combo1.AddItem "Quad"
Combo1.AddItem "Suite"
End Sub

Private Sub gg_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub hh_Click()
Form1.WindowState = 2
Form1.Show
End Sub

Private Sub hl_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub mm_Click()
Form4.WindowState = 2
Form4.Show
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub oo_Click()
Form10.WindowState = 2
Form10.Show
End Sub

Private Sub Text9_LostFocus()
Dim i As Integer
i = 0
Dim f As Boolean
f = False
Dim conn1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT roomid FROM room", conn1
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields("roomid").Value = Val(Text9.Text) Then
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

Private Sub vv_Click()
Form6.WindowState = 2
Form6.Show
End Sub
