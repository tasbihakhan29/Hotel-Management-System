VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00808080&
   Caption         =   "GUEST INFO"
   ClientHeight    =   1850
   ClientLeft      =   180
   ClientTop       =   2010
   ClientWidth     =   2980
   LinkTopic       =   "Form9"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   370
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   2650
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   850
      Left            =   6720
      Top             =   5880
      Visible         =   0   'False
      Width           =   3490
      _ExtentX        =   6156
      _ExtentY        =   1499
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":0000
      Height          =   5650
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   13930
      _ExtentX        =   24571
      _ExtentY        =   9966
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8421504
      ForeColor       =   -2147483634
      HeadLines       =   1
      RowHeight       =   29
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GUEST DATA"
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2410
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2410
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
   Begin VB.Menu ss 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aa_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub bb_Click()
Unload Me
End Sub

Private Sub cc_Click()
Form4.WindowState = 2
Form4.Show
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
Text1.SetFocus
If Form10.Check4 = 0 Then
DataGrid1.AllowUpdate = False
Else
DataGrid1.AllowUpdate = True
End If
If Form10.Check5 = 0 Then
DataGrid1.AllowDelete = False
Else
DataGrid1.AllowDelete = True
End If
End Sub

Private Sub hh_Click()
Form1.WindowState = 2
Form1.Show
End Sub

Private Sub ii_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub oo_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub ss_Click()
Form10.WindowState = 2
Form10.Show
End Sub

Private Sub Text1_LostFocus()
Dim i As Integer
i = 0
Dim f As Boolean
f = False
Dim conn1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT CusID FROM addguest ", conn1
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields("CusID").Value = Val(Text1.Text) Then
DataGrid1.Row = i
f = True

Exit Do
End If
rs.MoveNext
i = i + 1
Loop

End If
If Not f Then

End If
rs.Close
conn1.Close
End Sub

Private Sub tt_Click()
Form8.WindowState = 2
Form8.Show
End Sub
