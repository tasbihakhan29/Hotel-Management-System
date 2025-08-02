VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00808080&
   Caption         =   "CHECK OUT"
   ClientHeight    =   1850
   ClientLeft      =   180
   ClientTop       =   2010
   ClientWidth     =   2980
   LinkTopic       =   "Form7"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Height          =   610
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   2890
   End
   Begin VB.CommandButton Command2 
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
      Height          =   610
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   3250
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "CHEC&K OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2890
   End
   Begin VB.TextBox Text1 
      Height          =   370
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   2770
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   4330
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   12370
      _ExtentX        =   21819
      _ExtentY        =   7638
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   850
      Left            =   10440
      Top             =   6600
      Visible         =   0   'False
      Width           =   2770
      _ExtentX        =   4886
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK OUT CUSTOMER"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   3850
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
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2410
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose"
   End
   Begin VB.Menu bb 
      Caption         =   "&back"
   End
   Begin VB.Menu dd 
      Caption         =   "&home"
   End
   Begin VB.Menu cy 
      Caption         =   "a&dd guest"
   End
   Begin VB.Menu zz 
      Caption         =   "&check in"
   End
   Begin VB.Menu bm 
      Caption         =   "b&ill"
   End
   Begin VB.Menu gg 
      Caption         =   "&guest info"
   End
   Begin VB.Menu gf 
      Caption         =   "&add room"
   End
   Begin VB.Menu ar 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
   Begin VB.Menu at 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ar_Click()
Form8.WindowState = 2
Form8.Show
End Sub

Private Sub at_Click()
Form10.WindowState = 2
Form10.Show
End Sub

Private Sub bb_Click()
Form7.WindowState = 1
End Sub

Private Sub bm_Click()
Form6.WindowState = 2
Form6.Show
End Sub

Private Sub Command1_Click()

Dim conn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rv As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs1.Open "select * from bill where cusid=" & Text1.Text, conn
rv.Open "select checkoutdate  from checkin where cusid=" & Text1.Text, conn
If Not rv.EOF Then
 If rv.Fields("checkoutdate").Value <> Date Then
 MsgBox ("Can't checkout")
 Else
 If rs1.EOF Then
  MsgBox ("Bill not calculated")
 Else
  Adodc1.Recordset.Delete
  End If
End If
End If
Text1.SetFocus
 
rv.Close
rs1.Close


End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
MsgBox ("RECORD UPDATED")
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Command2.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Form7.WindowState = 1
End Sub

Private Sub cy_Click()
Form3.WindowState = 2
Form3.Show
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
Text1.SetFocus
Command2.Enabled = False
If Form10.Check4 = 0 Then
DataGrid1.AllowUpdate = False
Else
DataGrid1.AllowUpdate = True
End If
End Sub

Private Sub gf_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub gg_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub Text1_LostFocus()
Dim i As Integer

i = 0
Dim f As Boolean
f = False
Dim conn1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn1.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT cusid FROM checkin ", conn1
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If rs.Fields("cusid").Value = Val(Text1.Text) Then
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

Private Sub zz_Click()
Form4.WindowState = 2
Form4.Show
End Sub
