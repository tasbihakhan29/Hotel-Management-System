VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00808080&
   Caption         =   "BILL"
   ClientHeight    =   1990
   ClientLeft      =   180
   ClientTop       =   1710
   ClientWidth     =   3100
   LinkTopic       =   "Form6"
   ScaleHeight     =   8150
   ScaleWidth      =   14180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "BILL"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   3720
      Width           =   11890
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000010&
         Caption         =   "SUB&MIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2400
         Width           =   2170
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000010&
         Caption         =   "&VIEW/EDIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Width           =   2170
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000010&
         Caption         =   "ADD &NEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1200
         Width           =   2170
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000010&
         Caption         =   "&PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   2170
      End
      Begin VB.TextBox Text6 
         DataField       =   "roomcharges"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Width           =   2530
      End
      Begin VB.TextBox Text7 
         DataField       =   "eservicechar"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3960
         TabIndex        =   16
         Top             =   960
         Width           =   2530
      End
      Begin VB.TextBox Text8 
         DataField       =   "tex"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3960
         TabIndex        =   15
         Top             =   1560
         Width           =   2530
      End
      Begin VB.TextBox Text9 
         DataField       =   "discount"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3960
         TabIndex        =   14
         Top             =   2160
         Width           =   2530
      End
      Begin VB.TextBox Text10 
         DataField       =   "total"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   3960
         TabIndex        =   13
         Top             =   2760
         Width           =   2530
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   3480
         TabIndex        =   35
         Top             =   2760
         Width           =   370
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   3480
         TabIndex        =   34
         Top             =   2160
         Width           =   370
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   3480
         TabIndex        =   33
         Top             =   1560
         Width           =   370
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   3480
         TabIndex        =   32
         Top             =   360
         Width           =   370
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs."
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
         Left            =   3480
         TabIndex        =   31
         Top             =   960
         Width           =   370
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM CAHRGES"
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
         TabIndex        =   22
         Top             =   360
         Width           =   2170
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "DISOUNT"
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
         TabIndex        =   21
         Top             =   2160
         Width           =   2170
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX"
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
         TabIndex        =   20
         Top             =   1560
         Width           =   2170
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "EXTRA SERVICE CHARGES"
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
         TabIndex        =   19
         Top             =   960
         Width           =   3250
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL BILL"
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
         TabIndex        =   18
         Top             =   2760
         Width           =   2170
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   370
      Left            =   1200
      Top             =   240
      Visible         =   0   'False
      Width           =   970
      _ExtentX        =   1711
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
      RecordSource    =   "bill"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "CUSTOMER INFO"
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
      Height          =   2890
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   11890
      Begin VB.ComboBox Combo1 
         DataField       =   "cusid"
         DataSource      =   "Adodc1"
         Height          =   280
         Left            =   2640
         TabIndex        =   27
         Top             =   1080
         Width           =   2890
      End
      Begin VB.TextBox Text11 
         Height          =   370
         Left            =   8760
         TabIndex        =   24
         Top             =   1680
         Width           =   2530
      End
      Begin VB.TextBox Text3 
         DataField       =   "nofmem"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   8760
         TabIndex        =   11
         Top             =   1080
         Width           =   2290
      End
      Begin VB.TextBox Text4 
         DataField       =   "roomid"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   8760
         TabIndex        =   10
         Top             =   480
         Width           =   2290
      End
      Begin VB.TextBox Text12 
         Height          =   370
         Left            =   2640
         TabIndex        =   7
         Top             =   2280
         Width           =   2290
      End
      Begin VB.TextBox Text2 
         Height          =   370
         Left            =   2640
         TabIndex        =   5
         Top             =   1680
         Width           =   3370
      End
      Begin VB.TextBox Text5 
         DataField       =   "ID"
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   2290
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKIN DATE"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   2170
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NO. OF MEMBERS"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   2170
      End
      Begin VB.Label Label5 
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
         Left            =   6240
         TabIndex        =   8
         Top             =   480
         Width           =   2170
      End
      Begin VB.Label Label12 
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
         TabIndex        =   6
         Top             =   2280
         Width           =   2170
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER FULL NAME"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2650
      End
      Begin VB.Label Label1 
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
         TabIndex        =   2
         Top             =   1080
         Width           =   2170
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL ID"
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
         TabIndex        =   1
         Top             =   480
         Width           =   2170
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   370
      Left            =   1560
      TabIndex        =   28
      Top             =   240
      Width           =   2770
   End
   Begin VB.Menu ll 
      Caption         =   "c&lose"
   End
   Begin VB.Menu bb 
      Caption         =   "&back"
   End
   Begin VB.Menu hg 
      Caption         =   "&home"
   End
   Begin VB.Menu ck 
      Caption         =   "a&dd guest"
   End
   Begin VB.Menu cx 
      Caption         =   "&check in"
   End
   Begin VB.Menu qg 
      Caption         =   "check &out"
   End
   Begin VB.Menu bx 
      Caption         =   "&guest info"
   End
   Begin VB.Menu vc 
      Caption         =   "&add room"
   End
   Begin VB.Menu ci 
      Caption         =   "ho&tel status"
   End
   Begin VB.Menu nn 
      Caption         =   "bill i&nfo"
   End
   Begin VB.Menu ff 
      Caption         =   "customer in&fo"
   End
   Begin VB.Menu vy 
      Caption         =   "&settings"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer
Private Sub bb_Click()
Form6.WindowState = 1
End Sub

Private Sub bx_Click()
Form9.WindowState = 2
Form9.Show
End Sub

Private Sub ci_Click()
Form8.WindowState = 2
Form8.Show
End Sub

Private Sub ck_Click()
Form3.WindowState = 2
Form3.Show
End Sub

Private Sub Command1_Click()
Form6.PrintForm
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Text2.Text = ""
Text12.Text = ""
Text11.Text = ""
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM bill ORDER BY ID DESC", conn
If rs.EOF Then
Else
Text5.Text = rs.Fields("ID").Value + 1
End If
rs.Close
conn.Close
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT checkin.*,bill.* FROM checkin LEFT JOIN bill ON checkin.cusid=bill.cusid WHERE checkin.cusid IS NULL OR bill.Cusid IS NULL  ", conn
Combo1.Clear
Do While Not rs.EOF
Combo1.AddItem rs.Fields("checkin.cusid").Value
rs.MoveNext
Loop
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
End Sub

Private Sub Command3_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command5_Click()
Form6.WindowState = 1
End Sub

Private Sub cx_Click()
Form4.WindowState = 2
Form4.Show
End Sub

Private Sub ff_Click()
Form12.WindowState = 2
Form12.Show
End Sub

Private Sub Form_Activate()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM bill ORDER BY ID DESC", conn
If rs.EOF Then
Else
Text5.Text = rs.Fields("ID").Value + 1
End If
rs.Close
conn.Close
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT checkin.*,bill.* FROM checkin LEFT JOIN bill ON checkin.cusid=bill.cusid WHERE checkin.cusid IS NULL OR bill.Cusid IS NULL  ", conn
Combo1.Clear
Do While Not rs.EOF
Combo1.AddItem rs.Fields("checkin.cusid").Value
rs.MoveNext
Loop
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
End Sub

Private Sub Form_Load()

Adodc1.Recordset.AddNew

End Sub

Private Sub hg_Click()
Form1.WindowState = 2
Form1.Show
End Sub

Private Sub ll_Click()
Unload Me
End Sub

Private Sub nn_Click()
Form11.WindowState = 2
Form11.Show
End Sub

Private Sub qg_Click()
Form7.WindowState = 2
Form7.Show
End Sub

Private Sub combo1_LostFocus()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
a = Val(Combo1.Text)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN\PROJ1.accdb"
rs.Open "SELECT * FROM checkin WHERE cusid= " & a, conn
If rs.EOF Then

Combo1.Text = ""
Else
Text4.Text = rs.Fields("roomid").Value
b = rs.Fields("roomid").Value
Text11.Text = rs.Fields("dofarrival").Value
Text12.Text = rs.Fields("checkoutdate").Value
End If
Dim rs1 As New ADODB.Recordset
a = Val(Combo1.Text)
rs1.Open "SELECT * FROM addguest WHERE CusId= " & a, conn
If rs1.EOF Then
Else
c = rs1.Fields("fname").Value & " " & rs1.Fields("lname").Value
Text2.Text = c
Text3.Text = rs1.Fields("adult").Value + rs1.Fields("child").Value
End If
Dim rs2 As New ADODB.Recordset
rs2.Open "SELECT * FROM room WHERE roomid= " & b, conn
If rs2.EOF Then
MsgBox ("Data not found")
Else
Text6.Text = rs2.Fields("roomrate").Value * rs.Fields("nofdayres").Value
If rs2.Fields("roomrate").Value > 10000 Then
g = rs2.Fields("roomrate").Value
 Text8.Text = (g * Val(Form10.Text8.Text)) / 100
 Else
 If rs2.Fields("roomrate").Value > 7000 Then
g = rs2.Fields("roomrate").Value
 Text8.Text = (g * Val(Form10.Text9.Text)) / 100
 Else
 If rs2.Fields("roomrate").Value > 2000 Then
g = rs2.Fields("roomrate").Value
 Text8.Text = (g * Val(Form10.Text10.Text)) / 100
End If
End If
End If
End If
Dim rs3 As New ADODB.Recordset
rs3.Open "SELECT * FROM extraser WHERE cusid= " & a, conn
Text7.Text = 0
If rs3.EOF Then
Else
If rs3.Fields("swim").Value = True Then
d = Form10.Text1.Text
Text7.Text = 0 + Val(d)
End If
If rs3.Fields("laundry").Value = True Then
d = Form10.Text2.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
If rs3.Fields("spa").Value = True Then
d = Form10.Text4.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
If rs3.Fields("gym").Value = True Then
d = Form10.Text3.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
If rs3.Fields("roomservice").Value = True Then
d = Form10.Text5.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
If rs3.Fields("childcare").Value = True Then
d = Form10.Text6.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
If rs3.Fields("valetparking").Value = True Then
d = Form10.Text7.Text
Text7.Text = Val(Text7.Text) + Val(d)
End If
End If
rs.Close
rs1.Close
rs2.Close
rs3.Close
conn.Close
Set rs3 = Nothing
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set conn = Nothing
Dim tot As Double
Dim dis As Double
tot = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
If tot > 30000 Then

dis = (tot * Val(Form10.Text11.Text)) / 100
Text9.Text = dis
Text10.Text = tot - dis
Else
If tot > 20000 Then

dis = (tot * Val(Form10.Text12.Text)) / 100
Text9.Text = dis
Text10.Text = tot - dis
Else
 If tot > 10000 Then
 dis = (tot * Val(Form10.Text13.Text)) / 100
Text9.Text = dis
 Text10.Text = tot - dis
 End If
 End If
End If

End Sub

Private Sub vc_Click()
Form5.WindowState = 2
Form5.Show
End Sub

Private Sub vy_Click()
Form10.WindowState = 2
Form10.Show
End Sub
