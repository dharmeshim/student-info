VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "login"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc loginado 
      Height          =   375
      Left            =   13320
      Top             =   9000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\OOAD\project final\studentinfo.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\OOAD\project final\studentinfo.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from staffinfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3720
      TabIndex        =   10
      Top             =   4800
      Width           =   2600
   End
   Begin VB.TextBox dtxt 
      Height          =   600
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "MM/DD/YYYY"
      Top             =   3120
      Width           =   2600
   End
   Begin VB.TextBox sitxt 
      Height          =   600
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Enter your Student ID"
      Top             =   2040
      Width           =   2600
   End
   Begin VB.TextBox txtpass 
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   10560
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter password"
      Top             =   3120
      Width           =   2600
   End
   Begin VB.TextBox txtuser 
      Height          =   600
      Left            =   10560
      TabIndex        =   1
      ToolTipText     =   "Enter your staff ID"
      Top             =   2050
      Width           =   2600
   End
   Begin VB.CommandButton loginbtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10560
      TabIndex        =   0
      Top             =   4800
      Width           =   2600
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "About Us"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "HIDE"
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "SHOW"
      Height          =   255
      Left            =   10560
      TabIndex        =   16
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "  <"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot password?"
      Height          =   255
      Left            =   11880
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SIGN UP"
      Height          =   375
      Left            =   12600
      TabIndex        =   13
      Top             =   5640
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   8880
      X2              =   11640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   4920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   7920
      Y1              =   600
      Y2              =   6720
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "STAFF LOGIN"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8880
      TabIndex        =   12
      Top             =   600
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT LOGIN"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   3795
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't have a login id?"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   2160
      Width           =   885
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
login.Hide
End Sub


Private Sub Command4_Click()
loginado.RecordSource = "select * from studentinfo where studentid='" + sitxt.Text + "' and dob='" + dtxt.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Login failled..!", vbCritical, "Please Enter the correct userid and password"
Else
MsgBox "Login Successful.", vbExclamation, "Successful attempt"
student.Show
txtuser.Text = ""
txtpass.Text = ""
End If
End Sub

Private Sub Label10_Click()
login.Hide
End Sub

Private Sub Label11_Click()
txtpass.PasswordChar = ""
Label11.Visible = False
Label12.Visible = True
End Sub

Private Sub Label12_Click()
txtpass.PasswordChar = "*"
Label11.Visible = True
Label12.Visible = False
End Sub

Private Sub Label13_Click()
about.Show
End Sub

Private Sub Label8_Click()
signup.Show
End Sub

Private Sub Label9_Click()
forgot.Show
End Sub

Private Sub loginbtn_Click()
loginado.RecordSource = "select * from staffinfo where staffid='" + txtuser.Text + "' and password='" + txtpass.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Login failled..!", vbCritical, "Please Enter the correct userid and password"
Else
MsgBox "Login Successful.", vbExclamation, "Successful attempt"
staff.Show
txtuser.Text = ""
txtpass.Text = ""
End If
End Sub

Private Sub Text3_Change()

End Sub

