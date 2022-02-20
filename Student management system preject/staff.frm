VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form staff 
   Caption         =   "staff"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton lastbtn 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11400
      TabIndex        =   37
      Top             =   9480
      Width           =   1000
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2760
      TabIndex        =   36
      Top             =   9480
      Width           =   1000
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10200
      TabIndex        =   35
      Top             =   9480
      Width           =   1000
   End
   Begin VB.CommandButton previousbtn 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3960
      TabIndex        =   34
      Top             =   9480
      Width           =   1000
   End
   Begin VB.TextBox Text10 
      Height          =   500
      Left            =   12000
      TabIndex        =   32
      Top             =   7320
      Width           =   1500
   End
   Begin VB.TextBox Text9 
      Height          =   500
      Left            =   12000
      TabIndex        =   27
      Top             =   6360
      Width           =   1500
   End
   Begin VB.TextBox Text8 
      Height          =   500
      Left            =   12000
      TabIndex        =   26
      Top             =   5400
      Width           =   1500
   End
   Begin VB.TextBox Text7 
      Height          =   500
      Left            =   12000
      TabIndex        =   25
      Top             =   4440
      Width           =   1500
   End
   Begin VB.TextBox Text6 
      Height          =   500
      Left            =   12000
      TabIndex        =   24
      Top             =   3480
      Width           =   1500
   End
   Begin VB.TextBox Text5 
      Height          =   500
      Left            =   12000
      TabIndex        =   23
      Top             =   2520
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   17760
      Top             =   10080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      RecordSource    =   "select * from studentinfo"
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
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8520
      TabIndex        =   20
      Top             =   9480
      Width           =   1500
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6840
      TabIndex        =   19
      Top             =   9480
      Width           =   1500
   End
   Begin VB.CommandButton addnew 
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5160
      TabIndex        =   18
      Top             =   9480
      Width           =   1500
   End
   Begin VB.TextBox Text4 
      Height          =   500
      Left            =   3600
      TabIndex        =   17
      Top             =   7800
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      Height          =   500
      Left            =   3600
      TabIndex        =   8
      Top             =   6960
      Width           =   2500
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Text            =   "Select Section"
      Top             =   6240
      Width           =   2500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Text            =   "Select Grade"
      Top             =   5520
      Width           =   2500
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   3960
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   873
      _Version        =   393216
      Format          =   123142145
      CurrentDate     =   44493
   End
   Begin VB.TextBox Text2 
      Height          =   500
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   2500
   End
   Begin VB.Label Label18 
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
      Left            =   120
      TabIndex        =   39
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   7800
      Y1              =   1440
      Y2              =   8520
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT INFORMATIONS"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   38
      Top             =   1400
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   3000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Total "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   33
      Top             =   7440
      Width           =   840
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Environmental Science"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   31
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Science"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   30
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Mathematics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   29
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Language 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   28
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Language 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   22
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "MARKS"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   21
      Top             =   1400
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   7080
      Width           =   1500
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   6240
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Class of Studying"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   5520
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ds As Database
Public rs As Recordset
Dim str As String
Dim confirm As Integer


Private Sub addnew_Click()
rs.addnew
clear
End Sub

Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to delete?", vbYesNo, "Delete")
If confirm = vbYes Then
rs.Delete
MsgBox "Record Deleted", vbInformation, "Delete"
'rs.Update
refreshdata
Else
MsgBox "Can't able to Delete!", vbInformation, "Delete"
End If
End Sub
Sub refreshdata()
rs.Close
Set rs = ds.OpenRecordset("select * from studentinfo")
'rs.OpenRecordset "select * from studentinfo", ds, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No record found"
End If
End Sub

Private Sub firstbtn_Click()
rs.MoveFirst
display
End Sub

Private Sub Form_Load()
Set ds = OpenDatabase("E:\OOAD\project final\studentinfo.mdb")
Set rs = ds.OpenRecordset("select * from studentinfo")

Combo1.AddItem "5th Grade"
Combo1.AddItem "6th Grade"
Combo1.AddItem "7th Grade"
Combo1.AddItem "8th Grade"
Combo1.AddItem "9th Grade"
Combo1.AddItem "10th Grade"

Combo2.AddItem "Section A"
Combo2.AddItem "Section B"
Combo2.AddItem "Section C"
Combo2.AddItem "Section D"

display

End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = "01/01/2002"
Option1.Value = False
Option2.Value = False
Combo1.Text = "Select Class of studying"
Combo2.Text = "Select Section"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""

End Sub
Sub display()
Text1.Text = rs!studentname
Text2.Text = rs!studentid
DTPicker1.Value = rs!dob
If rs!gender = "Male" Then
Option1.Value = True
Else
Option2.Value = True
End If
Combo1.Text = rs!classstudy
Combo2.Text = rs!sec
Text3.Text = rs!phoneno
Text4.Text = rs!Address
Text5.Text = rs!lan1
Text6.Text = rs!lan2
Text7.Text = rs!maths
Text8.Text = rs!science
Text9.Text = rs!environt
Text10.Text = rs!total

End Sub

Private Sub Label18_Click()
staff.Hide
End Sub

Private Sub lastbtn_Click()
rs.MoveLast
display
End Sub

Private Sub nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If
End Sub

Private Sub previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub savebtn_Click()
'rs.Edit
rs.Fields("studentname").Value = Text1.Text
rs.Fields("studentid").Value = Text2.Text
rs.Fields("dob").Value = DTPicker1.Value
If Option1.Value = True Then
rs.Fields("gender") = Option1.Caption
Else
rs.Fields("gender") = Option2.Caption
End If
rs.Fields("classstudy").Value = Combo1.Text
rs.Fields("sec").Value = Combo2.Text
rs.Fields("phoneno").Value = Text3.Text
rs.Fields("address").Value = Text4.Text
rs.Fields("lan1").Value = Text5.Text
rs.Fields("lan2").Value = Text6.Text
rs.Fields("maths").Value = Text7.Text
rs.Fields("science").Value = Text8.Text
rs.Fields("environt").Value = Text9.Text
rs.Fields("total").Value = Text10.Text
MsgBox "Data saved successfully..!", vbInformation, "saved"
rs.Update

End Sub

Private Sub Text10_GotFocus()
Text10.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)
End Sub

