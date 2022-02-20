VERSION 5.00
Begin VB.Form signup 
   Caption         =   "Signup"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Sign up"
      Height          =   11175
      Left            =   -1200
      TabIndex        =   0
      Top             =   -240
      Width           =   20730
      Begin VB.CommandButton Command1 
         Caption         =   "SIGN UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6960
         TabIndex        =   9
         Top             =   5640
         Width           =   2500
      End
      Begin VB.TextBox ph 
         Height          =   500
         Left            =   6960
         TabIndex        =   8
         ToolTipText     =   "Enter your Phone number"
         Top             =   4320
         Width           =   2500
      End
      Begin VB.TextBox ps 
         Height          =   500
         Left            =   6960
         TabIndex        =   6
         ToolTipText     =   "Create password"
         Top             =   3480
         Width           =   2500
      End
      Begin VB.TextBox staffid 
         Height          =   500
         Left            =   6960
         TabIndex        =   4
         ToolTipText     =   "Enter your Staff ID"
         Top             =   2640
         Width           =   2500
      End
      Begin VB.TextBox fname 
         Height          =   500
         Left            =   6960
         TabIndex        =   2
         ToolTipText     =   "Enter your Fullname"
         Top             =   1800
         Width           =   2500
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   3960
         X2              =   6960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "STAFF SIGN UP"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Phone Number"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   4440
         Width           =   1620
      End
      Begin VB.Label Label3 
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
         Left            =   5520
         TabIndex        =   5
         Top             =   3600
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Id"
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
         Left            =   5760
         TabIndex        =   3
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   1920
         Width           =   1140
      End
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ds As Database
Public rs As Recordset


Private Sub Command1_Click()
rs.addnew
rs.Fields(0) = fname.Text
rs.Fields(3) = staffid.Text
rs.Fields(1) = ps.Text
rs.Fields(2) = ph.Text
rs.Update
rs.Requery
MsgBox "Sign up successfully", vbExclamation, "Done"
login.Show
signup.Hide

End Sub


Private Sub Form_Load()
Set ds = OpenDatabase("E:\OOAD\project final\studentinfo.mdb")
Set rs = ds.OpenRecordset("select * from staffinfo")
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Label6_Click()
signup.Hide
End Sub
