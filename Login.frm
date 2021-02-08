VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00008000&
   Caption         =   "Login"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6780
   LinkTopic       =   "Form11"
   ScaleHeight     =   4620
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Silahkan Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
      Begin VB.TextBox txtpassword 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtusername 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "Sistem Informasi Pengelolaan Rekam Medis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   240
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Silahkan Login Terlebih dahulu jika ingin masuk pada "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Koneksi As New ADODB.Connection
Dim RSBarang As ADODB.Recordset
Sub BukaDB()
Set Koneksi = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Koneksi.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Klinik1;Initial Catalog=klinik"

End Sub

Private Sub Command1_Click()
BukaDB
Set RSBarang = New ADODB.Recordset
RSBarang.Open "select * from tb_admin where username='" & txtusername.Text & "'and password='" & txtpassword.Text & "'", Koneksi
If RSBarang.EOF Then
   MsgBox "Password dan User Salah !", vbExclamation, "Informasi"
   txtusername.SetFocus
   txtusername.Text = ""
   txtpassword.Text = ""
Else
   MsgBox "Login Sukses !", vbInformation, "Informasi "
   
   Form3.Show
 
End If

     

End Sub

Private Sub Text2_Change()

End Sub
