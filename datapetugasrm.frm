VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00008000&
   Caption         =   "Data Petugas"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   LinkTopic       =   "Form12"
   ScaleHeight     =   7560
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Height          =   525
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000007&
      Caption         =   "Cari"
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Operasi 
      BackColor       =   &H00008000&
      Caption         =   "Operasi"
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
      Begin VB.CommandButton Command6 
         Caption         =   "Tambah"
         Height          =   615
         Left            =   720
         TabIndex        =   17
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Keluar"
         Height          =   615
         Left            =   720
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Simpan"
         Height          =   615
         Left            =   720
         MaskColor       =   &H008080FF&
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   615
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Inputan Data Petugas"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
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
         TabIndex        =   16
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Kode Petugas"
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
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "datapetugasrm.frx":0000
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   4800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Klinik1;Initial Catalog=klinik"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Klinik1;Initial Catalog=klinik"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "data_diagnosa"
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
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
      Caption         =   "Data Petugas "
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
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Menu ini digunakan untuk memasukan data petugas RM"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   720
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   120
      Picture         =   "datapetugasrm.frx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "Form12"
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
Call BukaDB
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim TambahAdmin As String
    TambahAdmin = "Insert Into tb_admin values ('" & Text1 & "','" & Text2 & "','" & Text3 & "')"
    Koneksi.Execute TambahAdmin
    MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
    Form_Load
End If

End Sub

Private Sub Command2_Click()
Call BukaDB
Dim HapusAdmin As String
    HapusAdmin = "Delete From tb_admin where kode_admin='" & Text1 & "'"
    Koneksi.Execute HapusAdmin
    MsgBox "Data Berhasil DiHapus", vbInformation, "Pemberitahuan"
    Form_Load
End Sub

Private Sub Command3_Click()
Unload Form12
End Sub

Private Sub Command4_Click()
Call BukaDB
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim EditAdmin As String
    EditAdmin = "update tb_admin Set username= '" & Text2 & "', password= '" & Text3 & "' where kode_admin='" & Text1 & "'"
    Koneksi.Execute EditAdmin
    MsgBox "Data Berhasil DiUpdate", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command6_Click()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Private Sub Command8_Click()
Call BukaDB

RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from tb_admin where username like '%" & Text9 & "%'", Koneksi

If Not RSBarang.EOF Then
    With RSBarang
        With DataGrid1
            Set .DataSource = RSBarang
                .Refresh
        End With
    End With
    End If
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    Call BukaDB
    RSBarang.Open "Select * from tb_admin where kode_admin = '" & DataGrid1.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
        Text1 = RSBarang!kode_admin
        Text2 = RSBarang!UserName
        Text3 = RSBarang!Password
       
       
        Command1.Enabled = True
        Else
        MsgBox "Data Tidak Ada!"
    End If
End Sub

Private Sub Form_Load()
Call BukaDB
Text1 = ""
Text2 = ""
Text3 = ""




Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "tb_admin"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

End Sub



