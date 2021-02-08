VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Output"
      Height          =   1935
      Left            =   11520
      TabIndex        =   24
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton Command6 
         Caption         =   "Tampil"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Keluar"
      Height          =   615
      Left            =   9600
      TabIndex        =   23
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Operasi 
      BackColor       =   &H0080C0FF&
      Caption         =   "Operasi"
      Height          =   3495
      Left            =   8880
      TabIndex        =   19
      Top             =   960
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   720
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   615
         Left            =   720
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Simpan"
         Height          =   615
         Left            =   720
         MaskColor       =   &H008080FF&
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Inputan Data Pasien"
      Height          =   4695
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   7815
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2040
         TabIndex        =   8
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2040
         TabIndex        =   7
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2040
         TabIndex        =   6
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   3720
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   2040
         List            =   "Form1.frx":000A
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "No RM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H008080FF&
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "TTL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
         Caption         =   "JenKel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Caption         =   "Pekerjaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H008080FF&
         Caption         =   "No Telp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cari"
      Height          =   495
      Left            =   11400
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   525
      Left            =   8880
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8880
      Top             =   5160
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
      RecordSource    =   "tb_pasien"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0024
      Height          =   3015
      Left            =   840
      TabIndex        =   0
      Top             =   5760
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5318
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
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "Data Pendaftaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim TambahAnggota As String
    TambahAnggota = "Insert Into tb_pasien values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "')"
    Koneksi.Execute TambahAnggota
    MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
    Form_Load
End If


End Sub

Private Sub Command2_Click()
Call BukaDB
Dim HapusAnggota As String
    HapusAnggota = "Delete From tb_pasien where No_RM='" & Text1 & "'"
    Koneksi.Execute HapusAnggota
    MsgBox "Data Berhasil DiHapus", vbInformation, "Pemberitahuan"
    Form_Load
End Sub

Private Sub Command3_Click()
X = MsgBox("yakin keluar?", vbQuestion + vbYesNo, "informasi")
If X = vbYes Then
End
End If
End Sub

Private Sub Command4_Click()
Call BukaDB
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim EditAnggota As String
    EditAnggota = "update tb_pasien Set Nama= '" & Text2 & "',TTL='" & Text3 & "',Jenis_kelamin='" & Combo1 & "',Alamat='" & Text5 & "',Pekerjaan='" & Text6 & "',No_hp='" & Text7 & "' where No_RM='" & Text1 & "'"
    Koneksi.Execute EditAnggota
    MsgBox "Data Berhasil DiUpdate", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command5_Click()
Call BukaDB


RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from tb_pasien where Nama like '%" & Text9 & "%'", Koneksi

If Not RSBarang.EOF Then
    With RSBarang
        With DataGrid1
            Set .DataSource = RSBarang
                .Refresh
        End With
    End With
End If
   
      
End Sub

Private Sub Command6_Click()
DataReport1.Show
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    Call BukaDB
    RSBarang.Open "Select * from tb_pasien where No_RM = '" & DataGrid1.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
        Text1 = RSBarang!No_RM
        Text2 = RSBarang!Nama
        Text3 = RSBarang!TTL
        Combo1 = RSBarang!Jenis_kelamin
        Text5 = RSBarang!Alamat
        Text6 = RSBarang!Pekerjaan
        Text7 = RSBarang!No_hp
       
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
Combo1 = ""
Text5 = ""
Text6 = ""
Text7 = ""



Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "tb_pasien"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1




End Sub

Private Sub Text1_Change()
Call BukaDB
If Text1.MaxLength <= 6 Then
MsgBox "Max 6 Karakter"
Else

End If
End Sub

