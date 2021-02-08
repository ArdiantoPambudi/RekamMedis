VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form4"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14055
   LinkTopic       =   "Form4"
   ScaleHeight     =   8490
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000007&
      Caption         =   "Cari"
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   525
      Left            =   1320
      TabIndex        =   17
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Output"
      Height          =   1935
      Left            =   7920
      TabIndex        =   15
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton Command6 
         Caption         =   "Tampil"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Inputan Data Dokter"
      Height          =   3495
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   3855
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   1560
         TabIndex        =   14
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   1560
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "No HP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Nama Dokter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Dokter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Operasi 
      BackColor       =   &H0080C0FF&
      Caption         =   "Operasi"
      Height          =   3495
      Left            =   5400
      TabIndex        =   1
      Top             =   960
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Simpan"
         Height          =   615
         Left            =   720
         MaskColor       =   &H008080FF&
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Keluar"
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   2535
      Left            =   1320
      TabIndex        =   19
      Top             =   5760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
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
      Left            =   1320
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
      RecordSource    =   "data_dokter"
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Data Dokter"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim TambahDokter As String
    TambahDokter = "Insert Into data_dokter values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
    Koneksi.Execute TambahDokter
    MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command2_Click()
Call BukaDB
Dim HapusDokter As String
    HapusDokter = "Delete From data_dokter where Kode_dokter='" & Text1 & "'"
    Koneksi.Execute HapusDokter
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim EditDokter As String
    EditDokter = "update data_dokter Set Nama_dokter= '" & Text2 & "',Alamat='" & Text3 & "',No_hp='" & Text4 & "' where Kode_dokter='" & Text1 & "'"
    Koneksi.Execute EditDokter
    MsgBox "Data Berhasil DiUpdate", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command6_Click()
DataReport3.Show
End Sub

Private Sub Command8_Click()
Call BukaDB

RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from data_dokter where Nama_dokter like '%" & Text9 & "%'", Koneksi

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
    RSBarang.Open "Select * from data_dokter where Kode_dokter = '" & DataGrid1.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
        Text1 = RSBarang!Kode_dokter
        Text2 = RSBarang!Nama_dokter
        Text3 = RSBarang!Alamat
        Text4 = RSBarang!No_hp
       
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
Text4 = ""
Text9 = ""




Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "data_dokter"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

