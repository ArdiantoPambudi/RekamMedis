VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00008000&
   Caption         =   "Entry Data Pendaftaran"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14340
   LinkTopic       =   "Form13"
   ScaleHeight     =   8280
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Search Data Pendaftaran"
      ForeColor       =   &H8000000B&
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   6015
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000007&
         Caption         =   "Cari"
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "pendaftaran.frx":0000
         Height          =   1935
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
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
         Height          =   375
         Left            =   4440
         Top             =   240
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
         RecordSource    =   "tb_pendaftaran"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Search Data Pasien"
      ForeColor       =   &H8000000B&
      Height          =   3375
      Left            =   8760
      TabIndex        =   17
      Top             =   1560
      Width           =   5295
      Begin VB.TextBox Text9 
         Height          =   525
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H80000007&
         Caption         =   "Cari"
         Height          =   495
         Left            =   2640
         MaskColor       =   &H00FFFFC0&
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   450
         Left            =   3840
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   794
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
         ForeColor       =   0
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Klinik"
         OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Klinik"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tb_pasien"
         Caption         =   "Adodc2"
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
   End
   Begin VB.Frame Operasi 
      BackColor       =   &H00008000&
      Caption         =   "Operasi"
      ForeColor       =   &H8000000B&
      Height          =   4095
      Left            =   6240
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
      Begin VB.CommandButton Command6 
         Caption         =   "Tambah"
         Height          =   615
         Left            =   720
         TabIndex        =   14
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   720
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         Height          =   615
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Simpan"
         Height          =   615
         Left            =   720
         MaskColor       =   &H008080FF&
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Keluar"
         Height          =   615
         Left            =   720
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Inputan Data Pendaftaran"
      ForeColor       =   &H8000000B&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox DTPicker1 
         Height          =   315
         ItemData        =   "pendaftaran.frx":0015
         Left            =   2160
         List            =   "pendaftaran.frx":002B
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Poli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "No Registrasi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "No RM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   14040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Menu ini digunakan untuk memasukan data pendaftaran"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   840
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   600
      Picture         =   "pendaftaran.frx":008E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "Form13"
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
Private Sub NomorOtomatis()
Call BukaDB
On Error Resume Next
RSBarang.Open ("select * from tb_pendaftaran Where No_registrasi In(Select Max(No_registrasi)From tb_pendaftaran)Order By No_registrasi Desc"), Koneksi
RSBarang.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RSBarang
        If .EOF Then
            Urutan = "AGT" + "001"
            Text1 = Urutan
        Else
            Hitung = Right(RSBarang!No_registrasi, 3) + 1
            Urutan = "AGT" + Right("000" & Hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub

Private Sub Combo1_Click()
Call BukaDB
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo1 & "%'", Koneksi
If Not RSBarang.EOF Then
Text2.Text = RSBarang!Nama
End If
End Sub

Private Sub Command1_Click()
Call BukaDB
If Text1 = "" Or Combo1 = "" Or Text2 = "" Or DTPicker1 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim TambahPasien As String
    TambahPasien = "Insert Into tb_pendaftaran values ('" & Text1 & "','" & Combo1 & "','" & Text2 & "','" & DTPicker1 & "')"
    Koneksi.Execute TambahPasien
    MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub
Private Sub Command2_Click()
Call BukaDB
Dim HapusPasien As String
    HapusPasien = "Delete From tb_pendaftaran where No_Registrasi='" & Text1 & "'"
    Koneksi.Execute HapusPasien
    MsgBox "Data Berhasil DiHapus", vbInformation, "Pemberitahuan"
    Form_Load
End Sub

Private Sub Command3_Click()
Unload Form13
End Sub

Private Sub Command4_Click()
Call BukaDB
If Text1 = "" Or Text2 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim EditObat As String
    EditObat = "update tb_pendaftaran Set No_RM= '" & Combo1 & "',Nama= '" & Text2 & "',Tanggal_periksa= '" & DTPicker1 & "' where No_Registrasi='" & Text1 & "'"
    Koneksi.Execute EditObat
    MsgBox "Data Berhasil DiUpdate", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command5_Click()
Call BukaDB

RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from tb_pendaftaran where No_RM like '%" & Text3 & "%'", Koneksi

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

Text2 = ""

Combo1.Clear
End Sub

Private Sub Command8_Click()
Call BukaDB

RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Text9 & "%'", Koneksi

If Not RSBarang.EOF Then
    With RSBarang
        With DataGrid2
            Set .DataSource = RSBarang
                .Refresh
        End With
    End With
    End If
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    Call BukaDB
    RSBarang.Open "Select * from tb_pendaftaran where No_Registrasi = '" & DataGrid1.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
        Text1 = RSBarang!No_registrasi
        Text2 = RSBarang!Nama
        Combo1 = RSBarang!No_RM
        DTPicker1 = RSBarang!Tanggal_periksa
       
       
        Command1.Enabled = True
        Else
        MsgBox "Data Tidak Ada!"
    End If
End Sub

Private Sub DataGrid2_Click()
On Error Resume Next
    Call BukaDB
    RSBarang.Open "Select * from tb_pasien where No_RM = '" & DataGrid2.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
       
        Text2 = RSBarang!Nama
        Combo1 = RSBarang!No_RM
        DTPicker1 = RSBarang!Tanggal_periksa
       
       
        Command1.Enabled = True
        Else
        MsgBox "Data Tidak Ada!"
    End If
End Sub

Private Sub Form_Load()
Call NomorOtomatis

BukaDB


Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "tb_pendaftaran"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Adodc2.ConnectionString = Koneksi
Adodc2.RecordSource = "tb_pasien"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2

RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo1 & "%'", Koneksi
Combo1.Clear
Do While Not RSBarang.EOF
Combo1.AddItem RSBarang!No_RM
RSBarang.MoveNext
Loop

End Sub


