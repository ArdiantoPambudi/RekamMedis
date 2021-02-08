VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15900
   LinkTopic       =   "Form7"
   ScaleHeight     =   8940
   ScaleWidth      =   15900
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7200
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   13440
      TabIndex        =   32
      Top             =   2040
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTtglsampai 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   31
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   125108227
      CurrentDate     =   43603
   End
   Begin MSComCtl2.DTPicker DTtgldari 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   30
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   125108227
      CurrentDate     =   43603
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cari"
      Height          =   495
      Left            =   11400
      TabIndex        =   28
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   525
      Left            =   11280
      TabIndex        =   27
      Top             =   4320
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   2655
      Left            =   720
      TabIndex        =   26
      Top             =   6600
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4683
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
      Left            =   11040
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "data_pemeriksaan"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Output"
      Height          =   1935
      Left            =   9240
      TabIndex        =   24
      Top             =   4320
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
   Begin VB.Frame Operasi 
      BackColor       =   &H0080C0FF&
      Caption         =   "Operasi"
      Height          =   3495
      Left            =   9240
      TabIndex        =   19
      Top             =   720
      Width           =   2415
      Begin VB.CommandButton Command3 
         Caption         =   "Keluar"
         Height          =   615
         Left            =   720
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Simpan"
         Height          =   615
         Left            =   720
         MaskColor       =   &H008080FF&
         TabIndex        =   22
         Top             =   360
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
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   615
         Left            =   720
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Inputan Data Pemeriksaan"
      Height          =   5655
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   7935
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125108225
         CurrentDate     =   43600
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   4200
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Pasien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Dokter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Diagnosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Tindakan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Obat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kode Pemeriksaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tgl Pemeriksaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form7"
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
Sub Pasien()
Call BukaDB
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo1 & "%'", Koneksi
Combo1.Clear
Do While Not RSBarang.EOF
Combo1.AddItem RSBarang!No_RM
RSBarang.MoveNext

Loop
End Sub
Sub Dokter()
Call BukaDB
RSBarang.Open "Select * from data_dokter where Kode_dokter like '%" & Combo2 & "%'", Koneksi
Combo2.Clear
Do While Not RSBarang.EOF
Combo2.AddItem RSBarang!Kode_dokter
RSBarang.MoveNext
Loop
End Sub
Sub Diagnosa()
Call BukaDB
RSBarang.Open "Select * from data_diagnosa where Kode_diagnosa like '%" & Combo3 & "%'", Koneksi
Combo3.Clear
Do While Not RSBarang.EOF
Combo3.AddItem RSBarang!Kode_diagnosa
RSBarang.MoveNext
Loop
End Sub
Sub Tindakan()
Call BukaDB
RSBarang.Open "Select * from data_Tindakan where Kode_tindakan like '%" & Combo4 & "%'", Koneksi
Combo4.Clear
Do While Not RSBarang.EOF
Combo4.AddItem RSBarang!Kode_tindakan
RSBarang.MoveNext
Loop
End Sub
Sub Obat()
Call BukaDB
RSBarang.Open "Select * from data_obat where Kode_obat like '%" & Combo5 & "%'", Koneksi
Combo5.Clear
Do While Not RSBarang.EOF
Combo5.AddItem RSBarang!Kode_obat
RSBarang.MoveNext
Loop

End Sub



Private Sub Combo1_Click()
Call BukaDB
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo1 & "%'", Koneksi
If Not RSBarang.EOF Then
Text3.Text = RSBarang!Nama
End If
End Sub



Private Sub Combo2_Click()
Call BukaDB


RSBarang.Open "Select * from data_dokter where Kode_dokter like '%" & Combo2 & "%'", Koneksi
If Not RSBarang.EOF Then
Text4.Text = RSBarang!Nama_dokter
End If
End Sub


Private Sub Combo3_Click()
Call BukaDB
RSBarang.Open "Select * from data_diagnosa where Kode_diagnosa like '%" & Combo3 & "%'", Koneksi
If Not RSBarang.EOF Then
Text5.Text = RSBarang!Nama_diagnosa
End If
End Sub
Private Sub Combo4_Click()
Call BukaDB
RSBarang.Open "Select * from data_tindakan where Kode_tindakan like '%" & Combo4 & "%'", Koneksi
If Not RSBarang.EOF Then
Text6.Text = RSBarang!Nama_tindakan
End If
End Sub


Private Sub Combo5_Click()
Call BukaDB
RSBarang.Open "Select * from data_obat where Kode_obat like '%" & Combo5 & "%'", Koneksi
If Not RSBarang.EOF Then
Text7.Text = RSBarang!Nama_obat
End If
End Sub

Private Sub Command1_Click()
Call BukaDB
If Text2 = "" Or DTPicker1 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim TambahPemeriksaan As String
    TambahPemeriksaan = "Insert Into data_pemeriksaan values ('" & Text2 & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "', '" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & Combo4 & "','" & Combo5 & "')"
    Koneksi.Execute TambahPemeriksaan
    MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
    Form_Load
    End If
End Sub

Private Sub Command2_Click()
Call BukaDB
Dim HapusPemeriksaan As String
    HapusPemeriksaan = "Delete From data_pemeriksaan where Kode_pemeriksaan='" & Text2 & "'"
    Koneksi.Execute HapusPemeriksaan
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
If Text1 = "" Or DTPicker1 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Then
MsgBox "Data Belum Lengkap"
Else
Dim EditAnggota As String
    EditAnggota = "update data_pemeriksaan Set Tgl_pemeriksaan= '" & DTPicker1 & "',Kode_pasien='" & Combo1 & "',Kode_dokter='" & Combo2 & "',Kode_diagnosa='" & Combo3 & "',Kode_tindakan='" & Combo4 & "',Kode_obat='" & Combo5 & "' where Kode_pemeriksaan='" & Text1 & "'"
    Koneksi.Execute EditAnggota
    MsgBox "Data Berhasil DiUpdate", vbInformation, "Pemberitahuan"
    Form_Load
End If
End Sub

Private Sub Command5_Click()
RSBarang.CursorLocation = adUseClient
RSBarang.Open "Select * from data_pemeriksaan where Kode_pemeriksaan like '%" & Text9 & "%'", Koneksi

If Not RSBarang.EOF Then
    With RSBarang
        With DataGrid1
            Set .DataSource = RSBarang
                .Refresh
        End With
    End With
End If
End Sub
Sub coba()
  DataEnvironment6.Tanggal
   
   DataReport6.Title = "Dari : " & DTtgldari & DTtglsampai
   Me.Hide
   DataReport6.Show
   Unload Me
   
End Sub

Private Sub Command7_Click()

Dim tglawal, tglakhir
tglawal = Format(DTtgldari.Value, "dd-mm-yyyy")
tglakhir = Format(DTtglsampai.Value, "dd-mm-yyyy")
  With CrystalReport1
        .DiscardSavedData = True
        .ReportFileName = App.Path & "\coba.rpt"
      .SelectionFormula = "{data_pemeriksaan.Tgl_pemeriksaan}>=date(" & Year(DTtgldari) & "," & Month(DTtgldari) & "," & Day(DTtgldari) & ") and {data_pemeriksaan.Tgl_pemeriksaan}<=date(" & Year(DTtglsampai) & "," & Month(DTtglsampai) & "," & Day(DTtglsampai) & ")"
        
 
           .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        End With
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    Call BukaDB
    RSBarang.Open "Select * from data_pemeriksaan where Kode_pemeriksaan = '" & DataGrid1.Columns(0) & "'", Koneksi
    If Not RSBarang.EOF Then
        Text2 = RSBarang!Kode_pemeriksaan
        DTPicker1 = RSBarang!Tgl_pemeriksaan
        Text3 = RSBarang!Kode_pasien
        Text4 = RSBarang!Kode_dokter
        Text5 = RSBarang!Kode_diagnosa
        Text6 = RSBarang!Kode_tindakan
        Text7 = RSBarang!Kode_obat
       
        Command1.Enabled = True
         Form_Load
        Else
        MsgBox "Data Tidak Ada!"
    End If
End Sub
Sub Tampillaporan()
If RSBarang.BOF Then
   MsgBox "Data Tidak Tersedia.", vbInformation + vbOKOnly, "informasi"
Else
   With DataReport6
   Set .DataSource = RSBarang
   .DataMember = ""
    Set .DataSource = RSBarang
      
    .Sections("Section4").Controls("Label7"). _
    Caption = Format(DTtgldari.Value, "dd-mm-YYYY")
     .Sections("Section4").Controls("Label7"). _
    Caption = Format(DTtglsampai, "dd-mm-YYYY")
    .LeftMargin = 100
    .RightMargin = 100
    .WindowState = 2
   End With
End If
End Sub

Private Sub Form_Load()
Call Pasien
Call Tindakan
Call Obat
Call Diagnosa
Call Dokter
Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "data_pemeriksaan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Text1 = ""
Me.DTPicker1.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo5 = ""
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo1 & "%'", Koneksi
Combo1.Clear
Do While Not RSBarang.EOF
Combo1.AddItem RSBarang!No_RM
RSBarang.MoveNext
Loop
RSBarang.Close



RSBarang.Open "Select * from data_dokter where Kode_dokter like '%" & Combo2 & "%'", Koneksi

Combo2.Clear
Do While Not RSBarang.EOF
Combo2.AddItem RSBarang!Kode_dokter
RSBarang.MoveNext
Loop
RSBarang.Close

RSBarang.Open "Select * from data_diagnosa where Kode_diagnosa like '%" & Combo3 & "%'", Koneksi
Combo3.Clear
Do While Not RSBarang.EOF
Combo3.AddItem RSBarang!Kode_diagnosa
RSBarang.MoveNext
Loop
RSBarang.Close


RSBarang.Open "Select * from data_Tindakan where Kode_tindakan like '%" & Combo4 & "%'", Koneksi
Combo4.Clear
Do While Not RSBarang.EOF
Combo4.AddItem RSBarang!Kode_tindakan
RSBarang.MoveNext
Loop
RSBarang.Close

RSBarang.Open "Select * from data_obat where Kode_obat like '%" & Combo5 & "%'", Koneksi
Combo5.Clear
Do While Not RSBarang.EOF
Combo5.AddItem RSBarang!Kode_obat
RSBarang.MoveNext
Loop
RSBarang.Close



End Sub




