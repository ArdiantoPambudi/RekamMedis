VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00008000&
   Caption         =   "Laporan Data Kunjungan Pasien Per Bulan"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   LinkTopic       =   "Form10"
   ScaleHeight     =   4680
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Operator 
      BackColor       =   &H00008000&
      Caption         =   "Operator"
      ForeColor       =   &H8000000B&
      Height          =   2655
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Preview"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker bln 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MMMM-yyyy"
         Format          =   106561539
         CurrentDate     =   43637
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Nama Bulan"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3960
      Top             =   2280
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
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   120
      Picture         =   "Form10GrafikkunjunganBulan.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Menu ini digunakan untuk melihat grafik data kunjungan pasien per bulan"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Grafik Data Kunjungan Pasien Per Bulan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Form10
End Sub

Private Sub Command8_Click()
With CrystalReport1
.DiscardSavedData = True
        .ReportFileName = App.Path & "\ReportGrafikKunjungan.rpt"
.SelectionFormula = _
    " YEAR({data_pemeriksaan.Tgl_pemeriksaan})= " & Year(bln) & _
    " and month({data_pemeriksaan.Tgl_pemeriksaan})= " & Month(bln) & ""
    .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
End With


End Sub

