VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00008000&
   Caption         =   "Laporan Data Kunujungan Pasien Per Hari"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form8"
   ScaleHeight     =   5250
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Operator 
      BackColor       =   &H00008000&
      Caption         =   "Operator"
      ForeColor       =   &H8000000B&
      Height          =   2655
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Grafik"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Preview"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   107347971
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
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   107347971
         CurrentDate     =   43603
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Sampai Tanggal"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Dari Tanggal"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   120
      Picture         =   "Form8.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Menu ini digunakan untuk melihat data kunjungan pasien perhari"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   840
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Laporan Data Kunujungan Pasien Per Hari"
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
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tglawal, tglakhir
tglawal = Format(DTtgldari.Value, "dd-mm-yyyy")
tglakhir = Format(DTtglsampai.Value, "dd-mm-yyyy")
  With CrystalReport1
        .DiscardSavedData = True
        .ReportFileName = App.Path & "\ReportGrafikKunjungan.rpt"
      .SelectionFormula = "{data_pemeriksaan.Tgl_pemeriksaan}>=date(" & Year(DTtgldari) & "," & Month(DTtgldari) & "," & Day(DTtgldari) & ") and {data_pemeriksaan.Tgl_pemeriksaan}<=date(" & Year(DTtglsampai) & "," & Month(DTtglsampai) & "," & Day(DTtglsampai) & ")"
        
 
           .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        End With
End Sub

Private Sub Command3_Click()
Unload Form8
End Sub

Private Sub Command7_Click()
Dim tglawal, tglakhir
tglawal = Format(DTtgldari.Value, "dd-mm-yyyy")
tglakhir = Format(DTtglsampai.Value, "dd-mm-yyyy")
  With CrystalReport1
        .DiscardSavedData = True
        .ReportFileName = App.Path & "\Reportdatakunjungan.rpt"
      .SelectionFormula = "{data_pemeriksaan.Tgl_pemeriksaan}>=date(" & Year(DTtgldari) & "," & Month(DTtgldari) & "," & Day(DTtgldari) & ") and {data_pemeriksaan.Tgl_pemeriksaan}<=date(" & Year(DTtglsampai) & "," & Month(DTtglsampai) & "," & Day(DTtglsampai) & ")"
        
 
           .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        End With
End Sub

