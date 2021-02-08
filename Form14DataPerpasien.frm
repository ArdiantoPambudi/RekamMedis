VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form14 
   BackColor       =   &H00008000&
   Caption         =   "Laporan Data Per Pasien"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6900
   FillColor       =   &H00008000&
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form14"
   ScaleHeight     =   5010
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Operator 
      BackColor       =   &H00008000&
      Caption         =   "Operator"
      ForeColor       =   &H8000000B&
      Height          =   2655
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Preview"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Nama Pasien"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "No RM"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   1920
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
      Left            =   3240
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
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Laporan Data Per Pasien"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   360
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Menu ini digunakan untuk melihat data Per Pasien"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   120
      Picture         =   "Form14DataPerpasien.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "Form14"
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












Private Sub Combo6_Click()
Call BukaDB
RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo6 & "%'", Koneksi
If Not RSBarang.EOF Then
Text5.Text = RSBarang!Nama
End If
End Sub


Private Sub Command7_Click()

End Sub

Private Sub Command3_Click()
Unload Form14
End Sub

Private Sub Command8_Click()
With CrystalReport1
        .DiscardSavedData = True
        .ReportFileName = App.Path & "\Reportdatapemeriksaan.rpt"
      .SelectionFormula = "{data_pemeriksaan.No_RM}='" + Combo6 + "'"
        
 
           .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        End With
End Sub




Private Sub Form_Load()

Call BukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "data_pemeriksaan"
Adodc1.Refresh



Combo1 = ""


RSBarang.Open "Select * from tb_pasien where No_RM like '%" & Combo6 & "%'", Koneksi
Combo6.Clear
Do While Not RSBarang.EOF
Combo6.AddItem RSBarang!No_RM
RSBarang.MoveNext
Loop
RSBarang.close

End Sub

